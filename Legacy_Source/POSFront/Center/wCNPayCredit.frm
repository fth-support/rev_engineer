VERSION 5.00
Object = "{D6810AA0-45D7-11CF-88F3-02608CA3BE43}#1.0#0"; "Msr.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.1#0"; "Codejock.SuiteCtrls.Unicode.v11.1.3.ocx"
Object = "{9D6F2CC0-07CE-11CF-A800-0020AF9EBB73}#1.0#0"; "Scanner.ocx"
Begin VB.Form wCNPayCredit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Credit Card;klg_AccNum"
   ClientHeight    =   4620
   ClientLeft      =   2400
   ClientTop       =   1410
   ClientWidth     =   7725
   ControlBox      =   0   'False
   Icon            =   "wCNPayCredit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   7725
   Tag             =   "2;"
   Begin VB.Frame ofaSwipeCard 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   2535
      Left            =   2900
      TabIndex        =   24
      Top             =   5400
      Width           =   3495
      Begin VB.Label olaForm 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Swipe or Insert Card into EDC"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   11
         Left            =   480
         TabIndex        =   25
         Top             =   960
         Width           =   2535
      End
   End
   Begin VB.Frame ofa3Sec 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   2535
      Left            =   120
      TabIndex        =   21
      Top             =   5400
      Width           =   3495
      Begin VB.Label olaForm 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Press ESC for manual or F1 for PRE-AUTH or wait for swipe card."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   10
         Left            =   360
         TabIndex        =   23
         Top             =   480
         Width           =   2775
      End
      Begin VB.Label olaForm 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "xx"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   9
         Left            =   1320
         TabIndex        =   22
         Top             =   1560
         Width           =   735
      End
   End
   Begin VB.Timer otmForm 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1200
      Top             =   3700
   End
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   4080
      Index           =   2
      Left            =   0
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   -240
      Width           =   7752
      _Version        =   720897
      _ExtentX        =   13674
      _ExtentY        =   7197
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   285
         Left            =   4920
         TabIndex        =   31
         Text            =   "gTAjLBBtYd1K03zzhSNn+uqsofhcoaYUwNnFh2QjHX4="
         Top             =   3240
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Decrypt"
         Enabled         =   0   'False
         Height          =   375
         Left            =   5400
         TabIndex        =   30
         Top             =   3600
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox otbRedeem 
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
         Left            =   5220
         MaxLength       =   12
         TabIndex        =   29
         Text            =   "0"
         Top             =   360
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Enabled         =   0   'False
         Height          =   375
         Left            =   5520
         TabIndex        =   27
         Top             =   2760
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox otbAmt 
         Height          =   375
         Left            =   4920
         TabIndex        =   18
         Text            =   "1"
         Top             =   2780
         Visible         =   0   'False
         Width           =   495
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
         TabIndex        =   2
         Top             =   1560
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
         Left            =   2400
         MaxLength       =   2
         TabIndex        =   1
         Top             =   1560
         Width           =   615
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
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   2160
         Width           =   3975
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
         TabIndex        =   4
         Top             =   3360
         Width           =   2415
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
         TabIndex        =   3
         Text            =   "0"
         Top             =   2760
         Width           =   2415
      End
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
         TabIndex        =   0
         Top             =   960
         Width           =   3975
      End
      Begin XtremeSuiteControls.ComboBox ocbType 
         Height          =   495
         Left            =   2400
         TabIndex        =   7
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
         Enabled         =   0   'False
         Style           =   2
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Type:;klg_Type"
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
         TabIndex        =   26
         Tag             =   "2;"
         Top             =   435
         Width           =   1950
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
         Left            =   4200
         TabIndex        =   17
         Tag             =   "2;"
         Top             =   1635
         Width           =   2775
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Card Type:;klg_CardType"
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
         TabIndex        =   16
         Tag             =   "2;"
         Top             =   2235
         Width           =   2130
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
         TabIndex        =   15
         Top             =   1560
         Width           =   210
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Approved Code :;klg_ApvCode"
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
         Width           =   2145
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount :;klg_Amt"
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
         TabIndex        =   13
         Tag             =   "2;"
         Top             =   2835
         Width           =   2160
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Account Number:;klg_AccNum"
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
         TabIndex        =   12
         Tag             =   "2;"
         Top             =   1035
         Width           =   3870
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
         Left            =   120
         TabIndex        =   11
         Tag             =   "2;"
         Top             =   1635
         Width           =   2130
      End
   End
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   1080
      Index           =   1
      Left            =   0
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   3564
      Width           =   7752
      _Version        =   720897
      _ExtentX        =   13674
      _ExtentY        =   1905
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
      Begin SCANNERLib.Scanner oScanner 
         Left            =   720
         Top             =   120
         _Version        =   65536
         _ExtentX        =   741
         _ExtentY        =   741
         _StockProps     =   0
      End
      Begin MSRLib.MSR OPOSMSR1 
         Left            =   0
         Top             =   135
         _Version        =   65536
         _ExtentX        =   741
         _ExtentY        =   741
         _StockProps     =   0
      End
      Begin VB.CommandButton ocmEDC 
         Caption         =   "EDC"
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
         Left            =   6240
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   400
         Width           =   1455
      End
      Begin VB.Timer otmEDC 
         Enabled         =   0   'False
         Interval        =   300
         Left            =   1680
         Top             =   120
      End
      Begin VB.CommandButton ocmCancel 
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
         Left            =   4680
         TabIndex        =   6
         TabStop         =   0   'False
         Tag             =   "2;"
         Top             =   400
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
         Left            =   3120
         TabIndex        =   5
         TabStop         =   0   'False
         Tag             =   "2;"
         Top             =   400
         Width           =   1455
      End
      Begin VB.Label olaForm 
         BackStyle       =   0  'Transparent
         Caption         =   "xx"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   8
         Left            =   2280
         TabIndex        =   20
         Top             =   360
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label olaForm 
         BackStyle       =   0  'Transparent
         Caption         =   "Cancel connect NET/EDC"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   19
         Top             =   600
         Visible         =   0   'False
         Width           =   2175
      End
   End
End
Attribute VB_Name = "wCNPayCredit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private oHyperCom As Object
Private oPosNet As Object
Private bW_FoundAcc As Boolean
Private bIsUnload As Boolean
Private bW_ChkCrd As Boolean
Private bW_KeyEnter As Boolean '***KT 51-11-20 ป้องกันการ Enter ซ้ำและมีการบันทึก Ontop มากกว่า 1 ครั้งใน CD
'NETS/EDC
Private bW_ManaulMode As Boolean
Private bW_HandleESC As Boolean
Private nW_TimeESC As Long
Private nW_TimeExpireEDC As Long
Private nW_Timer As Long
Private oEdc As Object

'NETS/EDC
'*Korn 55-10-10 Edit  เพิ่มตัวแปร 4 บรรทัด
Private bW_SendEcho As Boolean
Private bW_SendAmt As Boolean
Private bW_EdcCheck As Boolean
Private tW_EdcAction As String

'*Ao
Private tW_Comport As String, tW_EDCRun As String
Private nW_CaseERROR As Integer
Private bW_ChkESC As Boolean
'*NOT 58-04-16 EDC เพิ่มปุ่ม
Private bW_ChkocmEDC As Boolean
Private bW_EnableMSR  As Boolean '*Bump 58-04-18 CommSheet TK-ISS3000-558.02_22042015.xls
Private bW_C209  As Boolean '*Bump 58-04-30 กดปุ่ม EDC บนหน้าจอ
Public Function W_GETtMsgErrEDC(ByVal ptResCode As String) As String
'----------------------------------------------------------
' Call
' Cmt : *Eaw 56-10-24
'----------------------------------------------------------
Dim tMsg As String
 
On Error GoTo ErrHandle
    Select Case UCase(ptResCode)
        Case "00": tMsg = "APPROVED" '*Bump 57-07-07
        Case "01": tMsg = "REFER TO CARD ISSUER."
        Case "02": tMsg = "REFER TO CARD ISSUER’S SPECIAL CONDITION."
        Case "03": tMsg = "ERROR CALL HELP SN."
        Case "05": tMsg = "DO NOT HONOUR."
        Case "12": tMsg = "INVALID TRANSACTION (HELP TR)."
        Case "13": tMsg = "INVALID AMOUNT (HELP AM)."
        Case "14": tMsg = "INVALID CARD NUMBER (HELP RE)."
        Case "19": tMsg = "RE-ENTER TRANSACTION."
        Case "25": tMsg = "UNABLE TO LOCATE RECORD ON FILE (HELP NT)."
        Case "30": tMsg = "FORMAT ERROR (HELP FE)."
        Case "31": tMsg = "BANK NOT SUPPORTED BY SWITCH (HELP NS)."
        Case "41": tMsg = "LOST CARD."
        Case "43": tMsg = "STOLEN CARD PICK UP."
        Case "51": tMsg = "TRANSACTION DECLINED."
        Case "54": tMsg = "EXPIRED CARD."
        Case "55": tMsg = "INCORRECT PIN."
        Case "58": tMsg = "TRANSACTION NOT PERMITTED IN TERMINAL."
        Case "76": tMsg = "INVALID PRODUCT CODES (HELP DC)."
        Case "77": tMsg = "RECONCILE ERROR."
        Case "78": tMsg = "TRACE# NOT FOUND."
        Case "89": tMsg = "BAD TERMINAL ID."
        Case "91": tMsg = "ISSUER/SWITCH INOPERATIVE."
        Case "94": tMsg = "DUPLICATE TRANSMISSION."
        Case "96": tMsg = "SYSTEM MALFUNCTION."
        Case "SE": tMsg = "TERMINAL FULL."
        Case "PE": tMsg = "PIN ENTRY ERROR."
        Case "IC": tMsg = "INVALID CARD."
        Case "EC": tMsg = "CARD IS EXPIRED."
        Case "CE": tMsg = "CONNECTION ERROR."
        Case "RE": tMsg = "RECORD NOT FOUND."
        Case "HE": tMsg = "WRONG HOST NUMBER PROVIDED."
        Case "LE": tMsg = "LINE ERROR."
        Case "VB": tMsg = "TRANSACTION ALREADY VOIDED."
        Case "FE": tMsg = "FILE EMPTY / NO TRANSACTION TO VOID."
        Case "WC": tMsg = "CARD NUMBER DOES NOT MATCH."
'        Case "TA": tMsg = "TRANSACTION ABORTED BY USER."
        Case "TA": tMsg = "TIMEOUT ERROR/TRANSACTION ABORTED BY USER." '*Eaw 56-11-25 CommSheet TK-ISS3000-418-FAPL-149 00.00.01.xls
        Case "AE": tMsg = "AMOUNT DID NOT MATCH."
        Case "XX": tMsg = "TERMINAL NOT PROPERLY SETUP"
        Case "DL": tMsg = "LOGON NOT LOGON."
        Case "BT": tMsg = "BAD TLV COMMAND FORMAT."
        Case "IS": tMsg = "TRANSACTION NOT FOUND, INQUIRY SUCCESSFUL."
        Case "CD": tMsg = "CARD DECLINED TRANSACTION."
        Case "LH": tMsg = "LOYALTY HOST IS TEMPORARY OFFLINE."
        Case "IN": tMsg = "INVALID CARD (EZLINK CARD IS REFUNDED or BLACKLISTED)."
        Case "CO": tMsg = "CARD NOT READ PROPERLY."
        Case "TL": tMsg = "TOP UP LIMIT EXCEEDED."
        Case "PL": tMsg = "PAYMENT LIMIT EXCEEDED."
        Case "Y3": tMsg = "TRANSACTION FAILED."
        Case Else
            tMsg = "UNKNOWN RESPONSE CODE , TRANSACTION FAILED."
End Select
    W_GETtMsgErrEDC = tMsg
Exit Function
ErrHandle:
    W_GETtMsgErrEDC = "UNKNOWN RESPONSE CODE , TRANSACTION FAILED."
End Function

Private Function W_PRCbFindTCCardByNum(ByVal ptTCCardNo As String) As Boolean
'--------------------------------------------------------------------------------------------
'Call:
'Cmt: '*Eaw 56-09-04 ตรวจสอบ TC Card ว่ามีอยู่ใน Master หรือไม่
'----------------------------------------------------------------------------------------------
Dim tSql As String
Dim orsTemp As ADODB.Recordset
On Error GoTo ErrorHandles
    W_PRCbFindTCCardByNum = False
    tSql = W_PRCxStrSQL(ptTCCardNo)
    Call SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine, , , , , True, True) 'read master
    If Not (orsTemp.BOF And orsTemp.EOF) Then
        W_PRCbFindTCCardByNum = True
    Else
        W_PRCbFindTCCardByNum = False
    End If
Exit Function
ErrorHandles:
    W_PRCbFindTCCardByNum = False
End Function

Private Function W_CHKbFreeRangCard(ByVal ptCardType As String, ByRef ptRefCardType As String, ByVal ptTenderCode As String) As Boolean
'--------------------------------------------------------------------------------------------
'Call:
'Cmt: '*Eaw 56-09-03  CommSheet TK-ISS3000-330  ตรวจสอบ Free Rang Card
'----------------------------------------------------------------------------------------------
Dim tSql As String
Dim orsTemp As ADODB.Recordset
Dim tRefCode As String
On Error GoTo ErrorHandles
    W_CHKbFreeRangCard = False
'    '*Tao 60-06-08 ของเดิม
    '*Bmp 60-07-12  CommSheet TK-ISS3000-646 -Wrong Credit type updated for CUP transactions.xls        Compile ใน V.5.2.208
    '*เปิดใช้ อันเดิม เนื่องจาก Free Card Range จะกำหนดเป็น
    '*TCNMCreditCard.FTCdcType = 'C', หรือ 'D' หรือ X ที่เป็นประเภทบัตร Free  Card Range
    '* และ TCNMCreditCard.FTRemark = 'Y'
    '* โปรแกรม รับค่า ptCardType มาจาก Tag56 ที่รับมาจาก EDC
    tSql = "SELECT DISTINCT TOP 1 FTTdmCode,FTCdcCreditType,FTCdcDescCredit,FTCdcType, FTCdcStaBPSale,FTCdcStaBPReturn,FTRemark as FTFreeRangCard,FTCdcCreditGroup"
    tSql = tSql & vbCrLf & "FROM TCNMCreditCard WHERE ISNULL(FTRemark,'') = 'Y'"
    tSql = tSql & vbCrLf & "AND FTCdcType = '" & ptCardType & "'"
    tSql = tSql & vbCrLf & "AND FTTdmCode = '" & ptTenderCode & "'"
    tSql = tSql & vbCrLf & "AND FTCdcType IN(SELECT DISTINCT FTCdcType FROM TCNMCreditCard WHERE ISNULL(FTRemark,'N') = 'Y')" '*Bmp 60-07-18 เช็ค Free Range Card Compile ใน V.5.2.196, V.5.2.209 (02/08/2017)

    '*Tao 60-06-08 ให้เช็ค CardRange เพิ่มด้วย (ป้องกันได้ข้อมูลจากที่มี CardType และ FreeRange เหมือนกัน)                  Compile ใน V.5.2.207
'    tSql = "SELECT TOP 1 CC.FTCdcCreditType,CC.FTCdcDescCredit,CC.FTCdcCreditGroup,CC.FTCdcStaBPSale,CC.FTCdcStaBPReturn"
'    tSql = tSql & vbCrLf & "FROM TCNMCreditCard CC WITH(NOLOCK)"
'    tSql = tSql & vbCrLf & "INNER JOIN TCNMCreditCardRange CR WITH(NOLOCK) ON CC.FTTdmCode = CR.FTTdmCode AND CC.FTCdcCreditType = CR.FTCdcCreditType"
'    tSql = tSql & vbCrLf & "WHERE ISNULL(CC.FTRemark,'') = 'Y'"
'    tSql = tSql & vbCrLf & "AND CC.FTCdcType = '" & ptCardType & "'"
'    tSql = tSql & vbCrLf & "AND CC.FTTdmCode = '" & ptTenderCode & "'"
'    tSql = tSql & vbCrLf & "AND (LEFT('" & Trim(otbAcc.Text) & "' ,CC.FNCdcDigitChk)"
'    tSql = tSql & vbCrLf & "BETWEEN LEFT(CR.FTCcrRangeFrom,CC.FNCdcDigitChk) AND LEFT(CR.FTCcrRangeTo,CC.FNCdcDigitChk))"
'    tSql = tSql & vbCrLf & "ORDER BY LEN(CR.FTCcrRangeFrom) DESC,CR.FTCcrRangeTo DESC,CC.FNCdcDigitChk DESC"
    '=============================================================================
    
    Call SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine, , , , , True, True) 'read master
    If Not (orsTemp.BOF And orsTemp.EOF) Then
        ptRefCardType = SP_FEDtChkString(orsTemp, "FTCdcCreditType")
        tVB_TenderName = SP_FEDtChkString(orsTemp, "FTCdcCreditType") '*Eaw 56-09-13 ให้เปลี่ยนเป็น FTCdcCreditType แทน FTCdcDescCredit
        tVB_CreditGroup = SP_FEDtChkString(orsTemp, "FTCdcCreditGroup")         '*Ao 2013/11/05
        tVB_CrdStaBPSale = SP_FEDtChkString(orsTemp, "FTCdcStaBPSale")    '*Eaw 56-11-06
        tVB_CrdStaBPReturn = SP_FEDtChkString(orsTemp, "FTCdcStaBPReturn")
        
        W_CHKbFreeRangCard = True
        Call SP_CRTxLogMallCardEvent("Credit", "EDCResponse", "W_CHKbFreeRangCard=True=", tSql)
    Else
        W_CHKbFreeRangCard = False
        ptRefCardType = ""
        Call SP_CRTxLogMallCardEvent("Credit", "EDCResponse", "W_CHKbFreeRangCard=False=", tSql)
    End If
Exit Function
ErrorHandles:
    W_CHKbFreeRangCard = False
    ptRefCardType = ""
    Call SP_CRTxLogMallCardEvent("Credit", "EDCResponse", "W_CHKbFreeRangCard=ErrorHandles=", tSql)
End Function

Private Function W_CHKbCupsCard(ByVal ptType As String, ByRef ptRefCardType As String) As Boolean
'--------------------------------------------------------------------------------------------
'Call:
'Cmt: หา Card Type ของ CUP กรณีค่าจาก EDC เมื่อส่ง tag55 มาเป็น 5501C     '*Eaw 56-07-08 CommSheet TK-ISS3000-248-TR.xls
'----------------------------------------------------------------------------------------------
Dim tSql As String
Dim orsTemp As ADODB.Recordset
Dim tRefCode As String
On Error GoTo ErrorHandles
    W_CHKbCupsCard = False
    tSql = "SELECT TOP 1 FTCdcCreditType FROM TCNMCreditCard"
    tSql = tSql & " WHERE FTCdcType = '" & ptType & "'"
    tSql = tSql & " AND LEN('" & Trim(otbAcc.Text) & "') <= FNCdcDigits"
    tSql = tSql & vbCrLf & "AND FTTdmCode = '" & tVB_TenderCode & "'" '*Eaw 56-10-28
    Call SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine, , , , , True, True) 'read master
    If Not (orsTemp.BOF And orsTemp.EOF) Then
        ptRefCardType = SP_FEDtChkString(orsTemp, "FTCdcCreditType")
        W_CHKbCupsCard = True
    Else
        W_CHKbCupsCard = False
        ptRefCardType = ""
    End If
Exit Function
ErrorHandles:
    W_CHKbCupsCard = False
    ptRefCardType = ""
End Function

Private Sub W_PRCxloadEDC()
'_______________________________________________
'*NOT 58-04-16 เพิ่มปุม EDC
'
'
'_______________________________________________
If tVB_TypeOfEDC = 4 Then
        Dim oEdc As Object
        Dim tTime As String
        olaForm(8).Caption = ""
        olaForm(9).Caption = ""
        otbAcc.Locked = True
        otbAcc.BackColor = nVB_CNClrBLock
        nW_Timer = 0
        bW_HandleESC = True
        bW_ManaulMode = False
        ocmCancel.Cancel = False
        bW_ChkocmEDC = True
        Call SP_OBJxSetFocus(otbAcc)

        tTime = tVB_EDCWait
        If tTime = "" Then
            nW_TimeESC = CLng(tVB_EDCWait)
            Call SaveSetting("AdaPosMall", "POSFront", "TimeEDCEsc", tVB_EDCWait)
        Else
            If CLng(tTime) > 0 Then
                nW_TimeESC = CLng(tTime)
            Else
                nW_TimeESC = CLng(tVB_EDCWait)
                Call SaveSetting("AdaPosMall", "POSFront", "TimeEDCEsc", tTime)
            End If
        End If
        tTime = tVB_EDCTimeOut
        If tTime = "" Then
                nW_TimeESC = CLng(tVB_EDCTimeOut)
                Call SaveSetting("AdaPosMall", "POSFront", "TimeEDCEsc", tVB_EDCTimeOut)
        Else
            If CLng(tTime) > 0 Then
                nW_TimeExpireEDC = CLng(tTime)
            Else
                nW_TimeESC = CLng(tVB_EDCTimeOut)
                Call SaveSetting("AdaPosMall", "POSFront", "TimeEDCEsc", tVB_EDCTimeOut)
            End If
        End If
        '****
        If bVB_CNEdcPos Then
            Select Case Val(Format(tVB_TransactionSub))
                    Case tEN_TrnReturnManualTax, tEN_TrnReturnTaxInc, tEN_TrnSaleManualTax, tEN_TrnSaleTaxInc
                        otbAcc.Locked = False
                        otbAcc.BackColor = nVB_CNClrBUnLock
                        bW_HandleESC = False
                        bW_ManaulMode = True
                        otmEDC.Enabled = False
                        otmForm.Enabled = False
                        Call SP_OBJxSetFocus(otbAcc)
                        ocmEDC.Enabled = False
                    Case Else
                        olaForm(8).Caption = nW_TimeESC
                        olaForm(9).Caption = nW_TimeESC
                        ofa3Sec.Left = 1600     'center
                        ofa3Sec.Top = 600
                        ofaSwipeCard.Left = 2900    'ซ่อน
                        ofaSwipeCard.Top = 5400
                        otcForm1(1).Enabled = False     'lock text box
                        otcForm1(2).Enabled = False     'lock ปุ่ม
                        otmForm.Enabled = True
'                        Call otmForm_Timer
                End Select
        End If
    End If
End Sub

Private Function W_PRCxStrCupsCard(ByVal ptAccNum As String) As String
'---------------------------------------------------------------------------
'Call:
'Cmt: String ตรวจสอบว่าเป็น Free Rang Card หรือไม่  '*Eaw 56-05-15 CommSheet TK-ISS3000-159-IT2 00.00.03.xls
'----------------------------------------------------------------------------
    Dim tSql As String
    tSql = "SELECT TOP 1 * FROM TCNMCreditCard"
    tSql = tSql & " WHERE FTCdcCreditType = '" & Trim(ocbType.Text) & "'"
    tSql = tSql & " AND LEN('" & ptAccNum & "') <= FNCdcDigits"
    tSql = tSql & vbCrLf & "AND FTTdmCode = '" & tVB_TenderCode & "'" '*Eaw 56-10-28
    W_PRCxStrCupsCard = tSql
End Function

Private Function W_CHKbAccByID(ByRef ptRefCardType As String) As Boolean
'-------------------------------------------------------------------------------------------------
'Call:
'Cmt: หา Card Type ที่อยู่ใน master เมื่อรูดบัตรผ่าน EDC     '*Eaw 56-05-13 CommSheet TK-ISS3000-167-ST 00.00.02.xls
'-------------------------------------------------------------------------------------------------
Dim tSql As String
Dim orsTemp As ADODB.Recordset
On Error GoTo ErrorHandles
    W_CHKbAccByID = False
    tSql = W_PRCxStrSQL(Trim(otbAcc.Text))
    Call SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine, , , , , True, True) 'read master
    If Not (orsTemp.BOF Or orsTemp.EOF) Then
        W_CHKbAccByID = True
        ptRefCardType = SP_FEDtChkString(orsTemp, "FTCdcCreditType")  '*Eaw 56-06-13  CommSheet TK-ISS3000-086 00.00.08.xls
        tVB_CrdCardType = SP_FEDtChkString(orsTemp, "FTCdcType") '*Eaw 56-09-10
        tVB_TenderName = SP_FEDtChkString(orsTemp, "FTCdcCreditType") '*Eaw 56-09-13 ให้เปลี่ยนเป็น FTCdcCreditType แทน FTCdcDescCredit
        tVB_CreditGroup = SP_FEDtChkString(orsTemp, "FTCdcCreditGroup")         '*Ao 2013/11/05
        tVB_CrdStaBPSale = SP_FEDtChkString(orsTemp, "FTCdcStaBPSale") '*Eaw 56-11-06
        tVB_CrdStaBPReturn = SP_FEDtChkString(orsTemp, "FTCdcStaBPReturn")
        Call SP_CRTxLogMallCardEvent("Credit", "EDCResponse", "W_CHKbAccByID=True=", tSql)
    Else
        W_CHKbAccByID = False
        ptRefCardType = ""   '*Eaw 56-06-13  CommSheet TK-ISS3000-086 00.00.08.xls
        Call SP_CRTxLogMallCardEvent("Credit", "EDCResponse", "W_CHKbAccByID=False=", tSql)
    End If
    Exit Function
ErrorHandles:
    W_CHKbAccByID = False
    ptRefCardType = ""  '*Eaw 56-06-13  CommSheet TK-ISS3000-086 00.00.08.xls
    Call SP_CRTxLogMallCardEvent("Credit", "EDCResponse", "W_CHKbAccByID=ErrorHandles=", tSql)
End Function

'Private oHyperCom As AdaEDC.cHyperCom 'อย่าลบ[EDC]
'Private oPosNet As AdaEDC.cPosNet 'อย่าลบ[EDC]
Private Function W_PRCxStrSQL(ByVal ptAccNum$) As String
Dim tSql As String
Dim nLen As Long '*Eaw 56-07-10 CommSheet TK-ISS3000-272-TR.xls
    nLen = Len(ptAccNum) '*Eaw 56-07-10
    tSql = "SELECT TOP 1 TCNMCreditCard.*"
    tSql = tSql & " ,TCNMCreditCardRange.FTCcrRangeFrom,TCNMCreditCardRange.FTCcrRangeTo"
    tSql = tSql & " FROM TCNMCreditCard INNER JOIN"
    tSql = tSql & " TCNMCreditCardRange ON TCNMCreditCard.FTTdmCode = TCNMCreditCardRange.FTTdmCode AND"
    tSql = tSql & " TCNMCreditCard.FTCdcCreditType = TCNMCreditCardRange.FTCdcCreditType"
    tSql = tSql & " WHERE TCNMCreditCard.FTTdmCode = '" & tVB_TenderCode & "'"
'    If W_CHKbCupsCard(Trim(ocbType.Text)) Then    '*Eaw 56-05-07 CommSheet TK-ISS3000-159-IT2 00.00.02.xls
'        tSql = tSql & " AND TCNMCreditCard.FTCdcCreditType = '" & Trim(ocbType.Text) & "'"
'    ElseIf ocbType.Text <> "-" Then
    If tVB_TenderCode <> "T017" Then
        tSql = tSql & " AND ISNULL(TCNMCreditCard.FTRemark,'N') = 'N'"  '*Eaw 56-05-15 CommSheet TK-ISS3000-159-IT2 00.00.03.xls
    End If
'    tSql = tSql & " AND (LEFT('" & ptAccNum & "' ,TCNMCreditCard.FNCdcDigitChk) BETWEEN LEFT(TCNMCreditCardRange.FTCcrRangeFrom,TCNMCreditCard.FNCdcDigitChk)  AND LEFT(TCNMCreditCardRange.FTCcrRangeTo,TCNMCreditCard.FNCdcDigitChk)  "   '*Eaw 56-06-11
'    tSql = tSql & " OR (LEN('" & ptAccNum & "') = 16 AND '" & ptAccNum & "' BETWEEN RIGHT(TCNMCreditCardRange.FTCcrRangeFrom,16) AND RIGHT(TCNMCreditCardRange.FTCcrRangeTo,16)))"   '*Eaw 56-06-11

'    tSql = tSql & " AND LEN('" & ptAccNum & "') = TCNMCreditCard.FNCdcDigits"          '*Ao 2013/05/28  TC Card
    Select Case (tVB_TenderCode)   '*Eaw 56-06-11 แยกการหา Credit Card กับ TC Card
        Case "T017"
'            tSql = tSql & " AND (LEFT('" & ptAccNum & "' ,TCNMCreditCard.FNCdcDigitChk) BETWEEN LEFT(TCNMCreditCardRange.FTCcrRangeFrom,TCNMCreditCard.FNCdcDigitChk)  AND LEFT(TCNMCreditCardRange.FTCcrRangeTo,TCNMCreditCard.FNCdcDigitChk)  "
'            tSql = tSql & " OR (LEN('" & ptAccNum & "') = 16 AND '" & ptAccNum & "' BETWEEN RIGHT(TCNMCreditCardRange.FTCcrRangeFrom,16) AND RIGHT(TCNMCreditCardRange.FTCcrRangeTo,16)))"
            '*Eaw 56-07-10 CommSheet TK-ISS3000-272-TR.xls
            tSql = tSql & vbCrLf & " AND (LEFT( '" & ptAccNum & "' ,TCNMCreditCard.FNCdcDigitChk) BETWEEN LEFT(TCNMCreditCardRange.FTCcrRangeFrom,TCNMCreditCard.FNCdcDigitChk)  AND LEFT(TCNMCreditCardRange.FTCcrRangeTo,TCNMCreditCard.FNCdcDigitChk)"
            tSql = tSql & vbCrLf & " OR (LEN(  '" & ptAccNum & "') = " & nLen & " AND '" & ptAccNum & "' BETWEEN RIGHT(TCNMCreditCardRange.FTCcrRangeFrom, " & nLen & ") AND RIGHT(TCNMCreditCardRange.FTCcrRangeTo, " & nLen & ")))"
            tSql = tSql & vbCrLf & " AND (LEN(RIGHT(RTRIM(TCNMCreditCardRange.FTCcrRangeFrom), " & nLen & ")) = " & nLen & ")"
        Case "T002", "T003"
            tSql = tSql & " AND (LEFT('" & ptAccNum & "' ,TCNMCreditCard.FNCdcDigitChk) BETWEEN LEFT(TCNMCreditCardRange.FTCcrRangeFrom,TCNMCreditCard.FNCdcDigitChk)  AND LEFT(TCNMCreditCardRange.FTCcrRangeTo,TCNMCreditCard.FNCdcDigitChk) ) "
            tSql = tSql & " AND LEN('" & ptAccNum & "') = TCNMCreditCard.FNCdcDigits"
    End Select

    tSql = tSql & " ORDER BY TCNMCreditCard.FNCdcDigitChk DESC" '*Tao 60-06-08 ของเดิม
    'tSql = tSql & " ORDER BY LEN(TCNMCreditCardRange.FTCcrRangeFrom) DESC,TCNMCreditCardRange.FTCcrRangeTo DESC,TCNMCreditCard.FNCdcDigitChk DESC" '*Tao 60-06-08 ให้เอาข้อมูลที่ใกล้เคียงความเป็นจริงที่สุด
    W_PRCxStrSQL = tSql

End Function

Private Function W_PRCbFindCrdByMEM(ByVal ptAccNum As String) As Boolean
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
        tVB_CreditGroup = SP_FEDtChkString(orsTemp, "FTCdcCreditGroup")         '*Ao 2013/11/05
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

Private Sub W_CHKxCaseERROR(ByVal pnCase As Integer)
    Select Case pnCase
        Case 1
            Call SP_MSGnShowing("tms_CN008038", nCS_Warning)
        Case 2
            Call SP_MSGnShowing("tms_CN008039", nCS_Warning)      '*Ao 2013/03/15
        Case 3
            'Call SP_MSGnShowing("tms_CN008039", nCS_Warning)      '*Ao 2013/03/15
            MsgBox "Invalid Card Data from EDC", tCS_CNMsgWarn, tVB_CNAppTitle
        Case Else
            Call SP_MSGnShowing("tms_CN005065", nCS_Warning)
    End Select
End Sub

Private Sub W_EDCxSplitResponse(ByVal ptRet$)
'____________________________________________________________________________________
'   Cmt : split response to array
'   Create by : *Korn 55-10-12 Edit เพิ่ม procedure ใหม่
'รูดไม่ผ่าน
        'แคชเชียร ์กด end ที่เครื่องเลย ยังไม่ได้รูด    EDC คืนค่า=> R20004120000000001003902TA5706000001
        'แคชเชียร ์รูดบัตรไปครึ่งทาง แล้วกด end     EDC คืนค่า=> R20004120000000001003902XX5706000001
        'แคชเชียร์ ไม่กดอะไร รอจนเครื่อง end          EDC คืนค่า=> R20004120000000001003902TA5706000001
        'แคชเชียร์ รูดบัตรผิดประเภท, บัตรอื่น            EDC คืนค่า=> R20004120000000001003902IC5706000001
    
        'aTag[0,0]=>R200-
        'aTag [0,1] - [1,1] => 04 - 000000000100
        'aTag [0,2] - [1,2] => 39 - TA
        'aTag [0,3] - [1,3] => 57 - 1

'รูดผ่าน
        'Response:    R2000216XXXXXXXXXXXX93200412000000000100071010151522551404XXXX2201M3712DEMODEMODEMO3806DEMO5139020041081234567842151681680000000005203DBS5410MASTERCARD5501M5601D57060000016206000047640600000296101100000004
                            '   R2000216XXXXXXXXXXXX03230412000000002000071011211248121404XXXX2201C3712542624604920380604063539020041084001300842121681689971825204DBSI53620440008000E800A000000004101001DBS MasterCard  66D7C0FC562A19745410MASTERCARD5501M5601L5706000005620600001363600000130000000020000000000020000000000000000001000000000000006406000001961011000000049826 /SWITCH UAT 851
        'aTag[0,0]=>R200-                                           1.6.2 ประเภท Command & Response R200:Sale, R201:Pre-Auth, R203:Refund, R902:Echo ...
        'aTag[0/1,1]=>02-XXXXXXXXXXXX9320    *หมายเลขบัตร (ได้ค่า XXXX XXXX XXXX NNNN ตลอด)
        'aTag[0/1,2]=>04-000000000100                 ยอดเงิน 1.00
        'aTag[0/1,3]=>07-1015152255                     เดือนวันเวลา MMDDHHmmss ส่วนปีหามาเพิ่มเอง เป็นปีปัจจุบัน
        'aTag[0/1,4]=>14-XXXX                                  *Expired YYMM (ได้ค่า XXXX ตลอด)
        'aTag[0/1,5]=>22-M                                         1.6.11 ประเภทการรูด E:คีย์เอง, M:แถบแม่เหล็ก, C:ชิป, ...
        'aTag[0/1,6]=>37-DEMODEMODEMO     หมายเลขอ้างอิง
        'aTag[0/1,7]=>38-DEMO51                           *Approved Code
        'aTag[0/1,8]=>39-00                                        1.6.3 Response Code, 00: ไม่ต้องทำอะไร ...
        'aTag[0/1,9]=>41-12345678                           หมายเลขเครื่อง EDC
        'aTag[0/1,10]=>42-168168000000000        หมายเลขร้านค้าที่จดกับ EDC
        'aTag[0/1,11]=>52-DBS                                  Bank ของ EDC
        'aTag[0/1,12]=>54-MASTERCARD            *ชื่อประเภทบัตร VISA, MASTERCARD, AMEX
        'aTag[0/1,13]=>55-M                                        *1.6.6 ตัวย่อประเภทบัตร V:VISA, M:MASTER, E:EZ-LINK, P:PIN DEBIT ...
        'aTag[0/1,14]=>56-D                                        1.6.7 ประเภทของ Bank(Host) D:Normal, A:American Express, N:Diners ...
        'aTag[0/1,15]=>57-000001                             1.6.6 หมายเลขคำสั่ง เลข 6 หลักเริ่มที่ 000001 ...
        'aTag[0/1,16]=>62-000047                             หมายเลข Invoice
        'aTag[0/1,17]=>63-000013000000002000000000002000000000000000000100000000000000    Point Info
        'aTag[0/1,18]=>64-000002                             หมายเลข Batch
        'aTag[0/1,19]=>96-1100000004                    หมายเลขอ้างอิงในการพิมพ์
        
        '*Tao 62-07-19                                                                                                                                                                         '-LMS-' '--- Amount ---' '--- Redeem --'
        'aTag[0/1,99]=>63-Transaction Info            Additional transaction information (Redeem Tender) (63 60 000023 000000010000 000000001000 000000000000 000100 000000006000)
        'aTag[0/1,99]=>98-Card Holder Name      Card Holder's Name (98 26 WONG SOONG TECK)
'____________________________________________________________________________________
    Dim tTemp$, tCut$, nLen%, n%, tMsg$, tStrCard$
    Dim aTag() As String
    Dim tExpired As String
    Dim tCardType As String
    Dim bChk  As Boolean
    Dim tRefCardType As String  '*Eaw 56-06-13 เก็บค่า Card Type ที่เอามาจาก master เช่น TKAmex,CUP,Visa
    Dim tChkCardType As String '*Eaw 56-07-08 เก็บค่ากรณีเป็น ezLink เพื่อไม่เช็ควันหมดอายุบัตร
    Dim tTagDetail As String '*Eaw 56-07-24
    Dim tCardTag55 As String
    Dim tTag39 As String
    
    '*Tao 62-07-19 Redeem Tender
    Dim tHolderName As String
    Dim cAmount As Double
    Dim cRedeem As Double
    Dim cBalance As Double
    Dim tTmp As String
    Dim tTmpCard As String
    Dim tTmpDate As String
    Dim nLine%
    On Error GoTo ErrorHandles
    
    tTemp = ptRet
'    tVB_EDCDetail = ptRet
    'ตัดรอบแรกก่อน 4 ตัว(fixed)
    ReDim Preserve aTag(1, n)
    tCut = Left(tTemp, 4)
    aTag(0, n) = tCut
    tTemp = Right(tTemp, Len(tTemp) - 4)    'ตัด temp ออก
    
    'เริ่มวนตัดที่เหลือ
    Do While tTemp <> ""
        n = n + 1
        nLine = n '*Thirapong 64-12-10
        'ปรับ size array
        ReDim Preserve aTag(1, n)

        'เก็บค่า tag ลง aTag(0, n)
        tCut = Left(tTemp, 2)   'tag code 2 ตัว(fixed)
        tTag39 = tCut
        tTemp = Right(tTemp, Len(tTemp) - 2)    'ตัด temp ออก
        aTag(0, n) = tCut
        
        'เก็บค่า value ลง aTag(1, n)
        tCut = Left(tTemp, 2)   'tag length 2 ตัว(fixed)
        tTemp = Right(tTemp, Len(tTemp) - 2)   'ตัด temp ออก
        nLen = Int(tCut)        'ความยาวของ value
        tCut = Left(tTemp, nLen)    'tag value nLen ตัว
        If Len(tTemp) < nLen Then
            tTemp = ""  'ตัด temp ออก
        Else
             tTemp = Right(tTemp, Len(tTemp) - nLen)   'ตัด temp ออก
        End If
        aTag(1, n) = tCut
        '*Tao 62-12-10 ปิดไว้ เนื่องจากได้ข้อมูลครบ (ไม่ได้ Trim ptRet จะไม่มี error5)
'        If bW_C209 = True Then '*Bump 58-05-11 กรณีกดปุ่ม EDC ให้เช็คถึง Tag39 ก็พอ ป้องกัน error5
'            If tTag39 = "39" Then Exit Do
'        End If
'        Debug.Print "aTag[0/1," & n & "]=>" & aTag(0, n) & "-" & aTag(1, n)
'        tVB_EDCDetail = tVB_EDCDetail & aTag(0, n) & "-" & aTag(1, n)
    Loop
    
    nLine = n '*Thirapong 64-12-10
    Call SP_CRTxLogMallCardEvent("Credit", "EDCResponse", "SplitResponseLast n=", nLine)  '*Thirapong 64-12-10
    nLine = 1
    '*BG 57-01-15 CommSheet TK-ISS3000-449-FAPL-171.xls
    Call SP_CRTxLogMallCardEvent("Credit", "EDCResponse", "SplitResponseLastTagID=", n & ":" & tCut)
    Call SP_CRTxLogMallCardEvent("Credit", "EDCResponse", "SplitResponseLastText=", tTemp)
    nLine = 2
    'เช็คว่ามีค่ารหัสบัตรหรือไม่ (02:มีรหัสบัตร)
    If aTag(0, 1) <> "02" Then      'ถ้าไม่มีรหัสบัตร แสดงว่าอาจมีปัญหา
        If aTag(0, 1) = "39" Then '*Bump 58-04-30
            tMsg = SP_GETtMsgErrEDC(tCut) '
            Call SP_MSGnShowing("tms_CN005065" & ";" & tMsg, nCS_Error)
            Call SP_EDCxINSERTLog(tMsg)
            Unload Me
            Exit Sub
        Else
            If aTag(0, 2) = "39" Then    'error แบบ track ได้   ส่วนใหญ่รหัส 39
    '            Select Case aTag(1, 2)
    '                Case "XX": tMsg = "Terminal not properly setup."
    '                Case "IC": tMsg = "Invalid Card"
    '                Case "TA": tMsg = "Transaction aborted by user."
    '                '*Eaw 56-08-15 CommSheet TK-ISS3000-228-ST 00 00 02 FSBT NG.xls เพิ่ม case EC กับ case else
    '                Case "EC": tMsg = "Card is Expired, Please use a different card"
    '                Case Else
    '                    tMsg = "Invalid Card"
    '            End Select
                tMsg = SP_GETtMsgErrEDC(aTag(1, 2)) '*Eaw 56-10-24 เปลี่ยนไปเรียก Respone Code ใน Function แทน
                Call SP_EDCxINSERTLog(tMsg)
            Else
                'all error, แบบ track ไม่ได้
                tMsg = "Invalid Account Number !!"
                Call SP_EDCxINSERTLog(tMsg)
            End If
            
            Call SP_MSGnShowing("tms_CN005065" & ";" & tMsg, nCS_Error)
            Unload Me
            Exit Sub
        End If
    End If
    nLine = 3
    'แสดงว่ารูดบัตร ผ่าน
    'ซ่อน Swipe card แต่ไม่ปลด lock ปุ่ม ok, cancel
    ofaSwipeCard.Left = 2900        'ซ่อน
    ofaSwipeCard.Top = 5400
    
    'นำค่าลงในหน้าจอ
'    otbAcc.Text = aTag(1, 1)                                                     'รหัสบัตร
'    orbMonth.Text = Right(aTag(1, 4), 2)                                'expired Month เอา 2 ตัวหลัง
'    orbYear.Text = Left(aTag(1, 4), 2)                                      'expired Year เอา 2 ตัวหน้า
'    otbCardType.Text = aTag(1, 12)                                        '*ชื่อ EDC Type ?? หรืออาจจะเอา Tag(55) หรือ aTag(0,13) ไปค้นหาในตารางเอาเอง
''    otbNum.Text = SP_FMTtNum2Shw(cVB_Balance)     'ยอดเงิน ใช้ตัวแปรได้เลย, ไม่ต้องใช้จากบัตรเครดิต
'    otbNum.Text = SP_FMTtNum2Shw(cVB_PayCredit)
'    otbApvCode.Text = aTag(1, 7)                                            'approved code
'    ocbType.Text = aTag(1, 13)
'
    '*TU 2020-09-09 CR_Encrypted Creditcard
    '---------------------------------------------------------------------------------------------------------
    Call SP_CRTxLogMallCardEvent("Credit", "EDCResponse", "Tag02=", SP_EDCxCheckTag(aTag, n, "02")) '*Thirapong 64-12-10    Log monitor CommSheet TTK-ISS3000-794 TK-ISSUE - Error 5 from EDC Response
    tTmpCard = ""
    If Left(SP_EDCxCheckTag(aTag, n, "02"), 1) = "#" Then
        nLine = 4
        tTmpCard = SP_EDCtDecrypt(Trim(SP_EDCxCheckTag(aTag, n, "02")))
        nLine = 5
        Call SP_CRTxLogMallCardEvent("Credit", "EDCResponse", "DecryptTag02=", tTmpCard) '*Thirapong 64-12-10    Log monitor CommSheet TTK-ISS3000-794 TK-ISSUE - Error 5 from EDC Response
        If tTmpCard <> "" Then
            nLine = 6
            otbAcc.Text = Right(tTmpCard, Len(tTmpCard) - 2)
            tStrCard = Right(tTmpCard, Len(tTmpCard) - 2)
            nLine = 7
        Else
            nLine = 8
            otbAcc.Text = ""
        End If
    Else
        nLine = 9
        otbAcc.Text = SP_EDCxCheckTag(aTag, n, "02")
        tStrCard = SP_EDCxCheckTag(aTag, n, "02")
    End If
    
    tTmpDate = ""
    Call SP_CRTxLogMallCardEvent("Credit", "EDCResponse", "Tag14=", SP_EDCxCheckTag(aTag, n, "14")) '*Thirapong 64-12-10    Log monitor CommSheet TTK-ISS3000-794 TK-ISSUE - Error 5 from EDC Response
    
    If Left(SP_EDCxCheckTag(aTag, n, "14"), 1) = "#" Then
        nLine = 10
        tTmpDate = SP_EDCtDecrypt(Trim(SP_EDCxCheckTag(aTag, n, "14")))
        nLine = 11
        tExpired = Right(tTmpDate, Len(tTmpDate) - 2)
        nLine = 12
        Call SP_CRTxLogMallCardEvent("Credit", "EDCResponse", "DecryptTag14=", tTmpDate)  '*Thirapong 64-12-10    Log monitor CommSheet TTK-ISS3000-794 TK-ISSUE - Error 5 from EDC Response
    Else
        nLine = 13
        tExpired = SP_EDCxCheckTag(aTag, n, "14")
    End If
    'otbAcc.text = SP_EDCxCheckTag(aTag, n, "02")
    'tExpired = SP_EDCxCheckTag(aTag, n, "14")
    '---------------------------------------------------------------------------------------------------------
    orbMonth.Text = Right(tExpired, 2)
    orbYear.Text = Left(tExpired, 2)
    otbCardType.Text = SP_EDCxCheckTag(aTag, n, "54")

    nLine = 14

    tVB_CrdType = Trim(otbCardType.Text)
    
    otbRedeem.Text = 0 '*Tao 62-07-19 Clear Redeem
    
    'If Trim(otbNum.Text) = "" Then otbNum.Text = SP_FMTcCurPoint(cVB_PayCredit, True, True)   '*Bump 58-05-11
    otbNum.Text = SP_FMTcCurPoint(cVB_PayCredit, True, True)  '*KT 66-12-06 CR Bank Instant Point Redemption    read payment amt from tag 04
    'tTmp = SP_EDCxCheckTag(aTag, n, "04") '*KT 67-01-29 Downgrade CR Bank Instant Point Redemption
    'otbNum.Text = SP_FMTcCurPoint(Val(Format(tTmp)) / 100, True, True)
    otbApvCode.Text = SP_EDCxCheckTag(aTag, n, "38")
    nLine = 15
'*Eaw 56-03-27  CommSheet TK-ISS3000-086 00.00.03.xls *Eaw 56-09-05  CommSheet TK-ISS3000-330 ให้แยกว่าถ้าเป็น Free rang ที่ OcbType ให้แสดง '---' จึงย้ายลงไปแยกเป็นเคสตามประเภทบัตร
'    ocbType.AddItem SP_EDCxCheckTag(aTag, n, "55")
'    ocbType.Text = SP_EDCxCheckTag(aTag, n, "55")
'    ocbType.AddItem SP_EDCxCheckTag(aTag, n, "56"), 1
'    ocbType.Text = SP_EDCxCheckTag(aTag, n, "56")
    '*Eaw 56-07-24
    'tTagDetail = "02" & SP_EDCxCheckTag(aTag, n, "02") '*Tao 61-09-18 ของเดิม
    
    '*Thirapong 64-12-10    Log monitor CommSheet TTK-ISS3000-794 TK-ISSUE - Error 5 from EDC Response
    
    'tTagDetail = "02" & SP_PRCtStrMarkX(SP_EDCxCheckTag(aTag, n, "02")) '*Tao 61-09-18 Mark X ก่อน
    tTagDetail = "02" & tStrCard 'SP_PRCtStrMarkX(SP_EDCxCheckTag(aTag, n, "02")) '*Tao 61-09-18 Mark X ก่อน
    tTagDetail = tTagDetail & ",14" & tExpired 'SP_EDCxCheckTag(aTag, n, "14")
    nLine = 16
    
    If bW_C209 = True Then '*Bump 58-04-30
            '02           card Number
            '14  Expire Date
            '22  Entry Mode
            '39  Response code
            '98  card 's Holder name
            '*Tao 62-07-22 ของเดิม
            'tTagDetail = tTagDetail & ",14" & SP_EDCxCheckTag(aTag, n, "22")
            'tTagDetail = tTagDetail & ",14" & SP_EDCxCheckTag(aTag, n, "39")
            'tTagDetail = tTagDetail & ",14" & SP_EDCxCheckTag(aTag, n, "98")
            
            '*Tao 62-07-22 ผิดรึป่าว...
            tTagDetail = tTagDetail & ",23" & SP_EDCxCheckTag(aTag, n, "22")
            tTagDetail = tTagDetail & ",39" & SP_EDCxCheckTag(aTag, n, "39")
            tTagDetail = tTagDetail & ",98" & SP_EDCxCheckTag(aTag, n, "98")
            nLine = 17
    Else
        tTagDetail = tTagDetail & ",38" & SP_EDCxCheckTag(aTag, n, "38")
        tTagDetail = tTagDetail & ",56" & SP_EDCxCheckTag(aTag, n, "56")
        tTagDetail = tTagDetail & ",54" & SP_EDCxCheckTag(aTag, n, "54")
        tTagDetail = tTagDetail & ",55" & SP_EDCxCheckTag(aTag, n, "55")
        tTagDetail = tTagDetail & ",98" & SP_EDCxCheckTag(aTag, n, "98") '*Tao 62-07-22 ฝากเอาไปพิมพ์ (Tag98 = Card 's Holder Name)
        'tTagDetail = tTagDetail & ",98" & "DEMO TEST" 'ทดสอบ อย่าลืมเอาออก
        nLine = 18
    End If
    If Len(Trim(tTagDetail)) > 255 Then
        tVB_EDCDetail = Left(tTagDetail, 255)
    Else
        tVB_EDCDetail = tTagDetail
    End If
    nLine = 19
'    Select Case UCase(tVB_TenderCode)
'        Case UCase("T002"), UCase("T003")
'            tCardType = SP_EDCxCheckTag(aTag, n, "55")
'            tCardType = SP_GETtCardType(tCardType)
'            If tCardType = "" Then
'                tVB_TenderName = otbCardType.Text
'            Else
'                tVB_TenderName = tCardType
'            End If
'        Case Else
'    End Select
    If bW_C209 = True Then '*Bump 58-04-30
        'no action
    Else
        Select Case UCase(tVB_TenderCode)
            Case UCase("T002")
                '*Eaw 56-07-08 CommSheet TK-ISS3000-248-TR.xls แยกเคส Cup(5501C) กับอื่นๆ
    '            tCardType = SP_EDCxCheckTag(aTag, n, "55")
                tCardType = SP_EDCxCheckTag(aTag, n, "56") '*Eaw 56-09-03  CommSheet TK-ISS3000-330 เปลี่ยนไปใช้ Tag 56
                
                '*Eaw 56-09-10 CommSheet TK-ISS3000-349-UAT.xls Concept ใหม่สำหรับเครดิตการ์ดแยก T002 กับ T003 ด้วย Tag56='I'
                tCardTag55 = SP_EDCxCheckTag(aTag, n, "55")
                nLine = 20
                If UCase(tCardTag55) = "E" Then
                    If W_CHKbFreeRangCard(tCardTag55, tRefCardType, tVB_TenderCode) Then
                        tCardType = tCardTag55
                        tVB_CrdType = tRefCardType
                        bVB_CNCrdEzLink = True
                    End If
                Else
                    bVB_CNCrdEzLink = False
                End If
                
                If Not bVB_CNCrdEzLink Then
                    If UCase(tCardType) = "I" Then
                        Call SP_MSGnShowing("tms_CN005253", nCS_Warning)
                        otbAcc.Text = ""
                        otbNum.Text = ""
                        otbApvCode.Text = ""
                        orbMonth.Text = ""
                        orbYear.Text = ""
                        otbCardType.Text = ""
                        Unload Me
                        Exit Sub
                    Else
                        If W_CHKbFreeRangCard(tCardType, tRefCardType, tVB_TenderCode) Then
                            nLine = 21
                            tVB_CrdType = tRefCardType
                            otbCardType.Text = tVB_TenderName
                        ElseIf W_CHKbAccByID(tRefCardType) Then
                        'If W_CHKbAccByID(tRefCardType) Then
                            tVB_CrdType = tRefCardType
                            otbCardType.Text = tVB_TenderName
'*KT 67-01-29 Downgrade CR Bank Instant Point Redemption
'                        ElseIf SP_EDCxCheckTag(aTag, n, "63") <> vbNullString Then '*KT 66-12-07 CR Bank Instant Point Redemption
'                            tVB_TenderName = SP_EDCxCheckTag(aTag, n, "54")
'                            tVB_CrdType = tVB_TenderName
'                            otbCardType.Text = tVB_TenderName
                            
                        Else
                            Call SP_MSGnShowing("tms_CN005065", nCS_Warning)
                            otbAcc.Text = ""
                            otbNum.Text = ""
                            otbApvCode.Text = ""
                            orbMonth.Text = ""
                            orbYear.Text = ""
                            otbCardType.Text = ""
                            Unload Me
                            Exit Sub
                        End If
                    End If
                End If
                nLine = 22
                '*Tao 62-07-19 Redeem Tender
                tHolderName = SP_EDCxCheckTag(aTag, n, "98") 'Name of Card Holder/Payee
                Call SP_CRTxLogMallCardEvent("Credit", "EDCResponse", "SplitResponseTag98=", tHolderName)
                'nAmount = (SP_VALcText2Double(SP_EDCxCheckTag(aTag, n, "04")) / 100) 'Transaction Amount
                tTmp = SP_EDCxCheckTag(aTag, n, "63") 'Transaction Info
                nLine = 23
                Call SP_CRTxLogMallCardEvent("Credit", "EDCResponse", "SplitResponseTag63=", tTmp)
                If tTmp <> "" Then
                    'Tag62=Invoice , Tag63=Transaction Info
                    If (SP_EDCxCheckTag(aTag, n, "62") = Left(tTmp, 6)) Then 'ถ้า 6 ตัวแรกของ Tag63 = Invoice ถือว่าเป็นการ Redeem (เผื่อมีข้อมุลอื่นที่ไม่ใช้การ Redeem ส่งออกมาจาก EDC)
                        nLine = 24
                        cAmount = SP_VALcText2Double(Mid(tTmp, 7, 12)) / 100 'Sale Amount
                        cRedeem = SP_VALcText2Double(Mid(tTmp, 19, 12)) / 100 'Redeem Amount
                        cBalance = cAmount - cRedeem
                        tTmp = "Date:" & Date & " Time:" & Time$ & " Card Holder:" & tHolderName & " Sale Amount:" & SP_FMTcCurPoint(cAmount, False, False) & " Redeem Amount:" & SP_FMTcCurPoint(cRedeem, False, False) & " Balance Amount:" & SP_FMTcCurPoint(cBalance, False, False)
                    End If
                Else
'                    'ทดสอบ อย่าลืมเอาออก
'                    cAmount = cVB_PayCredit 'Sale Amount
'                    If cVB_PayCredit > 10 Then
'                        cRedeem = 10 'Redeem Amount
'                    Else
'                        cRedeem = 1 'Redeem Amount
'                    End If
'                    cBalance = cAmount - cRedeem
'                    '==============
                    nLine = 25
                    If tHolderName <> "" Then
                        tTmp = "Date:" & Date & " Time:" & Time$ & " Card Holder:" & tHolderName
                    Else
                        tTmp = ""
                    End If
                End If
                nLine = 26
                If tTmp <> "" Then Call SP_CRTxLogMallCardEvent("Credit", "EDCResponse", "Redeem Debug:", tTmp)
                otbRedeem.Text = SP_FMTcCurPoint(cRedeem, True, True) 'เก็บยอดที่ Redeem
                '=============================
                nLine = 27
            Case UCase("T003")
                tCardType = SP_EDCxCheckTag(aTag, n, "56")
                If UCase(tCardType) <> "I" Then
                    Call SP_MSGnShowing("tms_CN005253", nCS_Warning)
                    otbAcc.Text = ""
                    otbNum.Text = ""
                    otbApvCode.Text = ""
                    orbMonth.Text = ""
                    orbYear.Text = ""
                    otbCardType.Text = ""
                    Unload Me
                    Exit Sub
                ElseIf W_CHKbAccByID(tRefCardType) Then
                    tVB_CrdType = tRefCardType
                    otbCardType.Text = tVB_TenderName
                Else
                    Call SP_MSGnShowing("tms_CN005065", nCS_Warning)
                    otbAcc.Text = ""
                    otbNum.Text = ""
                    otbApvCode.Text = ""
                    orbMonth.Text = ""
                    orbYear.Text = ""
                    otbCardType.Text = ""
                    Unload Me
                    Exit Sub
                End If
                '------------------------------------------------------------------
    '            Select Case UCase(tCardType)
    '                Case "C"
    '                    If W_CHKbCupsCard(tCardType, tRefCardType) Then
    '                        tVB_CrdType = tRefCardType
    '                        tVB_TenderName = Trim(otbCardType.Text) '*Eaw 56-07-09
    '                        '*Eaw 56-09-05  CommSheet TK-ISS3000-330
    '                        ocbType.AddItem SP_EDCxCheckTag(aTag, n, "56"), 1
    '                        ocbType.Text = SP_EDCxCheckTag(aTag, n, "56")
    '                    Else
    '                        Call SP_MSGnShowing("tms_CN005065", nCS_Error)
    '                        otbAcc.Text = ""
    '                        otbNum.Text = ""
    '                        otbApvCode.Text = ""
    '                        orbMonth.Text = ""
    '                        orbYear.Text = ""
    '                        otbCardType.Text = ""
    '                        Unload Me
    '                        Exit Sub
    '                    End If
    '                Case Else
    '                    If W_CHKbFreeRangCard(tCardType, tRefCardType, tVB_TenderCode) Then '*Eaw 56-09-03  CommSheet TK-ISS3000-330 แยกเคส Free Rang กับ เคสปกติ
    '                        tVB_CrdType = tRefCardType
    '                        otbCardType.Text = tVB_TenderName
    '                        tVB_TenderName = Trim(otbCardType.Text)  '*Eaw 56-09-05
                            '*Eaw 56-09-09 Free rang card ให้ยึดการพิมพ์ตามแบบ Credit card ทั่วไป
    '                        tCardType = SP_EDCxCheckTag(aTag, n, "56")
    '                        If UCase(tCardType) = UCase("T") Then     ''*Eaw 56-05-04 CommSheet TK-ISS3000-086 00.00.07.xls Page 20130502
    '                            tCardType = SP_GETtCardTypeEx(tCardType, "56")
    '                        Else
    '                            tCardType = ""
    '                        End If
    '                        If tCardType = "" Then
    '                            tCardType = SP_EDCxCheckTag(aTag, n, "54")
    '                        End If
    '                        If tCardType = "" Then                                                                   '*Eaw 56-03-29 ให้หา tag 54 ก่อน 55
    '                            tCardType = SP_EDCxCheckTag(aTag, n, "55")
    '                            tCardType = SP_GETtCardTypeEx(tCardType, "55")
    '                        End If
    '                        If tCardType = "" Then
    '                            tVB_TenderName = otbCardType.Text
    '                        Else
    '                            tVB_TenderName = tCardType
    '                        End If
    
    '                        otbCardType.Text = tVB_TenderName '*Eaw 56-09-10 CommSheet TK-ISS3000-349-UAT.xls Credit Type จะพิมพ์ตาม Card Desc ใน master
    '                    Else
    '                        '*Eaw 56-09-05  CommSheet TK-ISS3000-330
    '                        ocbType.AddItem SP_EDCxCheckTag(aTag, n, "56"), 1
    '                        ocbType.Text = SP_EDCxCheckTag(aTag, n, "56")
                            
    '                        tCardType = SP_EDCxCheckTag(aTag, n, "56")
    '                        If UCase(tCardType) = UCase("T") Then     ''*Eaw 56-05-04 CommSheet TK-ISS3000-086 00.00.07.xls Page 20130502
    '                            tCardType = SP_GETtCardTypeEx(tCardType, "56")
    '                        Else
    '                            tCardType = ""
    '                        End If
    '            '            tCardType = SP_GETtCardTypeEx(tCardType, "56")  '*Eaw 56-05-04 CommSheet TK-ISS3000-086 00.00.07.xls  Page 20130502
    '            '            If tCardType = "" Then                                                                   '*Eaw 56-03-29 ให้หา tag 54 ก่อน 55
    '            '                tCardType = SP_EDCxCheckTag(aTag, n, "55")
    '            '                tCardType = SP_GETtCardTypeEx(tCardType, "55")
    '            '            End If
    '                        If tCardType = "" Then
    '                            tCardType = SP_EDCxCheckTag(aTag, n, "54")
    '                        End If
    '                        If tCardType = "" Then                                                                   '*Eaw 56-03-29 ให้หา tag 54 ก่อน 55
    '                            tCardType = SP_EDCxCheckTag(aTag, n, "55")
    '                            tCardType = SP_GETtCardTypeEx(tCardType, "55")
    '                        End If
    '                        If tCardType = "" Then
    '                            tVB_TenderName = otbCardType.Text
    '                        Else
    '                            tVB_TenderName = tCardType
    '                        End If
    '                        If Not W_CHKbAccByID(tRefCardType) Then '*Eaw 56-05-13 CommSheet TK-ISS3000-167-ST 00.00.02.xls
    '                            Call SP_MSGnShowing("tms_CN005065", nCS_Error)
    '                            otbAcc.Text = ""
    '                            otbNum.Text = ""
    '                            otbApvCode.Text = ""
    '                            orbMonth.Text = ""
    '                            orbYear.Text = ""
    '                            otbCardType.Text = ""
    '                            Unload Me
    '                            Exit Sub
    '                        Else
    '                            tVB_CrdType = tRefCardType  '*Eaw 56-06-13  CommSheet TK-ISS3000-086 00.00.08.xls
    '                            '*Eaw 56-09-10 CommSheet TK-ISS3000-349-UAT.xls Credit Type จะพิมพ์ตาม Card Desc ใน master
    '                            otbCardType.Text = tVB_CrdType
    '                            tVB_TenderName = tVB_CrdType
    '                        End If
    '
    '                    End If
    '            End Select
    
        End Select
    End If
    '-------------------------------------------------------------------------------------------
    'เรียก ok
    bW_EdcCheck = False  'เพื่อไม่ให้ทำการเช็คเลขที่บัตรอีก เมื่อกดปุ่ม Ok(แบบ Manual)
    tW_EDCRun = SP_EDCtRunning(tW_EDCRun, False)
    tVB_CrdCardNo = Trim(otbAcc.Text)
    tVB_CrdEvnID = Left(Trim(tVB_CrdType), 12)             '*Ao 2014/06/05   ใช้แทน
    nLine = 28
'     If Not W_CHKbAccByID(tRefCardType) Then '*Eaw 56-05-13 CommSheet TK-ISS3000-167-ST 00.00.02.xls
'        Call SP_MSGnShowing("tms_CN005065", nCS_Error)
'        otbAcc.Text = ""
'        otbNum.Text = ""
'        otbApvCode.Text = ""
'        orbMonth.Text = ""
'        orbYear.Text = ""
'        otbCardType.Text = ""
'        Unload Me
'        Exit Sub
'    Else
'        tVB_CrdType = tRefCardType  '*Eaw 56-06-13  CommSheet TK-ISS3000-086 00.00.08.xls
'    End If
    If bW_C209 = True Then '*Bump 58-04-30
        If Not W_PRCxFindCrdByAccNum(tVB_CrdCardNo) Then
                Call W_CHKxCaseERROR("3")
                SleepEx 1000, 0
                otbAcc.Text = ""
                otbNum.Text = ""
                otbApvCode.Text = ""
                orbMonth.Text = ""
                orbYear.Text = ""
                otbCardType.Text = ""
                Unload Me
                Exit Sub
        End If
        dVB_CrdDateExp = DateSerial(2000 + Val(Format(orbYear.Text)), Val(Format(orbMonth.Text)) + 1, 1)  '*Eaw 56-06-21CommSheet TK-ISS3000-238-ST.xls เพิ่มการตรวจสอบวันหมดอายุ
        If CDate(Format(dVB_CrdDateExp, "yyyy/MM/dd") < CDate(Format(dVB_CurrentDate, "yyyy/MM/dd"))) Then
            Call SP_MSGnShowing("tms_CN005143", nCS_Warning)
            otbAcc.Text = ""
            otbNum.Text = ""
            otbApvCode.Text = ""
            orbMonth.Text = ""
            orbYear.Text = ""
            otbCardType.Text = ""
            Unload Me
            Exit Sub
        End If
        '* Bump 58-05-06 Default Balance Amount
        'otbNum.Text = SP_FMTcCurPoint(cVB_Balance, True, True)
        If Trim(otbNum.Text) = "" Then
        otbNum.Text = SP_FMTcCurPoint(cVB_PayCredit, True, True)   '*Bump 58-05-11
        End If
        '++++++++++++++
        nLine = 29
        If otbApvCode.Text = "" Then
            MsgBox "Please entry approve code", tCS_CNMsgWarn, tVB_CNAppTitle
             otcForm1(1).Enabled = True    'lock text box
            otcForm1(2).Enabled = True     'lock ปุ่ม
            Call SP_OBJxSetFocus(otbApvCode)
            ocmCancel.Enabled = False
            ocmEDC.Enabled = False
            Exit Sub
        End If
    Else
        nLine = 30
        '*Eaw 56-07-08 CommSheet TK-ISS3000-248-TR.xls แยกเคส EzLink กับอื่นๆ
        tChkCardType = SP_EDCxCheckTag(aTag, n, "55")
        Select Case UCase(tChkCardType)
        Case "E"
            'none
            nLine = 31
        Case Else
            nLine = 32
            '*KT 67-01-29 Downgrade CR Bank Instant Point Redemption
            'If orbYear.Text <> "XX" And orbMonth.Text <> "XX" Then '*KT 66-12-07 CR Bank Instant Point Redemption
                dVB_CrdDateExp = DateSerial(2000 + Val(Format(orbYear.Text)), Val(Format(orbMonth.Text)) + 1, 1)  '*Eaw 56-06-21CommSheet TK-ISS3000-238-ST.xls เพิ่มการตรวจสอบวันหมดอายุ
                If CDate(Format(dVB_CrdDateExp, "yyyy/MM/dd") < CDate(Format(dVB_CurrentDate, "yyyy/MM/dd"))) Then
                    Call SP_MSGnShowing("tms_CN005143", nCS_Warning)
                    otbAcc.Text = ""
                    otbNum.Text = ""
                    otbApvCode.Text = ""
                    orbMonth.Text = ""
                    orbYear.Text = ""
                    otbCardType.Text = ""
                    Unload Me
                    Exit Sub
                End If
            'End If
        End Select
        End If
    nLine = 33
'    bChk = W_PRCbFindCrdByMEM(tVB_CrdCardNo)
    Call ocmOK_Click            'สมบูรณ์แล้ว จะ unload โดยอัติโนมัติ
    nLine = 34
    Exit Sub
ErrorHandles:
    Call SP_CRTxLogMallCardEvent("Credit", "EDCResponse Err:" & Err.Number & ",Desc:" & Err.Description, "nLine=", nLine)   '*Thirapong 64-12-10    Log monitor CommSheet TTK-ISS3000-794 TK-ISSUE - Error 5 from EDC Response
    '*BG 57-01-15
    ofaSwipeCard.Left = 2900
    ofaSwipeCard.Top = 5400
    tVB_CreditGroup = ""
    cVB_Pay = 0
    tVB_CrdCardNo = ""
    tVB_CrdApprCode = ""
    tVB_CrdType = ""
    tVB_CrdTypeName = ""
    cVB_PayCredit = 0
'    Call SP_MSGxShowError(Err.Number, Err.Description)
    Call SP_MSGxShowError(Err.Number, Err.Description & "NOT PROPERLY from EDC Response")
    Unload Me

End Sub

Private Function W_EDCtGetString2Send(ByVal ptType$) As String
'-----------------------------------------------------------
'   Cmt : get message to send
'   Create by : *Korn 55-10-12 Edit เพิ่ม procedure ใหม่
'-----------------------------------------------------------
    Dim tSend$, tAmt$
    Dim aPayAmt() As String
    Dim tDec As String

    tSend = ""

'    tAmt = Trim(Str(cVB_Balance))   'เปลี่ยนยอดเงินให้เป็น string
    tAmt = Trim(Str(cVB_PayCredit))   'เปลี่ยนยอดเงินให้เป็น string
    If InStr(tAmt, ".") > 0 Then
        '*Eaw 56-10-07
        aPayAmt = Split(tAmt, ".")
        tDec = aPayAmt(1) & "00"
        tDec = Left(tDec, 2)
        tAmt = aPayAmt(0) & tDec
        '----------------------------------------------------
    Else
        tAmt = tAmt & "00"      'ใส่ 0 ตามหลัง 2 ตัวเป็นทศนิยมแทน
    End If
    tAmt = Replace(tAmt, ".", "")        'เอาทศนิยมออก
    tAmt = "000000000000" & tAmt    'ทั้งหมด
    tAmt = Right(tAmt, 12)                     'ต้องการ 12 หลัก
    tW_EDCRun = SP_EDCtRunning("", True)
    'แยกตามประเภทการส่ง
    'C200 = Sale, C201=PreAuth, C203=Refund, C300=Void
    Select Case UCase(ptType)
        Case "SALE"
            tSend = "C200"                      'command
            tSend = tSend & "0412"       'tag 04 ยอดเงิน จำนวน 12 หลัก ,    fix
            tSend = tSend & tAmt            'ยอดเงิน รวมทศนิยม
            tSend = tSend & "5706"        'tag 57 ตาราง tag 1.6.4 จำนวน 06 หลัก, fix
            tSend = tSend & tW_EDCRun
'            tSend = tSend & "000001"    'sale fix 000001    ??ไม่แน่ใจว่าจะต้องเป็น running หรือไม่
'            If bW_ChkocmEDC Then
'                tSend = "C209"                      'command
''''''''                tSend = tSend & "0412"       'tag 04 ยอดเงิน จำนวน 12 หลัก ,    fix
''''''''                tSend = tSend & tAmt            'ยอดเงิน รวมทศนิยม
''''''''                tSend = tSend & "5706"        'tag 57 ตาราง tag 1.6.4 จำนวน 06 หลัก, fix
'                tSend = tSend & tW_EDCRun
'                bW_ChkocmEDC = False
'            Else
'                tSend = "C200"                      'command
'                tSend = tSend & "0412"       'tag 04 ยอดเงิน จำนวน 12 หลัก ,    fix
'                tSend = tSend & tAmt            'ยอดเงิน รวมทศนิยม
'                tSend = tSend & "5706"        'tag 57 ตาราง tag 1.6.4 จำนวน 06 หลัก, fix
'                tSend = tSend & tW_EDCRun
'            End If
        Case "PREAUTH"
            tSend = "C202"                      'command
            tSend = tSend & "0412"       'tag 04 ยอดเงิน จำนวน 12 หลัก ,    fix
            tSend = tSend & tAmt            'ยอดเงิน รวมทศนิยม
            tSend = tSend & "5706"        'tag 57 ตาราง tag 1.6.4 จำนวน 06 หลัก, fix
            tSend = tSend & tW_EDCRun
'            tSend = tSend & "000001"    'sale fix 000001    ??ไม่แน่ใจว่าจะต้องเป็น running หรือไม่
        Case "REFUND"
            tSend = "C203"                      'command
            tSend = tSend & "0412"       'tag 04 ยอดเงิน จำนวน 12 หลัก ,    fix
            tSend = tSend & tAmt            'ยอดเงิน รวมทศนิยม
            tSend = tSend & "5706"        'tag 57 ตาราง tag 1.6.4 จำนวน 06 หลัก, fix
            tSend = tSend & tW_EDCRun
'            tSend = tSend & "000001"    'sale fix 000001    ??ไม่แน่ใจว่าจะต้องเป็น running หรือไม่
        Case "VOID"
            tSend = "C300"                      'command
            'To do
        Case "BP"
            tSend = "C209"                      'command
    End Select
    W_EDCtGetString2Send = tSend
    Call SP_CRTxLogMallCardEvent("Credit", "EDCSendString", "SendStringValue=", tSend) '*Eaw 56-10-07 เก็บ Log
End Function

Private Function W_EDCtReturnCase(ByVal pnRet As Long) As String
'-----------------------------------------------------------
'   Cmt : check return message
'   Create by : *Korn 55-10-08 Edit เพิ่ม procedure ใหม่
'-----------------------------------------------------------
    Dim tMsg$
    
    tMsg = ""
    
    Select Case pnRet
        Case 0: tMsg = "PASS"
        Case 2: tMsg = "COM Port not found."
        Case 5: tMsg = "COM Port in use by another process."
        Case 13: tMsg = "Invalid data from Terminal(Bad Checksum)."
        Case 24: tMsg = "Command length is too long(Maximum = 1024 bytes)."
        Case 258: tMsg = "Terminal can not be reached thorugh this port."
        Case 592: tMsg = "Data is not accepted by Terminal(Bad Checksum)."
'        Case 1460: tMsg = "Time Out"
        Case 1460: tMsg = "Timeout error and not Transaction aborted by user." '*BG 56-11-11 เปลี่ยนข้อความเพื่อให้สื่อมากขึ้น
'        Case Else: tMsg = "Port Error"
        Case Else: tMsg = "No Response from EDC"
    End Select
    
    W_EDCtReturnCase = tMsg
    
End Function

Private Sub W_EDCxMagic3CreditCard(ByVal ptComPort$, ByVal pnSecWait As Long, ByVal ptType$)
'-----------------------------------------------------------
'   Cmt : EDC for Magic3
'              ** Bump ยืนยันว่าไม่มีการแบ่งจ่าย ถ้าเป็น EDC ให้รับยอดเต็มทั้งหมด    '*Ao แบ่งจ่ายได้
'   Create by : *Korn 55-10-08 Edit เพิ่ม procedure ใหม่
'-----------------------------------------------------------
    Dim oMagic3 As New MagIC3Connector
    Dim nGet As Long, tMsg$, tSend$, tRet$
    Dim tCut As String
    
    On Error GoTo ErrHandle
    
    bW_ChkESC = False           '*Ao 2013/12/03 CommSheet TK-ISS3000-FAPL-161
    pnSecWait = 100000
    If pnSecWait = 0 Then pnSecWait = 100000     'ถ้าค่าเป็น  0 ให้ตั้งค่ารอที่ 40 วินาที

    '*ได้ทดสอบแล้วว่า ถ้าใช้ Dll จากของเขายังไงก็รอ ไม่ว่าจะเป็น demo(.Net ที่gเขาให้มา) ก็ตาม
    'ดังนั้น ก็ไม่จำเป็นจะต้อง loop หรือให้ Timer ทำงาน จึงต้องปิด otmEDC(T2) ก่อนเลยเพื่อให้มีการเรียกฟังก์ชั่นนี้แค่รอบเดียว, แต่ถ้าจะทำซ้ำค่อยเปิด enable ใหม่
    otmForm.Enabled = False  'เผื่อไว้ ปกติก็ปิดอยู่แล้ว T1
    otmEDC.Enabled = False  'ปิด T2 ก่อน
    
    'แสดงว่าส่ง echo ผ่านแล้ว, ให้ส่งมูลค่าไปต่อ
    If bW_C209 = True Then
         tSend = W_EDCtGetString2Send("BP")        'หายอดเงินที่จะส่งไป    ?? fix SALE จริงๆต้องแยกตาม type
    Else
         tSend = W_EDCtGetString2Send(ptType)        'หายอดเงินที่จะส่งไป    ?? fix SALE จริงๆต้องแยกตาม type
    End If
    'tSend = W_PRCtRandom  '*KT 66-12-06  ต้องเอาออก
    If tSend = "" Then
        tMsg = "Sending string is empty"
        Call SP_EDCxINSERTLog(tMsg)
        '----------------------------------------------------------------------
        '*Bump 57-05-12 Phase2 เก็บไว้ตรวจสอบปัญหา
        Call SP_CRTxLogMallCardExtend(("Date:" & Date & " Time:" & Time$ & " Error: " & tMsg))
        '----------------------------------------------------------------------
        Call SP_MSGnShowing("tms_CN005065" & ";" & tMsg, nCS_Error)      '*Korn ?? เปลี่ยน msg
        Unload Me
        Exit Sub
    End If
    
'*Eaw 56-07-09 CommSheet TK-ISS3000-268-TR.xls
    'check send Echo
'    nGet = oMagic3.EchoTest(ptComPort, tRet)
'    tMsg = W_EDCtReturnCase(nGet)
'    If UCase(tMsg) <> "PASS" Then
'        Call SP_EDCxINSERTLog(tMsg)
'        tMsg = tMsg & vbCrLf & "Do you want to try again?"
'        If SP_MSGnShowing("tms_CN005065" & ";" & tMsg, nCS_OkCancel) = vbOK Then        'ถ้าไม่ผ่าน, ยืนยันทำซ้ำอีกครั้ง
'            otmEDC.Enabled = True   'เปิดใหม่อีกครั้ง
'            Exit Sub
'        Else
'            Unload Me       'ปิดไปเลย กลับไปเริ่มใหม่ point-1
'            Exit Sub
'        End If
'    End If

    'ส่งค่าไป
    Call SP_CRTxLogMallCardExtend(("Date:" & Date & " Time:" & Time$ & " ComPort: " & ptComPort & " ,Send: " & tSend)) '*Tao 61-05-17
    'ptComPort = "COM2"
    nGet = oMagic3.SendCommand(ptComPort, tSend)
    Call SP_CRTxLogMallCardExtend(("Date:" & Date & " Time:" & Time$ & " Get: " & nGet))  '*Tao 61-05-17
    tMsg = W_EDCtReturnCase(nGet)
    Call SP_CRTxLogMallCardExtend(("Date:" & Date & " Time:" & Time$ & " Return: " & tMsg))  '*Tao 61-05-17
    If UCase(tMsg) <> "PASS" Then
        'อาจจะได้ case error 1460 TimeOut
        Call SP_EDCxINSERTLog(tMsg)
        '----------------------------------------------------------------------
        '*Bump 57-05-12 Phase2 เก็บไว้ตรวจสอบปัญหา
        Call SP_CRTxLogMallCardExtend(("Date:" & Date & " Time:" & Time$ & " Error: " & tMsg))
        '----------------------------------------------------------------------
        tMsg = tMsg & vbCrLf & "Do you want to try again?"
        If SP_MSGnShowing("tms_CN005065" & ";" & tMsg, nCS_OkCancel) = vbOK Then       'ถ้าไม่ผ่าน, ยืนยันทำซ้ำอีกครั้ง
            otmEDC.Enabled = True   'เปิดใหม่อีกครั้ง
            bW_ChkESC = True       '*Ao 2013/12/03 CommSheet TK-ISS3000-FAPL-161
            Exit Sub
        Else
            Unload Me       'ปิดไปเลย
            Exit Sub
        End If
    End If
    Call SP_CRTxLogMallCardExtend(("Date:" & Date & " Time:" & Time$ & " Response: "))  '*Tao 61-05-17
    'แสดงว่าส่ง echo ผ่านแล้ว, ส่งมูลค่าผ่านแล้ว, ให้ส่งรอ response
    tRet = ""   'reset ค่า
    nGet = oMagic3.RecvResponse(ptComPort, tRet, pnSecWait)
   ' tRet = tSend '*KT 66-12-06  ต้องเอาออก
    Call SP_CRTxLogMallCardExtend(("Date:" & Date & " Time:" & Time$ & " Response: " & nGet))  '*Tao 61-05-17
    tMsg = W_EDCtReturnCase(nGet)
    Call SP_CRTxLogMallCardExtend(("Date:" & Date & " Time:" & Time$ & " Return: " & tMsg))  '*Tao 61-05-17
    If UCase(tMsg) <> "PASS" Then
           'อาจจะได้ case error 1460 TimeOut
        Call SP_EDCxINSERTLog(tMsg)
        '---------------------------
         '*Bump 57-05-12 Phase2 เก็บไว้ตรวจสอบปัญหา
        Call SP_CRTxLogMallCardExtend("Date:" & Date & " Time:" & Time$ & " Error: " & tMsg)
        '---------------------------
        tMsg = tMsg & vbCrLf & "Do you want to try again?"
        If SP_MSGnShowing("tms_CN005065" & ";" & tMsg, nCS_OkCancel) = vbOK Then       'ถ้าไม่ผ่าน, ยืนยันทำซ้ำอีกครั้ง
            otmEDC.Enabled = True   'เปิดใหม่อีกครั้ง
            bW_ChkESC = True       '*Ao 2013/12/03 CommSheet TK-ISS3000-FAPL-161
            Exit Sub
        Else
            Unload Me       'ปิดไปเลย
            Exit Sub
        End If
    Else        'ผ่าน
        If Trim(tRet) = "" Then   'ผ่าน แต่ไม่ได้ค่า
            Call SP_EDCxINSERTLog(tMsg)
            '---------------------------
            '*Bump 57-05-12 Phase2 เก็บไว้ตรวจสอบปัญหา
            Call SP_CRTxLogMallCardExtend(("Date:" & Date & " Time:" & Time$ & " Error: " & tMsg))
            '---------------------------
            tMsg = tMsg & vbCrLf & "Do you want to try again?"
            If SP_MSGnShowing("tms_CN005065" & ";" & tMsg, nCS_OkCancel) = vbOK Then       'ผ่าน แต่ไม่ได้ค่า return, ยืนยันทำซ้ำอีกครั้ง
                otmEDC.Enabled = True   'เปิดใหม่อีกครั้ง
                bW_ChkESC = True       '*Ao 2013/12/03 CommSheet TK-ISS3000-FAPL-161
                Exit Sub
            Else
                Unload Me       'ปิดไปเลย
                Exit Sub
            End If
        Else
            'ได้ค่ามาแล้ว, return กลับที่ tRet โดย
             'อาจจะได้ case แคชเชียร์ ไม่กดอะไร รอจนเครื่อง end(รอจน TimeOut)
             'อาจจะได้ case แคชเชียร ์ยังไม่ได้รูดอะไร กด end เลย
             'อาจจะได้ case แคชเชียร ์รูดบัตรไปครึ่งทาง แล้วกด end
               'อาจจะได้ case แคชเชียร ์รูดบัตรไปครึ่งทาง แล้วไม่กดอะไรเลย(รอจน TimeOut)
             'อาจจะได้ case แคชเชียร์ รูดบัตรผิดประเภท(ไม่ใช่บัตรเครดิต)
             'อาจจะได้ case แคชเชียร์ รูดบัตรสมบูรณ์
            cVB_Pay = 0
            tVB_CrdCardNo = ""
            tVB_CrdApprCode = ""
             'ส่งค่าไป split
            tCut = Left(tRet, 4)
            If bW_C209 = True Then
                Call SP_CRTxLogMallCardEvent("CreditC209", "EDCResponse", "ResponseValue=", tRet) 'Bump*56-04-26 เก็บ Log
            Else
                Call SP_CRTxLogMallCardEvent("Credit", "EDCResponse", "ResponseValue=", tRet) 'Bump*56-04-26 เก็บ Log
            End If
            
            If UCase(tCut) <> UCase("R922") Then
                'Call W_EDCxSplitResponse(Trim(tRet)) '*Tao 62-12-10 ของเดิม
                Call W_EDCxSplitResponse(tRet) '*Tao 62-12-10 ไม่ Trim
                '-----------------------------------------------------------
                '*Bump 57-05-12 Phase2 เก็บไว้ตรวจสอบปัญหา
                '* PRJ10001TK-PH2-External Spec-Goods Return-00.02.06 FSBT.docx
                '* Page.8
                '* TK-ISS3000-463 TK-ISS3000-FAPL-182   EDC error message and response code
                'Call SP_EDCxSplitResponseExtend("C", Trim(tRet)) '*C = Credit card, "E" = Staff discount/Employee discount '*Tao 62-12-10 ของเดิม
                Call SP_EDCxSplitResponseExtend("C", tRet) '*C = Credit card, "E" = Staff discount/Employee discount '*Tao 62-12-10 ไม่ Trim
                Call SP_CRTxLogMallCardExtend(("Date:" & Date & " Time:" & Time$ & " tVB_Transaction:" & tVB_Transaction))  '*Bump 57-11-04 TK_Issue 27/10/2017 รับชำระด้วย Credit ทำให้ Transaction Type เปลี่ยนจาก 03 เป็น 31
                '-----------------------------------------------------------
            Else
                If UCase(tCut) = UCase("R922") Then
                    'Call SP_EDCxSplitResponseExtend("CE", Trim(tRet)) '*Bump 57-07-07 *C = Credit card, "E" = Staff discount/Employee discount, "CE" คือ เอา staff discount cad มารูดที่ credit '*Tao 62-12-10 ของเดิม
                    Call SP_EDCxSplitResponseExtend("CE", tRet) '*Bump 57-07-07 *C = Credit card, "E" = Staff discount/Employee discount, "CE" คือ เอา staff discount cad มารูดที่ credit '*Tao 62-12-10 ไม่ Trim
                    Call SP_CRTxLogMallCardExtend(("Date:" & Date & " Time:" & Time$ & " tVB_Transaction:" & tVB_Transaction))  '*Bump 57-11-04 TK_Issue 27/10/2017 รับชำระด้วย Credit ทำให้ Transaction Type เปลี่ยนจาก 03 เป็น 31
                End If
                Call SP_MSGnShowing("tms_CN005065" & ";" & "Invalid Card", nCS_Error)
                Unload Me       'ปิดไปเลย
            End If
'            Call W_EDCxSplitResponse(Trim(tRet))
        End If
    End If
    Exit Sub
ErrHandle:
    cVB_Pay = 0
    tVB_CrdCardNo = ""
    tVB_CrdApprCode = ""
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Sub

Private Function W_PRCxFindCrdByAccNum(ByVal ptAccNum$) As Boolean
'--------------------------------------
'Call:
'Cmt: ค้นหา Format ของเลขที่บัตรเครดิตว่าอยุในช่วงใด
'การค้นหาบัตรเครดิตที่มีช่วงหมายเลขทับซ้อนกันหมากกว่า 1 ประเภทบัตรเครดิตทำให้ระบบดึงประเภทบัตรเครดิตตาม record ที่เกิดก่อนหลัง จึงมีบัตรที่เป็น Master Card 2 ประเภท คือ
'Master Inter  = 5000000000000001-5999999999999999
'SCB MC =        5577550000000000-5577729999999999
'ดังนั้นไม่จะระบุหมายเลขในช่วง SCB MC อย่างไรก็จะได้ Master Inter
'Order By FNCdcDigitChk จากมาก - น้อย
'--------------------------------------
    Dim tSql$
    Dim orsTemp As ADODB.Recordset
    Dim tCrdType As String
    
On Error GoTo ErrHandle '*KT 52-02-11  PH2-ML-SF-092 โปรแกรมหลุดมาที่ Desktop ตอนพิมพ์ receipt , ชำระด้วย Credit
    W_PRCxFindCrdByAccNum = False
    nW_CaseERROR = 0
    tVB_CrdStaBPSale = "Y"
    tVB_CrdStaBPReturn = "Y"
'    tSql = "SELECT  TOP 1 TCNMCreditCard.FCCdcAmtMin, TCNMCreditCard.FCCdcAmtMax, TCNMCreditCard.FTCdcOnTop, TCNMCreditCard.FTCdcMallCard,"
'    tSql = tSql & " TCNMCreditCard.FNCdcOntopRate, TCNMCreditCard.FNCdcDigitChk, TCNMCreditCard.FNCdcDigits, TCNMCreditCard.FTCdcDescRe,"
'    tSql = tSql & " TCNMCreditCard.FTCdcDescCredit , TCNMCreditCard.FTCdcCreditType, TCNMCreditCard.FTTdmCode,"
'    tSql = tSql & " TCNMCreditCard.FTCdcDepCode,TCNMCreditCard.FTCdcReg,TCNMCreditCard.FTCdcPro,"
'    tSql = tSql & " TCNMCreditCardRange.FTCcrRangeFrom,TCNMCreditCardRange.FTCcrRangeTo,"
'    tSql = tSql & " TCNMCreditCard.FDCdcStartDate,TCNMCreditCard.FDCdcEndDate,"
'    tSql = tSql & " TCNMCreditCard.FTCdcPOSNoOnTopFlag,TCNMCreditCard.FTCdcPOSNoAlwType,TCNMCreditCard.FTCdcFloorAlwOnTop,"  '*KT 54-06-01 V.4.0.0  Req.19 OnTop Specify POSNO
'    tSql = tSql & " TCNMCreditCard.FTCdcStaSubAllOnTop,"  '*KT 54-06-01 V.4.0.0  Req.19 OnTop Specify POSNO
'    tSql = tSql & " TCNMCreditCard.FTCdcStaTier,TCNMCreditCard.FTCdcStaTime,TCNMCreditCard.FTCdcStaAlwTierAmtLast,"  '*KT 54-06-01 V.4.0.0  Req.19 OnTop Specify POSNO
'    tSql = tSql & " TCNMCreditCard.FTEvnCode,"  '*KT 54-11-17 V.4.0.5 DCR_5411003 PRJ09010TheMall-R4-SE-External Spec.Upload Bonus Points-00.04.00-FSBT.doc Event ID ของบัตรเครดิตการ์ด"
'    tSql = tSql & " TCNMCreditCard.FTCdcType"    '*Eaw 55-11-29
    
'    tSql = "SELECT TOP 1 TCNMCreditCard.*"                '*Ao 2013/02/22   CommSheet TK-ISS3000-011
'    tSql = tSql & " ,TCNMCreditCardRange.FTCcrRangeFrom,TCNMCreditCardRange.FTCcrRangeTo"
'    tSql = tSql & " FROM TCNMCreditCard INNER JOIN"
'    tSql = tSql & " TCNMCreditCardRange ON TCNMCreditCard.FTTdmCode = TCNMCreditCardRange.FTTdmCode AND"
'    tSql = tSql & " TCNMCreditCard.FTCdcCreditType = TCNMCreditCardRange.FTCdcCreditType"
'    tSql = tSql & " WHERE TCNMCreditCard.FTTdmCode = '" & tVB_TenderCode & "'"
'    '*Eaw 56-01-29
''    If bW_ManaulMode And tVB_CNMnuCurrent <> "1831" Then  '*Eaw 55-12-18 -- 1831  คือ TC Card
''        tSQL = tSQL & " AND TCNMCreditCard.FTCdcType = '" & tVB_CrdCardType & "'"
''    End If
''---------------------------------------------------------------------------------------------------------------------------------------------
'    tSql = tSql & " AND LEFT('" & ptAccNum & "' ,TCNMCreditCard.FNCdcDigitChk) BETWEEN LEFT(TCNMCreditCardRange.FTCcrRangeFrom,TCNMCreditCard.FNCdcDigitChk)  AND LEFT(TCNMCreditCardRange.FTCcrRangeTo,TCNMCreditCard.FNCdcDigitChk)  "
'    tSql = tSql & " AND LEN('" & ptAccNum & "') = TCNMCreditCard.FNCdcDigits"
'
'    tSql = tSql & " ORDER BY TCNMCreditCard.FNCdcDigitChk DESC" '*KT 51-12-11 Phase ข้อ 14 Communication Sheet PH2-ML-SF-025 Credit Card Range.xls

'    tSql = W_PRCxStrSQL(ptAccNum)           '*Ao 2013/05/03
    If bVB_CNCrdEzLink Then  '*Eaw 56-09-10 เอามาครอบถ้าเป็น EzLink จะตรวจสอบแบบใหม่
        tSql = "SELECT TOP 1 * FROM TCNMCreditCard"
        tSql = tSql & " WHERE FTCdcType = 'E' AND FTRemark = 'Y'"
    Else
    '*BG 56-11-11 กรณี TC Card กรองการใส่เลขที่ไม่ถูกต้องออกเช่น เลขน้อยกว่า 15 หลัก TC Card ไม่มีแน่นอน
        If (Trim(ocbType.Text) = "-") Or (tVB_CNMnuCurrent = "1831") Then
            If tVB_CNMnuCurrent = "1831" And Len(ptAccNum) < 15 Then
                W_PRCxFindCrdByAccNum = False
                nW_CaseERROR = 3
                Exit Function
            Else
                tSql = W_PRCxStrSQL(ptAccNum)
            End If
        Else
            tSql = W_PRCxStrCupsCard(ptAccNum)
        End If
    End If
    
    Call SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine, , , , , True, True) 'read master
    If Not (orsTemp.BOF Or orsTemp.EOF) Then
        tVB_CreditGroup = SP_FEDtChkString(orsTemp, "FTCdcCreditGroup")         '*Ao 2013/11/05
           '*BG 56-11-26 CommSheet TK-ISS3000-433-FAPL-158_Bung_561126.xls
'        otbCardType.Text = SP_FEDtChkString(orsTemp, "FTCdcDescCredit")
'        tVB_TenderName = SP_FEDtChkString(orsTemp, "FTCdcDescCredit")
        '----------------------------------------------------
'        If tVB_TenderCode = "T017" Then
'            Select Case SP_VALnText2Double(tVB_TransactionSub)
'                Case tEN_TrnReturn, tEN_TrnDepositCancel, tEN_TrnAddDepositCancel, tEN_TrnCancelVoucher
'                    otbCardType.Text = SP_FEDtChkString(orsTemp, "FTCdcDescRe")
'                    tVB_TenderName = SP_FEDtChkString(orsTemp, "FTCdcDescRe")
'                Case tEN_TrnSale, tEN_TrnDeposit, tEN_TrnAddDeposit, tEN_TrnVoucher, tEN_TrnRedeem  ' เช่น Sale ก็มีปัญหากรณีไม่ได้ค่า Description ใช้  FTCdcDescCredit
'                    otbCardType.Text = SP_FEDtChkString(orsTemp, "FTCdcDescCredit")
'                    tVB_TenderName = SP_FEDtChkString(orsTemp, "FTCdcDescCredit")
'                Case Else
'                    otbCardType.Text = SP_FEDtChkString(orsTemp, "FTCdcDescCredit")
'                    tVB_TenderName = SP_FEDtChkString(orsTemp, "FTCdcDescCredit")
'            End Select
'        Else
'            otbCardType.Text = SP_FEDtChkString(orsTemp, "FTCdcDescCredit")
'            tVB_TenderName = SP_FEDtChkString(orsTemp, "FTCdcDescCredit")
'        End If
        '----------------------------------------------------
        Select Case SP_VALnText2Double(tVB_TransactionSub)                      '*Ao 2014/06/04       Return
            Case tEN_TrnReturn, tEN_TrnDepositCancel, tEN_TrnAddDepositCancel, tEN_TrnCancelVoucher
                otbCardType.Text = SP_FEDtChkString(orsTemp, "FTCdcDescRe")
                tVB_TenderName = SP_FEDtChkString(orsTemp, "FTCdcDescRe")
            Case Else
                otbCardType.Text = SP_FEDtChkString(orsTemp, "FTCdcDescCredit")
                tVB_TenderName = SP_FEDtChkString(orsTemp, "FTCdcDescCredit")
        End Select
        '----------------------------------------------------
        tVB_CrdType = SP_FEDtChkString(orsTemp, "FTCdcCreditType")
        tVB_CrdKey = SP_FEDtChkString(orsTemp, "FTTdmCode")
        cVB_CrdRate = SP_FEDvChkNumeric(orsTemp, "FNCdcOntopRate")
'        tVB_IsFreeRange = SP_FEDtChkString(orsTemp, "FTRemark")   '*Eaw 56-05-15 CommSheet TK-ISS3000-159-IT2 00.00.03.xls
        Select Case tVB_TenderCode   '*Eaw 56-06-18
            Case "T002"
                tVB_IsFreeRange = SP_FEDtChkString(orsTemp, "FTRemark")   '*Eaw 56-05-15 CommSheet TK-ISS3000-159-IT2 00.00.03.xls
            Case Else
                tVB_IsFreeRange = ""
        End Select
'        Select Case Val(Format(tVB_TransactionSub))
'            Case tEN_TrnReturn, tEN_TrnCancelVoucher, tEN_TrnDepositCancel, tEN_TrnAddDepositCancel, tEN_TrnReturnTaxInc, tEN_TrnReturnManualTax
'                tVB_TenderName = SP_FEDtChkString(orsTemp, "FTCdcDescRe")
'            Case Else
'                tVB_TenderName = SP_FEDtChkString(orsTemp, "FTCdcDescCredit")
'        End Select
        '*Ao 2013/02/22 '*Ao 2013/02/22   CommSheet TK-ISS3000-011
        tVB_CrdStaReEntry = SP_FEDtChkString(orsTemp, "FTTdmStaReEntry")
        If bVB_ReEntryMode Then
            If UCase(tVB_CrdStaReEntry) <> UCase("Y") Then
                W_PRCxFindCrdByAccNum = False
                nW_CaseERROR = 1
                Exit Function
            End If
        End If
        tVB_CrdStaBPSale = SP_FEDtChkString(orsTemp, "FTCdcStaBPSale")
        tVB_CrdStaBPReturn = SP_FEDtChkString(orsTemp, "FTCdcStaBPReturn")
        tVB_CrdStaMallCard = SP_FEDtChkString(orsTemp, "FTCdcMallCard")
        If UCase(tVB_CrdStaMallCard) <> UCase("Y") Then
            If SP_CHKbAllowMemEmpCrd(tVB_TenderCode) Then
                W_PRCxFindCrdByAccNum = False
                nW_CaseERROR = 2
                Exit Function
            End If
        End If
        '--------------------------------------------------------------------------------------------------------------------------------
        '*KT 54-06-01 V.4.0.0  Req.19 OnTop Specify POSNO
        bVB_CrdPOSNoOnTopFlag = IIf(SP_FEDtChkString(orsTemp, "FTCdcPOSNoOnTopFlag") = "Y", True, False)
        tVB_CrdPOSNoOnTopAlwType = SP_FEDtChkString(orsTemp, "FTCdcPOSNoAlwType")
        tVB_CrdFloorAlwOnTop = SP_FEDtChkString(orsTemp, "FTCdcFloorAlwOnTop")
        tVB_CrdStaTierOntopType = IIf(SP_FEDtChkString(orsTemp, "FTCdcStaTier") = "Y", "1", "0") '*KT 54-11-29 V.4.0.6 แก้ไขการได้  Ontop  FTCdcStaTier = Y (Tier %) FTCdcStaTime=Y(Tier Amt)  กำหนดให้  tVB_CrdStaTierOntopType = '' คือ Tier Amt โดย Default
       ' tVB_CrdStaTierOntopType = IIf(SP_FEDtChkString(orsTemp, "FTCdcStaTime") = "Y", "1", "0") '*KT 54-11-29 V.4.0.6 แก้ไขการได้  Ontop ใช้ FTCdcStaTier หรือ FTCdcStaTime อย่างใดอย่างหนึ่ง
        tVB_CrdStaAlwTierAmtLast = SP_FEDtChkString(orsTemp, "FTCdcStaAlwTierAmtLast")
'        tVB_CrdEvnID = SP_FEDtChkString(orsTemp, "FTEvnCode") '*KT 54-11-17 V.4.0.5 DCR_5411003 PRJ09010TheMall-R4-SE-External Spec.Upload Bonus Points-00.04.00-FSBT.doc Event ID ของบัตรเครดิตการ์ด
        tVB_CrdEvnID = Left(Trim(tVB_CrdType), 12)             '*Ao 2014/06/05    ใช้แทน
        tVB_CrdCardType = SP_FEDtChkString(orsTemp, "FTCdcType")   '*Eaw 56-01-29
        
'        tCrdType = SP_GETtCardType(tVB_CrdCardType)
        '*Eaw 56-05-15 CommSheet TK-ISS3000-159-IT2 00.00.03.xls
'        If UCase(tVB_CrdCardType) = UCase("C") And Trim(ocbType.Text) = "-" Then '*Eaw 56-05-07 CommSheet TK-ISS3000-159-IT2 00.00.02.xls
'            ocbType.Text = tVB_CrdType
'            If Trim(ocbType.Text) = "-" Then  '*Eaw 56-05-07 เพิ่มเติมกรณีหาด้วยหมายเลข CUPS แต่ไม่เลือก Type CUPS
'                GoTo ErrHandle
'            End If
'        End If
'--------------------------------------------------------------------------------
'        If tCrdType = "" Then        '*Ao 2013/03/01  CommSheet TK-ISS3000-086 00.00.01_Aoo.xlsx   '*Eaw 56-04-05   CommSheet TK-ISS3000-086 00.00.05 FSBT.xlsx
'            Select Case Val(Format(tVB_TransactionSub))
'                Case tEN_TrnReturn, tEN_TrnCancelVoucher, tEN_TrnDepositCancel, tEN_TrnAddDepositCancel, tEN_TrnReturnTaxInc, tEN_TrnReturnManualTax
'                    tVB_TenderName = SP_FEDtChkString(orsTemp, "FTCdcDescRe")
'                Case Else
'                    tVB_TenderName = SP_FEDtChkString(orsTemp, "FTCdcDescCredit")
'            End Select
'        Else
'            tVB_TenderName = tCrdType
'        End If
        W_PRCxFindCrdByAccNum = True
    Else
        tVB_CreditGroup = ""
        otbCardType.Text = ""
        tVB_CrdKey = ""
        tVB_CrdType = ""
        cVB_CrdRate = 0
        
        '*KT 54-06-01 V.4.0.0  Req.19 OnTop Specify POSNO
        bVB_CrdPOSNoOnTopFlag = False
        tVB_CrdPOSNoOnTopAlwType = ""
        tVB_CrdFloorAlwOnTop = ""
        tVB_CrdStaTierOntopType = ""
        tVB_CrdStaAlwTierAmtLast = ""
        tVB_CrdEvnID = "" '*KT 54-11-17 V.4.0.5 DCR_5411003 PRJ09010TheMall-R4-SE-External Spec.Upload Bonus Points-00.04.00-FSBT.doc Event ID ของบัตรเครดิตการ์ด

    End If
    oDisc.orsOntop = orsTemp
    Call SP_DATxRsNothing(orsTemp)
    Exit Function
ErrHandle:
    Call SP_DATxRsNothing(orsTemp)
    W_PRCxFindCrdByAccNum = False
    otbCardType.Text = ""
    tVB_CrdKey = ""
    tVB_CrdType = ""
    cVB_CrdRate = 0
    '*KT 54-06-01 V.4.0.0  Req.19 OnTop Specify POSNO
    bVB_CrdPOSNoOnTopFlag = False
    tVB_CrdPOSNoOnTopAlwType = ""
    tVB_CrdFloorAlwOnTop = ""
    tVB_CrdStaTierOntopType = ""
    tVB_CrdStaAlwTierAmtLast = ""
    tVB_CrdEvnID = "" '*KT 54-11-17 V.4.0.5 DCR_5411003 PRJ09010TheMall-R4-SE-External Spec.Upload Bonus Points-00.04.00-FSBT.doc Event ID ของบัตรเครดิตการ์ด
End Function

Private Sub Command1_Click()
    Dim tRet As String

'    tRet = "R20004120000000100003902EC5706000024"
'    tRet = "R200021610000800441743060412000000000100071008221535351404111537121308221535323806      39020041089715100342159797151000000025405EZL-P5501E57060000316336000000000665000000000100000000000565"
'    tRet = "R20002161000080044174306041200000000100007100822153644140411153712130822153641380612345639020041089715100342159797151000000025405EZL-P5501E57060000316336000000000565000000000001000000000564"
'    tRet = "R20002194563510800021529366041200000000100007100902152458140416082201M3712DEMODEMODEMO3806DEMO5139020041081234567842158006500000000005205TXAUD5404VISA5501V5601T57060000506206000304634400000000100000366100000033500888888888888AUD640600000196101101010004" 'CUP Free Rang card
'    tRet = "R200021610000800441743060412000000000100071008221535351404111537121308221535323806      39020041089715100342159797151000000025405EZL-P5501E57060000316336000000000665000000000100000000000565"
   
'    tRet = "R200021610000800441743060412000000000100071008221535351404111537121308221535323806      39020041089715100342159797151000000025405EZL-P5501E57060000316336000000000665000000000100000000000565" 'Ezlink
'    tRet = "R20002194563510800021529366041200000000100007100902152458140416082201M3712DEMODEMODEMO3806DEMO5139020041081234567842158006500000000005205TXAUD5404VISA5501V5601T57060000506206000304634400000000100000366100000033500888888888888AUD640600000196101101010004"
'    tRet = "R20002165520380016678530041200000005000007100909142350140413122201C3712DEMODEMODEMO3806DEMO5139020041081234567842151681680000000005204IP1253620000008000E800A0000000041010  MasterCard      01191829A6CC3AE55410MASTERCARD5501M5601I57060000656206000327640600000296101110000004"
'    tRet = "R20002194563510800021529366041200000000100007100902150934140416082201M3712DEMODEMODEMO3806DEMO5139020041081234567842151683573000000005203CUP5403CUP5501C5601I57060000496206000302640600000396101100000004" 'HP CUP
'    tRet = "R20002165414963706119128041200000000010007100625102004140416012201C3712DEMODEMODEMO3806DEMO5139020041081234567842158006500000000005205TXTHB53620000008000E800A0000000041010  MASTERCARD      48CC4623D8BE7D5B5410MASTERCARD5501M5601T57060000016206000272634400000000233207645233247033500888888888888THB640600000396101101100004" 'Master
'    tRet = "R20002165124920000045038041200000000010007100625102004140415032201C3712DEMODEMODEMO3806DEMO5139020041081234567842158006500000000005205TXTHB53620000008000E800A0000000041010  MASTERCARD      48CC4623D8BE7D5B5410MASTERCARD5501M5601T57060000016206000272634400000000233207645233247033500888888888888THB640600000396101101100004" 'TanachatMC
'    tRet = "R20002194563510800021529366041200000000100007100902150934140416082201M3712DEMODEMODEMO3806DEMO5139020041081234567842151683573000000005203CUP5403CUP5501C5601C57060000496206000302640600000396101100000004" 'CUP
'    tRet = "R200021610000800441743060412000000000100071008221535351404111537121308221535323806      39020041089715100342159797151000000025405EZL-P5501E57060000316336000000000665000000000100000000000565" 'EzLink
   
'    tRet = "R20002165414963706119128041200000000010007100625102004140416012201C3712DEMODEMODEMO3806DEMO5139020041081234567842158006500000000005205TXTHB53620000008000E800A0000000041010  MASTERCARD      48CC4623D8BE7D5B5410MASTERCARD5501M5601I57060000016206000272634400000000233207645233247033500888888888888THB640600000396101101100004" 'HP Master
'    tRet = "R20002165124920000045038041200000000010007100625102004140415032201C3712DEMODEMODEMO3806DEMO5139020041081234567842158006500000000005205TXTHB53620000008000E800A0000000041010  MASTERCARD      48CC4623D8BE7D5B5410MASTERCARD5501M5601I57060000016206000272634400000000233207645233247033500888888888888THB640600000396101101100004" 'HP Thanachat
   
'    tRet = "R20002165414963706119128041200000000010007100625102004140416012201C3712DEMODEMODEMO3806DEMO5139020041081234567842158006500000000005205TXTHB53620000008000E800A0000000041010  MASTERCARD      48CC4623D8BE7D5B5410MASTERCARD5501M5601T57060000016206000272634400000000233207645233247033500888888888888THB640600000396101101100004"
'    tRet = "R20002165124920000045038041200000000010007100625102004140415032201C3712DEMODEMODEMO3806DEMO5139025441081234567842158006500000000005205TXTHB53620000008000E800A0000000041010  MASTERCARD      48CC4623D8BE7D5B5410MASTERCARD5501M5601I57060000016206000272634400000000233207645233247033500888888888888THB640600000396101101100004" 'HP Thanachat
'    tRet = "R20002165124920000045038041200000000010007100625102004140415032201C3712DEMODEMODEMO3806DEMO5139025141081234567842158006500000000005205TXTHB53620000008000E800A0000000041010  MASTERCARD      48CC4623D8BE7D5B5410MASTERCARD5501M5601I57060000016206000272634400000000233207645233247033500888888888888THB640600000396101101100004" 'HP Thanachat
'    tRet = "R20002165124920000045038041200000000010007100625102004140415032201C3712DEMODEMODEMO3806DEMO5139025441081234567842158006500000000005205TXTHB53620000008000E800A0000000041010  MASTERCARD      48CC4623D8BE7D5B5410MASTERCARD5501M5601M57060000016206000272634400000000233207645233247033500888888888888THB640600000396101101100004" 'HP Thanachat
'    tRet = "R20002165124920000045038041200000000010007100625102004140415032201C3712DEMODEMODEMO3806DEMO513902Y341081234567842158006500000000005205TXTHB53620000008000E800A0000000041010  MASTERCARD      48CC4623D8BE7D5B5410MASTERCARD5501M5601I57060000016206000272634400000000233207645233247033500888888888888THB640600000396101101100004" 'HP Thanachat

'    tRet = "R20002165124920000045038041200000000010007100625102004140415032201C3712DEMODEMODEMO3806DEMO5139025441081234567842158006500000000005205TXTHB53620000008000E800A0000000041010  MASTERCARD      48CC4623D8BE7D5B5410MASTERCARD5501M5601M57060000016206000272634400000000233207645233247033500888888888888THB640600000396101101100004" 'HP Thanachat
'    tRet = "R20002165124920000045038041200000000010007100625102004140415032201C3712DEMODEMODEMO3806DEMO513902Y341081234567842158006500000000005205TXTHB53620000008000E800A0000000041010  MASTERCARD      48CC4623D8BE7D5B5410MASTERCARD5501M5601M57060000016206000272634400000000233207645233247033500888888888888THB640600000396101101100004" 'HP Thanachat
'    tRet = "R20004120000000050003902Y35706000055" 'Y3
'    tRet = "R20004120000000050003902545706000055" '54
'    tRet = "R20004120000000050003902515706000055" '51
'    tRet = "R20004120000000050003902995706000055" '99 Unknown
'    tRet = "R20004120000001110003902TA5706000200" 'TA
'    tRet = "R200021610000800441743060412000000000010071010281158381404111537121310281158353806      39020041089715100342159797151000000025405EZL-P5501E57060001096336000000000041000000000010000000000031"
'    tRet = "R20002164182382020637777041200000000176007101031130326140416062201C3712330405013817380690410139020041083797879542121681683318725203DBS53620000008000F800A0000000031010  VISA            E060CF9B45BAC7B45404VISA5501V5601D57060003276206013817640600071096101100000004"
'    tRet = "C20004120000000017605706000327"
'    tRet = "R200021610000900173654080412000000000001071011061641541404031837121311061641513806      39020041089715100342159797151000000025405EZL-P5501E57060001956336000000000380000000000001000000000379"
'    tRet = "R2000215377911240257173041200000004000007101028155050140414092201C3712DEMODEMODEMO3806DEMO5139020041081234567842159805110000000005204AMEX53620000008000E800A0000000250104  DBS Amex        BDF923020E57F42D5406TKAMEX5501A5601A57060002006206000491640600000596101100000063" 'TK Amex
'    tRet = "R20002164417701001218629041200000003400007101110161216140414092201C3712331408081955380600643239020041083791577842121681680458475203DBS53620000008000F800A0000000031010  VISA PLATINUM   5CD637B28CBF64985404VISA5501V5601D57060005676206081955640600100096101100000004"
'    tRet = "R20004120000000100003902XX5706000112" 'XX
'    tRet = "R20004120000000050023902TA5706000111"  'TA
    'tRet = "R20002165577912115757302041200000000956807100113112954140418082201M3712401303001369380609411339020041083791582442121681680458395203DBS5410MASTERCARD5501M5601D5706011126620600136964060000259610110000000498226    VILIA S PUTRI    "
    
'    tRet = "R2000245#gTAjLBBtYd1K03zzhSNn+uqsofhcoaYUwNnFh2QjHX4=0412000000000001071009041054581425#0OjPrIAyg8sYFRNZLP/DUg==2201C371276593321626238063777433902004108160688884211168168188885203DBS53620000008000E800A000000004101001DBS MasterCard  56A6F475EE6006C65410MASTERCARD5501M5601D570600000162060000386406000008961011000000149826 /RAJESH CHAURASIA        "
'
'    Call W_EDCxSplitResponse(tRet)
'    bW_ManaulMode = False
End Sub

Private Sub Command2_Click()
'If Trim(Text1.Text) = "" Then Exit Sub
'Dim tInput$
'tInput = "R2000245#pHTMSwbxjEM1xJhVItm5rEkVRHFS913gDe6NkglDAtk=0412000000026280071012091439301425#+Bms/biUQQMc1x00yuAmuA==2201P3712134306032837380600200939020041083791603542121681683259245205TDVIS53620000000000    A000000003101000DBS VISA        468E6A049F83EAD75406TDVISA5501V5601D57061361396206032837640600340596101100000004"
'tInput = "R2000245#DalCcyWiXfB24NWHoRnM8y3uansYMFuH2g39tW8UcyI=0412000000010190071011301207491425#s2zhNhOQiwG76iDf5otK5Q==2201P37128119373685583806807586390200410837915875421098056090885204AMEX53620000008000E800A000000025010901AMERICAN EXPRESS896355BFD1AC6CA25404AMEX5501A5601A57060809466206031551640600393596100100000063"
'tInput = "R2000245#7xfJZGRbZ0g9A9chagDAK+CqDPHfOBANMZgjodr3pjY=0412000000006912071012101239081425#179K6y6YTlErLSknlAp0HA==2201P37125806996747573806687364390200410837915825421098056090885204AMEX53620000008000E800A000000025010402DBS Amex        B46FA9035C0F27865406TKAMEX5501A5601A57060191946206000304640600186796100100000063"
'Call W_EDCxSplitResponse(tInput)

'MsgBox SP_EDCtDecrypt(Trim(Text1.Text))
End Sub

'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'    Select Case KeyCode
'            Case vbKeyEscape
'                If bW_HandleESC Then
'                    Set oEdc = Nothing
'                    bW_ManaulMode = True
'                    ocmCancel.Cancel = True
'                    otbAcc.Locked = False
'                    otbAcc.BackColor = nVB_CNClrBUnLock
'                    Call SP_OBJxSetFocus(otbAcc)
'                End If
'    End Select
'End Sub
'*Ao 2012/10/31
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If tVB_CNMnuCurrent <> "1831" Then     '*Eaw 56-01-03                           1831 : TC
        If bW_ChkESC Then               '*Ao 2013/12/03 CommSheet TK-ISS3000-FAPL-161
            Select Case KeyCode
                Case vbKeyEscape
                    If bW_HandleESC Then    '*Eaw 55-11-29
                        '*Korn 55-10-10 Edit 9 บรรทัดล่าง
                        bW_HandleESC = False    'อาจจะไม่จำเป็น เพราะจะ stop timer แทน
                        otmForm.Enabled = False
                        otmEDC.Enabled = False
                        ofa3Sec.Left = 120      'ซ่อน
                        ofa3Sec.Top = 5400
                        ofaSwipeCard.Left = 2880        'ซ่อน
                        ofaSwipeCard.Top = 5400
                        otcForm1(1).Enabled = True     'unlock text box
                        otcForm1(2).Enabled = True     'unlock ปุ่ม
                        
                        Set oEdc = Nothing
                        bW_ManaulMode = True
                        ocmCancel.Cancel = True
                        otbAcc.Locked = False
                        otbAcc.BackColor = nVB_CNClrBUnLock
                        'otbNum.Text = SP_FMTcCurPoint(cVB_Balance, True, True)
                        otbNum.Text = SP_FMTcCurPoint(cVB_PayCredit, True, True)   '*Bump 58-05-11
                        
                        Call SP_OBJxSetFocus(otbAcc)
                        
    '                    ocbType.Enabled = True '*Eaw 55-11-29                      '*Ao 2013/01/09
                    Else
                        Call ocmCancel_Click                    '*Ao 2013/01/09
                    End If
    
                 '*Korn 55-10-10 Edit รอทำเพิ่มเติมส่วนของ PRE-AUTH
                Case vbKeyF1   'PRE-AUTH
                    'To Do
                    otmForm.Enabled = False     'stop T1
                    tW_EdcAction = "PREAUTH"
                    otmEDC.Enabled = True
            End Select
        End If
    Else
        Select Case KeyCode
            Case vbKeyEscape
                Call ocmCancel_Click                    '*Ao 2013/01/09
        End Select
    End If
End Sub

Private Sub Form_Load()
On Error GoTo ErrHandle
    Dim tTime As String
    Dim orsTemp As ADODB.Recordset
    Dim tEdcPort As String
    Dim tSql As String
    bW_C209 = False
    If tVB_CNMnuCurrent = "1831" Then bW_ManaulMode = True '* Bump 56-04-11 TC Card
    bW_ChkESC = True         '*Ao 2013/12/03 CommSheet TK-ISS3000-FAPL-161
    tVB_EDCDetail = ""
    tVB_CrdEvnID = ""               '*Ao 2014/06/05
    cVB_Pay = 0
    tVB_CrdCardNo = ""
    tVB_CrdApprCode = ""
    bW_EdcCheck = True  '*Korn 55-10-16 (1 บรรทัด) Default จะเป็น True แต่ถ้ารูดบัตรผ่าน EDC จะให้ค่าเป็น False
    bW_KeyEnter = False
    bW_ChkCrd = True
    cVB_AdjustMalCrdOfOntop = 0
    cVB_AdjustDiscMalCrdOfCrd = 0
    cVB_CreditRedeemAmt = 0 '*Tao 62-07-19 Clear Redeem
    bVB_CNCrdEzLink = False '*Eaw 56-10-28
    Call SP_FRMxChildCenter(Me)
    If bVB_Res640X480 Then
        Me.Top = 2500
    Else
        Me.Top = 2900
    End If
    Call SP_FRMxGetCaptionFrmDB2Tag(Me)
    Call SP_FRMxSetCapForm(Me)
    Call W_FRMxSetInitial
    Call W_OBJxAddDataToCombo    '*Eaw 55-11-29
    Me.Caption = aVB_TenderMnu(nVB_CurTenderSeqNo).Menu
'    DoEvents            '*KT 53-01-14
    'Call oOposMsr.ReadMSR(tVB_CashMsr)
    
    '*Bump 58-04-24 'CommSheet TK-ISS3000-558.02_22042015.xls
    
    bW_EnableMSR = True
    If tVB_CNMnuCurrent = "1817" Then
        bW_EnableMSR = False
    End If
    If tVB_CNMnuCurrent = "1818" Then
        bW_EnableMSR = False
    End If
    '**************
    If tVB_CNMnuCurrent = "1831" Then '*Bump 58-05-06 ซ่อนปุ่ม EDC กรณี TC Card
        ocmEDC.Visible = False
        ocmEDC.Enabled = False

        ocmOk.Left = ocmCancel.Left - 60
        ocmCancel.Left = ocmEDC.Left - 60
    End If
    If bVB_OPOSMSREnable And bW_EnableMSR Then  'bVB_CNMsrPos
    'If bVB_OPOSMSREnable Then
        With OPOSMSR1
            'Open the device
            'Use a Logical Device Name which has been set on the SetupPOS.
            .Open tVB_CashMsr
    
            'Get the exclusive control right for the opened device.
            'Then the device is disable from other application.
            '(Notice:When using an old CO, use the Claim.)
'            .ClaimDevice 0 'TramPos3000
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
'    If bVB_OPOSScannerEnable Then              '*Ao 2013/11/07
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
    'olaForm(7).Caption = ""
    olaForm(8).Caption = ""
    olaForm(9).Caption = ""
    otbAcc.Locked = True
    otbAcc.BackColor = nVB_CNClrBLock
    nW_Timer = 0
    '*Eaw 55-12-18 กรณีเป็น TC card
     ' 1831 คือ TC Card
    If tVB_CNMnuCurrent <> "1831" Then
        '--- Credit Card ติดต่อกับ EDC
        '--- Hire Purchase/Credit Installment  ติดต่อกับ EDC
        bW_HandleESC = True
        bW_ManaulMode = False
        '*BG อ่านค่าจาก TSysConfig แทน เดิมเอาค่าจาก Register
'        tTime = GetSetting("AdaPosMall", "POSFront", "TimeEDCEsc")
        tTime = tVB_EDCWait   '*BG 56-11-12 ป้องกันค่าว่าง แต่เงื่อนไขยังปล่อยไว้เหมือนเดิม
        If tTime = "" Then
'            nW_TimeESC = 3
'            Call SaveSetting("AdaPosMall", "POSFront", "TimeEDCEsc", "3")
            nW_TimeESC = CLng(tVB_EDCWait)
            Call SaveSetting("AdaPosMall", "POSFront", "TimeEDCEsc", tVB_EDCWait)
        Else
            If CLng(tTime) > 0 Then
                nW_TimeESC = CLng(tTime)
            Else
'                nW_TimeESC = 3
'                Call SaveSetting("AdaPosMall", "POSFront", "TimeEDCEsc", "3")
                nW_TimeESC = CLng(tVB_EDCWait)
                Call SaveSetting("AdaPosMall", "POSFront", "TimeEDCEsc", tTime)
            End If
        End If
'        tTime = GetSetting("AdaPosMall", "POSFront", "TimeEDCExpire")
        tTime = tVB_EDCTimeOut     '*BG 56-11-12 ป้องกันค่าว่าง แต่เงื่อนไขยังปล่อยไว้เหมือนเดิม
        If tTime = "" Then
'            nW_TimeExpireEDC = 60
'            Call SaveSetting("AdaPosMall", "POSFront", "TimeEDCExpire", "60")
                nW_TimeESC = CLng(tVB_EDCTimeOut)
                Call SaveSetting("AdaPosMall", "POSFront", "TimeEDCEsc", tVB_EDCTimeOut)
        Else
            If CLng(tTime) > 0 Then
                nW_TimeExpireEDC = CLng(tTime)
            Else
'                nW_TimeExpireEDC = 60
'                Call SaveSetting("AdaPosMall", "POSFront", "TimeEDCExpire", "60")
                nW_TimeESC = CLng(tVB_EDCTimeOut)
                Call SaveSetting("AdaPosMall", "POSFront", "TimeEDCEsc", tVB_EDCTimeOut)
            End If
        End If
        '****
        Select Case tVB_TypeOfEDC
            Case "1"
                'Dim oEdcHyperCom As New AdaEDC.cHyperCom
                'Set oEdc = oEdcHyperCom
                Call SP_OBJbCreate(oEdc, "AdaEDC.cHyperCom")
            Case "2"
                'Dim oEdcPosNet As New AdaEDC.cPosNet
                'Set oEdc = oEdcPosNet
                Call SP_OBJbCreate(oEdc, "AdaEDC.cPosNet")
            Case "3"
                'Dim oEdcVerifone  As New AdaEDC.cVerifone
                'Set oEdc = oEdcVerifone
                Call SP_OBJbCreate(oEdc, "AdaEDC.cVerifone")
            Case "4" 'EDC
                '*Ao 2012/10/30
'                tSql = "SELECT *"
'                tSql = tSql & vbCrLf & "FROM TCNMTerminalMtn"
'                tSql = tSql & vbCrLf & "WHERE FTTmnNum = '" & tVB_CNTerminalNum & "'"
'                Call SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine, , , , , True, True) 'read master
'                If Not (orsTemp.BOF And orsTemp.EOF) Then
'                    If SP_FEDtChkString(orsTemp, "FTTmnEdc") = "1" Then
'                        bVB_CNEdcPos = True
'                        tEdcPort = SP_FEDtChkString(orsTemp, "FTTmnEdcPort")
'                        tW_Comport = SP_EDCtSetPort(tEdcPort)
'                        Select Case Val(Format(tVB_TransactionSub))
'                            Case tEN_TrnReturn, tEN_TrnReturnTaxInc, tEN_TrnReturnManualTax
'                                tW_EdcAction = "REFUND"
'                            Case Else
'                                tW_EdcAction = "SALE"
'            '                    tW_EdcAction = "VOID"
'            '                    tW_EdcAction = "PREAUTH"
'                        End Select
'                    Else
'                        bVB_CNEdcPos = False
'                    End If
'                Else
'                    bVB_CNEdcPos = False
'                End If
                Select Case Val(Format(tVB_TransactionSub))
                    Case tEN_TrnReturn, tEN_TrnReturnTaxInc, tEN_TrnReturnManualTax, tEN_TrnDepositCancel, tEN_TrnAddDepositCancel, tEN_TrnOtherDisbursement    '*Ao 2013/11/02   เพิ่ม tEN_TrnDepositCancel, tEN_TrnAddDepositCancel
                        tW_EdcAction = "REFUND"
                    Case tEN_TrnCancelVoucher
                        tW_EdcAction = "REFUND"
                    Case Else
                        tW_EdcAction = "SALE"
    '                    tW_EdcAction = "VOID"
    '                    tW_EdcAction = "PREAUTH"
                End Select
                tW_Comport = tVB_EDCComport
                '________________________________________________________________________________
                If Not bVB_CNEdcPos Then
                    otbAcc.Locked = False
                    otbAcc.BackColor = nVB_CNClrBUnLock
                    bW_HandleESC = False
                    bW_ManaulMode = True
                    otmEDC.Enabled = False
                    otmForm.Enabled = False
                    Call SP_OBJxSetFocus(otbAcc)
                    ocmEDC.Enabled = False  '* NOT 58-04-16 เพิ่มปุ่ม
                Else
    '                    olaForm(7).Caption = "You're cancel EDC  within "
    '                    olaForm(8).Caption = nW_TimeESC
    '                    Call SP_OBJbCreate(oEdc, "AdaEDC.cVerifone")
    '                    otmForm.Enabled = True
    '                    otmEDC.Enabled = True
                        '*Korn 55-10-10 Edit (8 บรรทัด)
'                    olaForm(8).Caption = nW_TimeESC
'                    olaForm(9).Caption = nW_TimeESC
'                    ofa3Sec.Left = 1600     'center
'                    ofa3Sec.Top = 600
'                    ofaSwipeCard.Left = 2900    'ซ่อน
'                    ofaSwipeCard.Top = 5400
'                    otcForm1(1).Enabled = False     'lock text box
'                    otcForm1(2).Enabled = False     'lock ปุ่ม
'                    otmForm.Enabled = True
                    '*Eaw 56-06-11
                    Select Case Val(Format(tVB_TransactionSub))
                        Case tEN_TrnReturnManualTax, tEN_TrnReturnTaxInc, tEN_TrnSaleManualTax, tEN_TrnSaleTaxInc
                            otbAcc.Locked = False
                            otbAcc.BackColor = nVB_CNClrBUnLock
                            bW_HandleESC = False
                            bW_ManaulMode = True
                            otmEDC.Enabled = False
                            otmForm.Enabled = False
                            Call SP_OBJxSetFocus(otbAcc)
                            ocmEDC.Enabled = False '* NOT 58-04-16 เพิ่มปุ่ม
                        Case Else
                            olaForm(8).Caption = nW_TimeESC
                            olaForm(9).Caption = nW_TimeESC
                            ofa3Sec.Left = 1600     'center
                            ofa3Sec.Top = 600
                            ofaSwipeCard.Left = 2900    'ซ่อน
                            ofaSwipeCard.Top = 5400
                            otcForm1(1).Enabled = False     'lock text box
                            otcForm1(2).Enabled = False     'lock ปุ่ม
                            otmForm.Enabled = True
                    End Select
'-----------------------------------------------------------------------------------------------------
                End If
            Case "5" 'NETS
                If Not bVB_CNNETSPos Then
                    otbAcc.Locked = False
                    otbAcc.BackColor = nVB_CNClrBUnLock
                    bW_HandleESC = False
                    bW_ManaulMode = True
                    otmEDC.Enabled = False
                    otmForm.Enabled = False
                    Call SP_OBJxSetFocus(otbAcc)
                Else
'                    olaForm(7).Caption = "You're cancel NETS  within "
'                    olaForm(8).Caption = nW_TimeESC
'                    Call SP_OBJbCreate(oEdc, "AdaEDC.cVerifone")
'                    otmForm.Enabled = True
'                    otmEDC.Enabled = True
'*Eaw 56-06-11
                    Select Case Val(Format(tVB_TransactionSub))
                        Case tEN_TrnReturnManualTax, tEN_TrnReturnTaxInc, tEN_TrnSaleManualTax, tEN_TrnSaleTaxInc
                            otbAcc.Locked = False
                            otbAcc.BackColor = nVB_CNClrBUnLock
                            bW_HandleESC = False
                            bW_ManaulMode = True
                            otmEDC.Enabled = False
                            otmForm.Enabled = False
                            Call SP_OBJxSetFocus(otbAcc)
                        Case Else
                            olaForm(7).Caption = "You're cancel NETS  within "
                            olaForm(8).Caption = nW_TimeESC
                            Call SP_OBJbCreate(oEdc, "AdaEDC.cVerifone")
                            otmForm.Enabled = True
                            otmEDC.Enabled = True
                    End Select
                End If
            Case Else
                otbAcc.Locked = False
                otbAcc.BackColor = nVB_CNClrBUnLock
                bW_HandleESC = False
                bW_ManaulMode = True
                otmEDC.Enabled = False
                otmForm.Enabled = False
                Call SP_OBJxSetFocus(otbAcc)
            End Select
        Else
            '*Eaw 55-12-18 กรณีเป็น TC Card
            '--TC--------
            olaForm(7).Visible = False
            olaForm(8).Visible = False
            ocbType.Enabled = False
            otbAcc.BackColor = nVB_CNClrBUnLock
            otbAcc.Locked = False
'            bW_HandleESC = True
'            bW_ManaulMode = True
            otmEDC.Enabled = False  'ปิด Timmer
            otmForm.Enabled = False 'ปิด Timmer
            otbNum.Text = SP_FMTcCurPoint(cVB_Balance, True, True) '*Eaw 56-07-01CommSheet TK-ISS3000-253.xls
        End If
    Exit Sub
ErrHandle:
    Call SP_CRTxLogPrg("ErrorNumber:" & Err.Number & " ErrorDesc:" & Err.Description & " Form:" & Me.Name & " Control:Form" & " Events:Form_Load")
End Sub

Private Function W_OBJxAddDataToCombo()
'--------------------------------------------------------------------------------------
'Call:
'Cmt: '*Eaw 55-11-29 เพิ่ม Type ของ Credit Card ลงใน Combo box เพื่อให้พนักงานเลือกในกรณีที่คีย์มือ
'----------------------------------------------------------------------------------------
    Dim tSql$
    Dim tCrdNull$
    Dim orsTemp As ADODB.Recordset
    Dim orsCups As ADODB.Recordset  '*Eaw 56-05-07 CommSheet TK-ISS3000-159-IT2 00.00.02.xls
 On Error GoTo ErrHandle
    '*Eaw 56-05-07 CommSheet TK-ISS3000-159-IT2 00.00.02.xls
    ocbType.Clear
    If tVB_CNMnuCurrent = "1831" Then
        ocbType.Enabled = False
        Exit Function
    Else
        ocbType.Enabled = True
    End If
    ocbType.AddItem ("-")
    tSql = "SELECT FTCdcCreditType FROm TCNMCreditCard"
'    tSql = tSql & vbCrLf & "WHERE FTCdcType = 'C' AND FTRemark='Y'"
    tSql = tSql & vbCrLf & "WHERE FTRemark= 'Y'"   '*Eaw 56-05-15 CommSheet TK-ISS3000-159-IT2 00.00.03.xls
    tSql = tSql & vbCrLf & "AND FTTdmCode = '" & tVB_TenderCode & "'" '*Eaw 56-09-04
    Call SP_SQLvExecute(tSql, orsCups, oDB.ocnOnLine, , , , , True, True) 'read master
    If Not (orsCups.EOF And orsCups.BOF) Then
        orsCups.MoveFirst
        Do While Not orsCups.EOF
            ocbType.AddItem SP_FEDtChkString(orsCups, "FTCdcCreditType")
            orsCups.MoveNext
        Loop
    End If
    
    '*Eaw 56-05-15 CommSheet TK-ISS3000-159-IT2 00.00.03.xls  แสดงเฉพาะ Card ที่เป็น CUPS เท่านั้น
'    tCrdNull = "99"
'    tSql = "SELECT  DISTINCT FTCdcType "
'    tSql = tSql & "FROM TCNMCreditCard "
'    tSql = tSql & "WHERE ISNULL(FTCdcType, '" & tCrdNull & "')<>'" & tCrdNull & "'"
'     tSql = tSql & "AND FTCdcType<>'" & "'"
'     tSql = tSql & "AND FTCdcType<> 'C'"  '*Eaw 56-05-07 CommSheet TK-ISS3000-159-IT2 00.00.02.xls
'    Call SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine, , , , , True, True) 'read master
''    ocbType.Clear
'    If Not (orsTemp.EOF And orsTemp.BOF) Then
'        orsTemp.MoveFirst
'        Do While Not orsTemp.EOF
'             ocbType.AddItem SP_FEDtChkString(orsTemp, "FTCdcType")
'             orsTemp.MoveNext
'         Loop
'    End If
    ocbType.ListIndex = 0  '*Eaw 56-05-07 CommSheet TK-ISS3000-159-IT2 00.00.02.xls
    Exit Function
ErrHandle:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Function

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrHandle
    bW_C209 = False
    'Call oOposMsr.CloseMSR
    With OPOSMSR1
        'Cancel the device
        .DeviceEnabled = False
        'Release the device exclusive control right.
        '(Notice:When using an old CO, use the Release.)
'        .ReleaseDevice 'Tram3000
        .Release ''Tram2000
        'Finish using the device.
        .Close
    End With
'    With oScanner
'        If bVB_OPOSScannerEnable Then
'            If .DeviceEnabled Then
'                .DeviceEnabled = False
'                .Release
'                .Close
'            End If
'        End If
'    End With

    Call SP_OBJxSetNothing(Me)
    oDisc.oRsOnTopNothing = True '*KT 53-11-12 ปรับปรุง
    Set oEdc = Nothing
'    Set wCNPayCredit = Nothing  '*Korn 55-10-11 Edit (เพิ่ม code 1 บรรทัด)
    Exit Sub
ErrHandle:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Sub

Private Sub ocbType_Click()
'    MsgBox ocbType.Text
    Select Case UCase(ocbType.Text)
        'Case UCase("CUP"), UCase("Ezlink"), UCase("MCC-FX")
        Case UCase("Ezlink")
            ocmEDC.Enabled = False
        Case Else
        ocmEDC.Enabled = True
    End Select
End Sub

'Private Sub ocbType_LostFocus()
''*Eaw 55-11-29
'
'    tVB_CrdCardNo = Trim(otbAcc.Text)
'    tVB_CrdCardType = ocbType.Text
'    If bW_ManaulMode Then
'        If Not W_PRCxFindCrdByAccNum(tVB_CrdCardNo) Then
'            Call SP_MSGnShowing("tms_CN005065", nCS_Error)
'            bW_FoundAcc = False
'            otbAcc.Text = ""
'            orbMonth.Text = ""
'            orbYear.Text = ""
'            otbCardType.Text = ""
'            otbApvCode.Text = ""
'            otbNum.Text = ""
'            ocbType.Text = ""
'            Call W_OBJxAddDataToCombo    '*Eaw 56-01-03
'            Call SP_OTBxSelText(otbAcc)
'        Else
'            bW_FoundAcc = True
'            Call SP_OBJxSetFocus(orbMonth)
'        End If
'    End If
'End Sub
Private Sub ocmCancel_Click()
    Dim cNetSale  As Double
    Dim cRoundRcv As Double
    Dim cNetSaleRound As Double
    Dim cTotalDiscAft As Double
    Dim cBalance As Double
On Error GoTo ErrHandle

    If bVB_OPOSMSREnable Then OPOSMSR1.DataEventEnabled = False         '*Ao 10-02-09
'    If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = False
    tVB_CreditGroup = ""
    cVB_Pay = 0
    tVB_CrdCardNo = ""
    tVB_CrdApprCode = ""
    tVB_CrdType = ""                '*Bump 56-12-04 V.4.0.109
    tVB_CrdTypeName = ""    '*Bump 56-12-04 V.4.0.109
    cVB_PayCredit = 0 '*Eaw 56-07-10
    
    '*KT 51-08-18 เมื่อกดปุ่ม Cancel ต้องกลับมาเป็นเหมือนเดิม
    '----------------------------------------------------------------------------------------------
    If cVB_AdjustMalCrdOfOntop > 0 Or cVB_AdjustDiscMalCrdOfCrd > 0 Then
        Select Case Val(Format(tVB_TransactionSub))
            Case tEN_TrnDeposit, tEN_TrnRedeem, tEN_TrnAddDeposit, tEN_TrnDepositCancel, tEN_TrnAddDepositCancel
                 '*KT 52-02-23 640 or 800
                wSub.olaDiscTotal.Caption = SP_FMTcCurPoint(cVB_TotalDiscAft, False, True)
                wSub.olaNetSale.Caption = SP_FMTcCurPoint(cVB_Total - cVB_TotalDiscAft, False, True)
                wSub.olaBalance.Caption = SP_FMTcCurPoint(cVB_Balance, False, True)
                wSub.lblDepositAmt.Caption = SP_FMTcCurPoint(cVB_DepositAmt, False, True)
                wSub.Refresh '*KT 53-02-03 แก้ไขตาม POSFront 2.0.33
'                    DoEvents            '*KT 53-01-14
            Case Else
                 '*KT 52-02-23 640 or 800
                wSub.olaDiscTotal.Caption = SP_FMTcCurPoint(cVB_TotalDiscAft, False, True)
                wSub.olaNetSale.Caption = SP_FMTcCurPoint(cVB_NetSale, False, True)
                wSub.olaBalance.Caption = SP_FMTcCurPoint(cVB_Balance, False, True)
                wSub.Refresh '*KT 53-02-03 แก้ไขตาม POSFront 2.0.33
'                    DoEvents            '*KT 53-01-14
        End Select
    End If
    '-----------------------------------------------------------------------------------------------
    Unload Me
    Exit Sub
ErrHandle:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Sub

Private Sub ocmEDC_Click()
'    wCNPayCreditPayEx.Show vbModal
'    If bVB_C209 = True Then
'        bW_C209 = True
'        Call W_PRCxloadEDC
'    Else
'        Call SP_OBJxSetFocus(otbAcc)
'    End If
    '*Bump 58-05-06 ไม่ต้องแสดง wCNPayCreditPayEx
    bW_C209 = True
    Call W_PRCxloadEDC
End Sub

Private Sub ocmOK_Click()
    Dim cPay As Double
    Dim tSql As String
    Dim orsTemp As New ADODB.Recordset
On Error GoTo ErrHandle '*KT 52-02-11  PH2-ML-SF-092 โปรแกรมหลุดมาที่ Desktop ตอนพิมพ์ receipt , ชำระด้วย Credit
    If bVB_OPOSMSREnable And bW_EnableMSR Then OPOSMSR1.DataEventEnabled = False         '*Ao 10-02-09
'    If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = False
    
    tVB_CrdCardNo = Trim(otbAcc.Text)
'    tVB_CrdCardType = Trim(ocbType.Text)   '*Eaw  55-11-29  '*Eaw 56-05-07 CommSheet TK-ISS3000-159-IT2 00.00.02.xls
    
    If bW_ManaulMode Then      '*Eaw 55-11-29  กรณีเป็น Manaul Mode
        If Not W_PRCxFindCrdByAccNum(tVB_CrdCardNo) Then
'            Call SP_MSGnShowing("tms_CN005065", nCS_Error)
            Call W_CHKxCaseERROR(nW_CaseERROR)                  '*Ao 2013/02/22   CommSheet TK-ISS3000-011
            bW_FoundAcc = False
'            ocbType.ListIndex = 0  '*Eaw 56-05-07 CommSheet TK-ISS3000-159-IT2 00.00.02.xls
            Call SP_OBJxSetFocus(otbAcc)
            Exit Sub
        Else
            bW_FoundAcc = True
        End If
        
    Else    ''*Ao 2013/02/22
'        tVB_CrdStaBPSale = "N"
'        tVB_CrdStaBPReturn = "N"
'        tVB_CrdStaReEntry = "Y"
    End If
'If bW_C209 = True Then '*Bump 58-04-29
'    If Not W_PRCxFindCrdByAccNum(tVB_CrdCardNo) Then
'            Call W_CHKxCaseERROR(nW_CaseERROR)
'            SleepEx 1000, 0
'            otbAcc.Text = ""
'            otbNum.Text = ""
'            otbApvCode.Text = ""
'            orbMonth.Text = ""
'            orbYear.Text = ""
'            otbCardType.Text = ""
'            Unload Me
'            Exit Sub
'    End If
'End If
    '***KT 51-11-20
'    If Not bW_FoundAcc Then '*KT 52-02-12 เพิ่มประสิทธิภาพการทำงานกรณีที่มีการตรวจสอบหมายเลข Credit Card มาแล้ว ไม่ต้องตรวจสอบซ้ำอีกครั้ง
    If (bW_EdcCheck) And (Not bW_FoundAcc) Then    '*KT 52-02-12 เพิ่มประสิทธิภาพการทำงานกรณีที่มีการตรวจสอบหมายเลข Credit Card มาแล้ว ไม่ต้องตรวจสอบซ้ำอีกครั้ง
        If Not W_PRCxFindCrdByAccNum(tVB_CrdCardNo) Then
'            Call SP_MSGnShowing("tms_CN005065", nCS_Error)
            Call W_CHKxCaseERROR(nW_CaseERROR)                      '*Ao 2013/02/22   CommSheet TK-ISS3000-011
            bW_FoundAcc = False
            Call SP_OBJxSetFocus(otbAcc)
            bW_KeyEnter = False
            If bVB_OPOSMSREnable And bW_EnableMSR Then OPOSMSR1.DataEventEnabled = True          '*Ao 10-02-09
'            If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True
            Exit Sub
        Else
            bW_FoundAcc = True
        End If
    End If
    'Check Expire Date Comsheet ML-SF-170
    dVB_CrdDateExp = DateSerial(2000 + Val(Format(orbYear.Text)), Val(Format(orbMonth.Text)) + 1, 1) '*KT 51-12-09 V 1.0.3 Phase 2 ข้อ 26  ให้วันที่ที่ป้อนเป็นวันที่ 1 ของเดือนถัดไปที่ระบุ Communication Sheet PH2-ML-SF-027 Credit Expire.xls
'    If dVB_CrdDateExp <= Format(dVB_CurrentDate, "mm/dd/yyyy") Then '*KT 51-12-09 V 1.0.3 Phase 2 ข้อ 26  ให้วันที่ที่ป้อนเป็นวันที่ 1 ของเดือนถัดไปที่ระบุ Communication Sheet PH2-ML-SF-027 Credit Expire.xls
    '*Ao 52-02-05  เนื่องจาก dCurrentExpDate Format เป็น dd/mm/yyyy  ทำให้การเปรียบเทียบผิด
    'เช่น วันนี้ 05/02/2009 -> Format(05/02/2009, "mm/dd/yyyy")=02/05/2009
    '01/04/2009 <  02/05/2009
'    If CDate(Format(dVB_CrdDateExp, "YYYY/MM/DD")) < CDate(Format(dVB_CurrentDate, "YYYY/MM/DD")) Then
'    If (bW_EdcCheck) And (CDate(Format(dVB_CrdDateExp, "YYYY/MM/DD")) < CDate(Format(dVB_CurrentDate, "YYYY/MM/DD"))) Then
    If bVB_CNCrdEzLink Then '*Eaw 56-09-10 เอามาครอบถ้าเป็น EzLink จะไม่เช็ควันหมดอายุบัตร
        'none
    Else
        If (bW_EdcCheck) And (CDate(Format(dVB_CrdDateExp, "YYYY/MM/DD")) < CDate(Format(dVB_CurrentDate, "YYYY/MM/DD"))) Then    '*Ao 2013/08/26
            Call SP_MSGnShowing("tms_CN005143", nCS_Warning)
            orbMonth.Text = ""
            orbYear.Text = ""
            bW_KeyEnter = False
            Call SP_OBJxSetFocus(orbMonth)
            If bVB_OPOSMSREnable And bW_EnableMSR Then OPOSMSR1.DataEventEnabled = True           '*Ao 10-02-09
'            If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True
            Exit Sub
        End If
    End If
        
    cPay = SP_VALcText2Double(otbNum.Text)
    If cPay = 0 Then 'check amount zero
        Call SP_MSGnShowing("tms_CN005064", nCS_Error)
        bW_KeyEnter = False
        Call SP_OBJxSetFocus(otbNum)
        If bVB_OPOSMSREnable And bW_EnableMSR Then OPOSMSR1.DataEventEnabled = True          '*Ao 10-02-09
'        If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True
        Exit Sub
    End If
    
  Select Case SP_VALnText2Double(tVB_TransactionSub)
        Case tEN_TrnDeposit, tEN_TrnRedeem, tEN_TrnAddDeposit, tEN_TrnDepositCancel, tEN_TrnAddDepositCancel
            '*KT 51-11-12
            'If cPay > SP_FMTcCurPoint((cVB_Balance - cVB_AdjustMalCrdOfOntop + cVB_AdjustDiscMalCrdOfCrd - IIf(cVB_AdjustMalCrdOfOntop > 0, cVB_RoundRcv, 0)), True, True) Then 'check amount > balance
             If cPay > SP_VALcText2Double(SP_FMTcCurPoint((cVB_Balance - cVB_AdjustMalCrdOfOntop + cVB_AdjustDiscMalCrdOfCrd) - cVB_RoundRcv, True, True)) Then    'check amount > balance
                 Call SP_MSGnShowing("tms_CN005066", nCS_Error)
                 If otbNum.Locked Then
                        Call SP_OBJxSetFocus(otbAcc)
                        otbNum.Text = ""
                        otbNum.Locked = False
                        otbNum.BackColor = nVB_CNClrBUnLock
                        otbApvCode.Text = ""
                        orbMonth.Text = ""
                        orbYear.Text = ""
                        otbCardType.Text = ""
                        bVB_IsOnTop = False
                        Select Case SP_VALnText2Double(tVB_TransactionSub)
                            Case tEN_TrnDeposit, tEN_TrnRedeem, tEN_TrnAddDeposit, tEN_TrnDepositCancel, tEN_TrnAddDepositCancel
                                 '*KT 52-02-23 640 or 800
                                wSub.olaDiscTotal.Caption = SP_FMTcCurPoint(cVB_TotalDiscAft, False, True)
                                wSub.olaNetSale.Caption = SP_FMTcCurPoint(cVB_Total - cVB_TotalDiscAft, False, True)
                                wSub.olaBalance.Caption = SP_FMTcCurPoint(cVB_Balance, False, True)
                                wSub.lblDepositAmt.Caption = SP_FMTcCurPoint(cVB_DepositAmt, False, True)
                                wSub.Refresh '*KT 53-02-03 แก้ไขตาม POSFront 2.0.33
'                                    DoEvents            '*KT 53-01-14
                            Case Else
                                 '*KT 52-02-23 640 or 800
                                wSub.olaDiscTotal.Caption = SP_FMTcCurPoint(cVB_TotalDiscAft, False, True)
                                wSub.olaNetSale.Caption = SP_FMTcCurPoint(cVB_NetSale, False, True)
                                wSub.olaBalance.Caption = SP_FMTcCurPoint(cVB_Balance, False, True)
                                wSub.Refresh '*KT 53-02-03 แก้ไขตาม POSFront 2.0.33
'                                    DoEvents        '*KT 53-01-14
                        End Select
                 Else
                        Call SP_OBJxSetFocus(otbNum)
                End If
                bW_KeyEnter = False
                If bVB_OPOSMSREnable And bW_EnableMSR Then OPOSMSR1.DataEventEnabled = True          '*Ao 10-02-09
'                If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True
                Exit Sub
             End If
        Case Else
            If cPay > SP_VALcText2Double(SP_FMTcCurPoint((cVB_Balance - cVB_AdjustMalCrdOfOntop + cVB_AdjustDiscMalCrdOfCrd) - cVB_RoundRcv, True, True)) Then  'check amount > balance
                     Call SP_MSGnShowing("tms_CN005066", nCS_Error)
                     If otbNum.Locked Then
                            Call SP_OBJxSetFocus(otbAcc)
                            bVB_IsOnTop = False
                            Select Case SP_VALnText2Double(tVB_TransactionSub)
                                Case tEN_TrnDeposit, tEN_TrnRedeem, tEN_TrnAddDeposit, tEN_TrnDepositCancel, tEN_TrnAddDepositCancel
                                     '*KT 52-02-23 640 or 800
                                    wSub.olaDiscTotal.Caption = SP_FMTcCurPoint(cVB_TotalDiscAft, False, True)
                                    wSub.olaNetSale.Caption = SP_FMTcCurPoint(cVB_Total - cVB_TotalDiscAft, False, True)
                                    wSub.olaBalance.Caption = SP_FMTcCurPoint(cVB_Balance, False, True)
                                    wSub.lblDepositAmt.Caption = SP_FMTcCurPoint(cVB_DepositAmt, False, True)
                                    wSub.Refresh '*KT 53-02-03 แก้ไขตาม POSFront 2.0.33
'                                        DoEvents        '*KT 53-01-14
                                Case Else
                                     '*KT 52-02-23 640 or 800
                                    wSub.olaDiscTotal.Caption = SP_FMTcCurPoint(cVB_TotalDiscAft, False, True)
                                    wSub.olaNetSale.Caption = SP_FMTcCurPoint(cVB_NetSale, False, True)
                                    wSub.olaBalance.Caption = SP_FMTcCurPoint(cVB_Balance, False, True)
                                    wSub.Refresh '*KT 53-02-03 แก้ไขตาม POSFront 2.0.33
'                                        DoEvents        '*KT 53-01-14
                            End Select
                    Else
                            Call SP_OBJxSetFocus(otbNum)
                    End If
                    bW_KeyEnter = False
                    If bVB_OPOSMSREnable And bW_EnableMSR Then OPOSMSR1.DataEventEnabled = True          '*Ao 10-02-09
'                    If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True
                    Exit Sub
             End If
    End Select
    
'    dVB_CrdDateExp = DateSerial(2000 + Val(Format(orbYear.Text)), Val(Format(orbMonth.Text)) + 1, 1)           '*Ao 2013/08/26   เช็คข้างบนแล้ว
''    If Format(dVB_CrdDateExp, "YYYY/MM/DD") <= Format(dVB_CurrentDate, "YYYY/MM/DD") Then '*KT 51-12-09 V 1.0.3 Phase 2 ข้อ 26  เปรียบเทียบน้อยกว่าเพราะdVB_CrdDateExp + ขึ้นไป 1 วันจากที่ระบุ Communication Sheet PH2-ML-SF-027 Credit Expire.xls
'    If (bW_EdcCheck) And (Format(dVB_CrdDateExp, "YYYY/MM/DD") <= Format(dVB_CurrentDate, "YYYY/MM/DD")) Then '*KT 51-12-09 V 1.0.3 Phase 2 ข้อ 26  เปรียบเทียบน้อยกว่าเพราะdVB_CrdDateExp + ขึ้นไป 1 วันจากที่ระบุ Communication Sheet PH2-ML-SF-027 Credit Expire.xls
'        Call SP_MSGnShowing("tms_CN005079", nCS_Error)
'        bW_KeyEnter = False
'        Call SP_OBJxSetFocus(orbMonth)
'        If bVB_OPOSMSREnable Then OPOSMSR1.DataEventEnabled = True          '*Ao 10-02-09
'        If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True
'        Exit Sub
'    End If

'*Eaw 56-09-10 ไม่ต้องเช็ค Approve Code มีหรือไม่มีก็ได้
'    If (bW_EdcCheck) And Trim(otbApvCode.Text) = "" Then 'check approve code            '*Ao 2013/08/27   เพิ่ม bW_EdcCheck
'        Call SP_MSGnShowing("tms_CN005080", nCS_Error)
'        bW_KeyEnter = False
'        Call SP_OBJxSetFocus(otbApvCode)
'        If bVB_OPOSMSREnable Then OPOSMSR1.DataEventEnabled = True          '*Ao 10-02-09
'        If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True
'        Exit Sub
'    End If
    
'    'ทดสอบ อย่าลืมเอาออก
'    If cVB_PayCredit > 10 Then
'        otbRedeem.Text = 10 'Redeem Amount
'    Else
'        otbRedeem.Text = 1 'Redeem Amount
'    End If
'    otbRedeem.Text = cVB_PayCredit 'Full Redeem
'    '==============
    
    cVB_Pay = cPay
    tVB_CrdCardNo = Trim(otbAcc.Text)
'    tVB_CrdExpM = orbMonth.Text                                '*Ao 2013/03/24
    tVB_CrdExpM = Format(orbMonth.Text, "00")
'    tVB_CrdExpY = orbYear.Text
    tVB_CrdExpY = Format(orbYear.Text, "00")
    tVB_CrdTypeName = Trim(otbCardType.Text)
    cVB_CrdAmt = cVB_Pay
    tVB_CrdApprCode = Trim(otbApvCode.Text)
'    tVB_CrdCardType = Trim(ocbType.Text)    '*Eaw 55-11-29
'    tVB_FTRteCode = Trim(ocbType.Text)  '*Eaw 55-11-29 ใช้ FTRteCode เก็บ Type ใน ตาราง TPStSalRC
    tVB_FTRteCode = tVB_CrdCardType  '*Eaw 56-05-15 CommSheet TK-ISS3000-159-IT2 00.00.03.xls
'    If W_CHKbCupsCard(Trim(ocbType.Text)) Then  '*Eaw 56-05-07 CommSheet TK-ISS3000-159-IT2 00.00.02.xls
'        tVB_FTRteCode = "C"
'    End If

    cVB_CreditRedeemAmt = SP_VALcText2Double(otbRedeem.Text) '*Tao 62-07-19 Credit Redeem Amount
    
    oDisc.DISxClearAll
    oDisc.dTranDate = tUT_SaleDate
    oDisc.tTerminalNo = tVB_CNTerminalNum
    oDisc.tTransNo = tVB_Running
    oDisc.tCardID = tVB_CrdCardNo
    oDisc.tCardType = tVB_CrdType
    oDisc.tCardKey = tVB_CrdKey
    oDisc.cDiscValue = cVB_Pay
    oDisc.tStaChkExpire = "Y"
    dVB_CrdDateExp = DateSerial(2000 + Val(Format(orbYear.Text)), Val(Format(orbMonth.Text)), 1)
    oDisc.dCardExpDate = dVB_CrdDateExp
    oDisc.cDiscOntopAlw = cVB_AdjustMalCrdOfOntop
    
    '*KT 54-06-01 V.4.0.0  Req.19 OnTop Specify POSNO
    oDisc.bPOSNoOnTopFlag = bVB_CrdPOSNoOnTopFlag
    oDisc.tPOSNoOnTopAlwType = tVB_CrdPOSNoOnTopAlwType
    oDisc.tFloorAlwOnTop = tVB_CrdFloorAlwOnTop
    oDisc.cSubTotalAll = cUT_SubTotal
    oDisc.cSubTotal = cVB_NetSale
    oDisc.tStaOnTopAlwTierAmtLast = tVB_CrdStaAlwTierAmtLast
    
    'สำหรับให้ทำการ Recal ส่วนลด Mall Card ใหม่ให้เหมือน Special Gift Voucher,Special Coupon
    Select Case SP_VALnText2Double(tVB_TransactionSub)
        Case tEN_TrnRedeem
            oDisc.cBalance = SP_FMTcCurPoint(cVB_DepossitAmtRedeem, False, True)
        Case tEN_TrnAddDeposit
            oDisc.cBalance = SP_FMTcCurPoint(cVB_DepossitAmtBAddDpt, False, True)
    End Select

    If otbNum.Locked Then           '*AO51-03-21  Comm.Sheet  ML-SF-126
        Select Case SP_VALnText2Double(tVB_TransactionSub)
                Case tEN_TrnDeposit, tEN_TrnRedeem, tEN_TrnAddDeposit, tEN_TrnDepositCancel, tEN_TrnAddDepositCancel
                     cVB_AdjustMalCrdOfOntop = SP_FMTcCurPoint(oDisc.DIScGetOntop, True, False)
                     cVB_AdjustDiscMalCrdOfCrd = SP_FMTcCurPoint(oDisc.cAmtAllMall, True, False)
                    '*KT 52-06-04 PH1.8-ML-SF-042
                     If bVB_DepositAll Then
                         cVB_NetSale = cVB_NetSale - cVB_AdjustMalCrdOfOntop + cVB_AdjustDiscMalCrdOfCrd
                    Else
                         cVB_NetSale = cVB_NetSale - cVB_AdjustMalCrdOfOntop
                    End If
'                     cVB_NetSale = cVB_NetSale - cVB_AdjustMalCrdOfOntop
'                     cVB_RoundRcv = IIf(bVB_CNAlwRnd, SP_FMTcCurPoint(UT_GetnRoundDiff(cVB_NetSale - cVB_Pay), True, False), 0)    '*Eaw 56-01-29   แก้ไขการ Rounding สำหรับ TK Project
                     cVB_NetSaleRound = SP_FMTcCurPoint(cVB_NetSale + cVB_RoundRcv, True, False)
                     cVB_TotalDiscAft = SP_FMTcCurPoint(cVB_TotalDiscAft + cVB_AdjustMalCrdOfOntop - cVB_AdjustDiscMalCrdOfCrd, True, False)
                    '*KT 52-06-04 PH1.8-ML-SF-042
                     If bVB_DepositAll Then
                         cVB_DepositAmt = SP_FMTcCurPoint(cVB_DepositAmt - cVB_AdjustMalCrdOfOntop + cVB_AdjustDiscMalCrdOfCrd, True, False)
                    Else
                         cVB_DepositAmt = SP_FMTcCurPoint(cVB_DepositAmt - cVB_AdjustMalCrdOfOntop, True, False)
                    End If
'                     cVB_DepositAmt = SP_FMTcCurPoint(cVB_DepositAmt - cVB_AdjustMalCrdOfOntop, True, False)
                     cVB_Balance = SP_FMTcCurPoint(cVB_NetSaleRound - CDbl(Val(Format(wSub.olaAmtTender.Caption))), True, False)  '*KT 52-02-23 640 or 800
                      '*KT 52-02-23 640 or 800
                     wSub.olaDiscTotal.Caption = SP_FMTcCurPoint(cVB_TotalDiscAft, False, True)
                     wSub.olaNetSale.Caption = SP_FMTcCurPoint(cVB_Total - cVB_TotalDiscAft, False, True)
                     wSub.lblDepositAmt.Caption = SP_FMTcCurPoint(cVB_DepositAmt, False, True)
                     wSub.olaBalance.Caption = SP_FMTcCurPoint(cVB_Balance, False, True)
                     wSub.Refresh '*KT 53-02-03 แก้ไขตาม POSFront 2.0.33
'                     DoEvents           '*KT 53-01-14
                     
                    Select Case SP_VALnText2Double(tVB_TransactionSub)
                        Case tEN_TrnDeposit
                            tSql = "UPDATE " & tVB_CNTblHD
                            tSql = tSql & " SET FCShdDepositAmt=" & cVB_DepositAmt
                            '*KT 52-06-04 PH1.8-ML-SF-042
                            If Not bVB_DepositAll Then
                                tSql = tSql & ",FCShdDepositBal = FCShdDepositBal + " & cVB_AdjustDiscMalCrdOfCrd
                            End If
'                            tSql = tSql & ",FCShdDepositBal = FCShdDepositBal + " & cVB_AdjustDiscMalCrdOfCrd
                            tSql = tSql & ",FCShdGrand =" & (cVB_Total - cVB_TotalDiscAft)
                            tSql = tSql & ",FCShdRnd=" & cVB_RoundRcv
                            tSql = tSql & "," & UT_tLastUpdSql
                            tSql = tSql & "  WHERE FTTmnNum='" & tVB_CNTerminalNum & "'"
                            tSql = tSql & "  AND FTShdTransNo='" & tVB_Running & "'"
                            tSql = tSql & "  AND FDShdTransDate ='" & tUT_SaleDate & "'"
                            tSql = tSql & "  AND FTShdTransType='" & tVB_Transaction & "'"
                            tSql = tSql & "  AND FTShdStaDoc = '2'"
                            Call SP_SQLbExecute(tSql)
                        Case tEN_TrnAddDeposit
                            cVB_DepossitAmtAddDpt = cVB_DepositAmt
                            tSql = "UPDATE " & tVB_CNTblHD
                            tSql = tSql & " SET FCShdDepositAmt=" & cVB_DepossitAmtAddDpt
                            '*KT 52-06-04 PH1.8-ML-SF-042
                            If Not bVB_DepositAll Then
                                tSql = tSql & " ,FCShdDepositBal=FCShdDepositBal + " & cVB_AdjustDiscMalCrdOfCrd '" & SP_FMTcCurPoint(cVB_AccBalAddDpt - cVB_DepossitAmtAddDpt, True, False)
                            End If
'                            tSql = tSql & " ,FCShdDepositBal=FCShdDepositBal + " & cVB_AdjustDiscMalCrdOfCrd '" & SP_FMTcCurPoint(cVB_AccBalAddDpt - cVB_DepossitAmtAddDpt, True, False)
                            tSql = tSql & "," & UT_tLastUpdSql
                            tSql = tSql & "  WHERE FTTmnNum='" & tVB_CNTerminalNum & "'"
                            tSql = tSql & "  AND FTShdTransNo='" & tVB_Running & "'"
                            tSql = tSql & "  AND FDShdTransDate ='" & tUT_SaleDate & "'"
                            tSql = tSql & "  AND FTShdTransType='" & tVB_Transaction & "'"
                            tSql = tSql & "  AND FTShdStaDoc = '2'"
                            Call SP_SQLbExecute(tSql)
                    End Select
                Case tEN_TrnSaleManualTax, tEN_TrnReturnManualTax
                    'none
                Case Else
                    cVB_AdjustMalCrdOfOntop = SP_FMTcCurPoint(oDisc.DIScGetOntop, True, True)
                    cVB_AdjustDiscMalCrdOfCrd = SP_FMTcCurPoint(oDisc.cAmtAllMall, True, True)
                    
                    cVB_NetSale = cVB_NetSale - cVB_AdjustMalCrdOfOntop + cVB_AdjustDiscMalCrdOfCrd
'                    cVB_RoundRcv = IIf(bVB_CNAlwRnd, SP_FMTcCurPoint(UT_GetnRoundDiff(cVB_NetSale + cVB_NetDutyFreeDep - cVB_Pay), True, False), 0)   '*Eaw 56-01-29   แก้ไขการ Rounding สำหรับ TK Project
                    cVB_NetSaleRound = SP_FMTcCurPoint(cVB_NetSale + cVB_RoundRcv, True, False)
                    cVB_TotalDiscAft = SP_FMTcCurPoint(cVB_TotalDiscAft + cVB_AdjustMalCrdOfOntop - cVB_AdjustDiscMalCrdOfCrd, True, False)
                    cVB_Balance = cVB_NetSaleRound - CDbl(Val(Format(wSub.olaAmtTender.Caption)))  '*KT 52-02-23 640 or 800
                    '*KT 52-02-23 640 or 800
                    wSub.olaDiscTotal.Caption = SP_FMTcCurPoint(cVB_TotalDiscAft, False, True)
                    wSub.olaNetSale.Caption = SP_FMTcCurPoint(cVB_NetSale, False, True)
                    wSub.olaBalance.Caption = SP_FMTcCurPoint(cVB_Balance, False, True)
                    wSub.Refresh '*KT 53-02-03 แก้ไขตาม POSFront 2.0.33
'                    DoEvents            '*KT 53-01-14
        End Select
    Else
        
        '*KT 51-08-16 ทำการคำนวณค่าใหม่กรณีที่มีการชำระด้วยบัตรเครดิตและมีการเฟกหรือไม่เฟก Mall Card ที่ Credit Card
        '----------------------------------------------------------------------------------------------------
            Select Case Val(Format(tVB_TransactionSub))
                Case tEN_TrnDeposit, tEN_TrnRedeem, tEN_TrnAddDeposit, tEN_TrnDepositCancel, tEN_TrnAddDepositCancel
                     cVB_AdjustMalCrdOfOntop = 0
                     cVB_AdjustDiscMalCrdOfCrd = SP_FMTcCurPoint(oDisc.DIScGetCredit, True, True)
                    '*KT 52-06-04 PH1.8-ML-SF-042
                     If bVB_DepositAll Then
                         cVB_NetSale = cVB_NetSale - cVB_AdjustMalCrdOfOntop + cVB_AdjustDiscMalCrdOfCrd
                    Else
                         cVB_NetSale = cVB_NetSale - cVB_AdjustMalCrdOfOntop
                    End If
'                     cVB_NetSale = cVB_NetSale - cVB_AdjustMalCrdOfOntop
'                     cVB_RoundRcv = IIf(bVB_CNAlwRnd, SP_FMTcCurPoint(UT_GetnRoundDiff(cVB_NetSale - cVB_Pay), True, False), 0)     '*Eaw 56-01-29   แก้ไขการ Rounding สำหรับ TK Project
                     cVB_NetSaleRound = SP_FMTcCurPoint(cVB_NetSale + cVB_RoundRcv, True, False)
                     cVB_TotalDiscAft = SP_FMTcCurPoint(cVB_TotalDiscAft + cVB_AdjustMalCrdOfOntop - cVB_AdjustDiscMalCrdOfCrd, True, False)
                    '*KT 52-06-04 PH1.8-ML-SF-042
                     If bVB_DepositAll Then
                         cVB_DepositAmt = SP_FMTcCurPoint(cVB_DepositAmt - cVB_AdjustMalCrdOfOntop + cVB_AdjustDiscMalCrdOfCrd, True, False)
                    Else
                         cVB_DepositAmt = SP_FMTcCurPoint(cVB_DepositAmt - cVB_AdjustMalCrdOfOntop, True, False)
                    End If
'                     cVB_DepositAmt = SP_FMTcCurPoint(cVB_DepositAmt - cVB_AdjustMalCrdOfOntop, True, False)
                     cVB_Balance = SP_FMTcCurPoint(cVB_NetSaleRound - CDbl(Val(Format(wSub.olaAmtTender.Caption))), True, False)  '*KT 52-02-23 640 or 800
                      '*KT 52-02-23 640 or 800
                     wSub.olaDiscTotal.Caption = SP_FMTcCurPoint(cVB_TotalDiscAft, False, True)
                     wSub.olaNetSale.Caption = SP_FMTcCurPoint(cVB_NetSale, False, True)
                     wSub.lblDepositAmt.Caption = SP_FMTcCurPoint(cVB_DepositAmt, False, True)
                     wSub.olaBalance.Caption = SP_FMTcCurPoint(cVB_Balance, False, True)
                     wSub.Refresh '*KT 53-02-03 แก้ไขตาม POSFront 2.0.33
'                     DoEvents               '*KT 53-01-14
                     
                     Select Case Val(Format(tVB_TransactionSub))
                        Case tEN_TrnDeposit
                            tSql = "UPDATE " & tVB_CNTblHD
                            tSql = tSql & " SET FCShdDepositAmt=" & cVB_DepositAmt
                            '*KT 52-06-04 PH1.8-ML-SF-042
                            If Not bVB_DepositAll Then
                                tSql = tSql & ",FCShdDepositBal = FCShdDepositBal + " & cVB_AdjustDiscMalCrdOfCrd
                            End If
'                            tSql = tSql & ",FCShdDepositBal = FCShdDepositBal + " & cVB_AdjustDiscMalCrdOfCrd
                            tSql = tSql & "," & UT_tLastUpdSql
                            tSql = tSql & "  WHERE FTTmnNum='" & tVB_CNTerminalNum & "'"
                            tSql = tSql & "  AND FTShdTransNo='" & tVB_Running & "'"
                            tSql = tSql & "  AND FDShdTransDate ='" & tUT_SaleDate & "'"
                            tSql = tSql & "  AND FTShdTransType='" & tVB_Transaction & "'"
                            tSql = tSql & "  AND FTShdStaDoc = '2'"
                            Call SP_SQLbExecute(tSql)
                        Case tEN_TrnAddDeposit
                            cVB_DepossitAmtAddDpt = cVB_DepositAmt
                            tSql = "UPDATE " & tVB_CNTblHD
                            tSql = tSql & " SET FCShdDepositAmt=" & cVB_DepossitAmtAddDpt
                            '*KT 52-06-04 PH1.8-ML-SF-042
                            If Not bVB_DepositAll Then
                                tSql = tSql & " ,FCShdDepositBal=FCShdDepositBal + " & cVB_AdjustDiscMalCrdOfCrd '" & SP_FMTcCurPoint(cVB_AccBalAddDpt - cVB_DepossitAmtAddDpt, True, False)
                            End If
'                            tSql = tSql & " ,FCShdDepositBal=FCShdDepositBal + " & cVB_AdjustDiscMalCrdOfCrd '" & SP_FMTcCurPoint(cVB_AccBalAddDpt - cVB_DepossitAmtAddDpt, True, False)
                            tSql = tSql & "," & UT_tLastUpdSql
                            tSql = tSql & "  WHERE FTTmnNum='" & tVB_CNTerminalNum & "'"
                            tSql = tSql & "  AND FTShdTransNo='" & tVB_Running & "'"
                            tSql = tSql & "  AND FDShdTransDate ='" & tUT_SaleDate & "'"
                            tSql = tSql & "  AND FTShdTransType='" & tVB_Transaction & "'"
                            tSql = tSql & "  AND FTShdStaDoc = '2'"
                            Call SP_SQLbExecute(tSql)
                    End Select
                    
                Case Else
        
                    cVB_AdjustMalCrdOfOntop = 0
                    cVB_AdjustDiscMalCrdOfCrd = SP_FMTcCurPoint(oDisc.DIScGetCredit, True, True)
                    
                    cVB_NetSale = SP_FMTcCurPoint(cVB_NetSale - cVB_AdjustMalCrdOfOntop + cVB_AdjustDiscMalCrdOfCrd, True, False)
'                    cVB_RoundRcv = IIf(bVB_CNAlwRnd, SP_FMTcCurPoint(UT_GetnRoundDiff(cVB_NetSale + cVB_NetDutyFreeDep - cVB_Pay), True, False), 0)   '*Eaw 56-01-29   แก้ไขการ Rounding สำหรับ TK Project
                    cVB_NetSaleRound = SP_FMTcCurPoint(cVB_NetSale + cVB_RoundRcv, True, False)
                    cVB_TotalDiscAft = SP_FMTcCurPoint(cVB_TotalDiscAft + cVB_AdjustMalCrdOfOntop - cVB_AdjustDiscMalCrdOfCrd, True, False)
                    cVB_Balance = SP_FMTcCurPoint(cVB_NetSaleRound - CDbl(Val(Format(wSub.olaAmtTender.Caption))), True, False)  '*KT 52-02-23 640 or 800
                    '*KT 52-02-23 640 or 800
                    wSub.olaDiscTotal.Caption = SP_FMTcCurPoint(cVB_TotalDiscAft, False, True)
                    wSub.olaNetSale.Caption = SP_FMTcCurPoint(cVB_NetSale, False, True)
                    wSub.olaBalance.Caption = SP_FMTcCurPoint(cVB_Balance, False, True)
                    wSub.Refresh '*KT 53-02-03 แก้ไขตาม POSFront 2.0.33
'                    DoEvents            '*KT 53-01-14
            End Select
        '---------------------------------------------------------------------------------------------------
    End If
    
    '*Korn 55-10-17 Edit ถ้าเป็นการรูดผ่าน EDC จะ hold หน้าจอไว้สัก 1-2 วินาที
    If (Not bW_EdcCheck) And (Not bW_C209) Then
        ofaSwipeCard.Left = 2900      'ซ่อน
        ofaSwipeCard.Top = 5400
        DoEvents
        SleepEx 1000, 0     '2 วินาที     '*Ao 2013/011/13   ผู้ใช้ไม่อยากรอนาน ปรับเหลือ 1 วิ
    End If
    If bW_C209 = True Then '*Bump 58-05-02
        If otbApvCode.Text = "" Then
            MsgBox "Please entry approve code", tCS_CNMsgWarn, tVB_CNAppTitle
            otcForm1(1).Enabled = True    'lock text box
            otcForm1(2).Enabled = True     'lock ปุ่ม
            Call SP_OBJxSetFocus(otbApvCode)
            ocmCancel.Enabled = False
            ocmEDC.Enabled = False
            ocbType.Enabled = False
            Exit Sub
        End If
    End If
    Unload Me
    Exit Sub
ErrHandle:
    bW_FoundAcc = False
    bW_KeyEnter = False
    Call SP_MSGxShowError(Err.Number, Err.Description)
    Call SP_OBJxSetFocus(otbAcc)
    If bVB_OPOSMSREnable And bW_EnableMSR Then OPOSMSR1.DataEventEnabled = True          '*Ao 10-02-09
'    If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True
End Sub



Private Sub OPOSMSR1_DataEvent(ByVal Status As Long)
    Dim tExpDate$
    Dim tWhe$
    Dim cDiscOnTopAmt As Double
    Dim cNetSale  As Double
    Dim cRoundRcv As Double
    Dim cNetSaleRound As Double
    Dim cTotalDiscAft As Double
    Dim cBalance As Double
    Dim bTime5 As Boolean
    Dim nI As Long
    Dim tTrack1Data$ '*Bump 56-04-02 เพิ่ม Code การเช็ค TC Card Comsheet TK-ISS3000-023
    Dim tTrack2Data$
    Dim tCheckCardRang As String
    Dim tCardID As String
    Dim nPos As Long
On Error GoTo ErrHandle '*KT 52-02-11  PH2-ML-SF-092 โปรแกรมหลุดมาที่ Desktop ตอนพิมพ์ receipt , ชำระด้วย Credit

    If tVB_CNMnuCurrent = "1831" Then bW_ManaulMode = True '* Bump 56-04-11 TC Card
    
    If bW_ManaulMode = True Then
'            With oScanner                                                           '*Ao 10-02-09
'                  If bVB_OPOSScannerEnable Then
'                         'Enable the event.
'                         .DataEventEnabled = False
'                  End If
'             End With
            With OPOSMSR1
                otbAcc.Text = .AccountNumber     '(Account No.)
                tExpDate = .ExpirationDate    '(Valid date)
                orbYear.Text = Mid(tExpDate, 1, 2)
                orbMonth.Text = Mid(tExpDate, 3, 2)
                '*Bump 56-04-02 เพิ่ม Code การเช็ค TC Card Comsheet TK-ISS3000-023
                tTrack1Data = .Track1Data '-->"s60000120345389706440548160000000000000040200000000000000000000000060"
                tTrack2Data = .Track2Data   '-->"4538970644054816=04021011387700000947"
                Call SP_CRTxLogMallCardEvent("MSR/PayTCCard", "MSR", "1831:,1817:Credit:", tVB_CNMnuCurrent & ":")
                Call SP_CRTxLogMallCardEvent("MSR/PayTCCard", "MSR", "Track1Data:", tTrack1Data & ":")
                Call SP_CRTxLogMallCardEvent("MSR/PayTCCard", "MSR", "Track2Data:", tTrack2Data & ":")
                'Call SP_CRTxLogMallCardEvent("MSR/PayTCCard", "MSR", "AccountNumber:", .AccountNumber & ":") '*Tao 61-09-18 ของเดิม
                Call SP_CRTxLogMallCardEvent("MSR/PayTCCard", "MSR", "AccountNumber:", SP_PRCtStrMarkX(.AccountNumber) & ":") '*Tao 61-09-18 Mark X ก่อน
                Call SP_CRTxLogMallCardEvent("MSR/PayTCCard", "MSR", "ExpirationDate:", .ExpirationDate & ":")
                
                If tVB_CNMnuCurrent = "1831" Then 'TC Card
'                    If Trim(tTrack1Data) <> "" And Trim(tTrack2Data) <> "" Then
'                        nPos = InStrRev(tTrack2Data, "=")
'                        If nPos > 1 Then
'                            tCardID = Trim(Mid(tTrack2Data, 1, nPos - 1))  '-->"4538970644054816" คือ หมายเลขบนบัตร TC Card
'                            nPos = InStrRev(tTrack1Data, tCardID)
'                                If nPos > 1 Then
'                                    tCheckCardRang = Mid(tTrack1Data, (nPos - 10), 26)
'                                    If Len(tCheckCardRang) < 26 Then tCheckCardRang = tCardID
'                                    'otbAcc.Text = .AccountNumber
'                                     'tExpDate = .ExpirationDate    '(Valid date)
'                                     'orbYear.Text = Mid(tExpDate, 1, 2)
'                                     'orbMonth.Text = Mid(tExpDate, 3, 2)
'                                Else
'                                     tCheckCardRang = tCardID
'                                End If
'                                Call SP_CRTxLogMallCardEvent("MSR/PayTCCard", "MSR", "CardID:", tCardID & ":")
'                                Call SP_CRTxLogMallCardEvent("MSR/PayTCCard", "MSR", "IDMark->Validaty Code:", tCheckCardRang)
'                        End If
'                   End If
'                    tTrack1Data = "s60000120345389706440548160000000000000140200000000000000000000000060"
'                    tTrack2Data = "4538970644054816=14021011387700000947"
'                    tTrack1Data = "s1000000353777842211180100000000000061170300000000000000000000000000"
'                    tTrack2Data = "377784221118010=1212"
'
'                    tTrack1Data = "s1000000353777841891330120000000000061170300000000000000000000000000"
'                    tTrack2Data = "377784189133012=1212"
                    If (tTrack1Data <> "") Or (tTrack2Data <> "") Then              '*Ao 2013/05/28   TC Card
                        If tTrack1Data <> "" Then
                            tCheckCardRang = Mid(tTrack1Data, 11, 16)
'                            If Right(tCheckCardRang, 1) = "0" Then tCheckCardRang = Left(tCheckCardRang, 15) '*Eaw 56-07-10 CommSheet TK-ISS3000-272-TR.xls
                            If Not W_PRCbFindTCCardByNum(tCheckCardRang) Then tCheckCardRang = Left(tCheckCardRang, 15) '*Eaw 56-09-04
                            otbAcc.Text = tCheckCardRang
                            tExpDate = Mid(tTrack1Data, 40, 4)
                        Else
                            If tTrack2Data <> "" Then
                                If Len(tTrack2Data) > 20 Then
                                    nPos = InStrRev(tTrack2Data, "=")
                                    If nPos > 1 Then
                                        tCheckCardRang = Trim(Mid(tTrack2Data, 1, InStr(1, tTrack2Data, "=") - 1))
                                        otbAcc.Text = tCheckCardRang
                                        tExpDate = Trim(Mid(tTrack2Data, InStr(1, tTrack2Data, "=") + 1, Len(tTrack2Data) - InStr(1, tTrack2Data, "=")))
                                    End If
                                Else
                                    tCheckCardRang = tTrack2Data
                                    otbAcc.Text = tCheckCardRang
                                End If
                            End If
                        End If
                    End If
                    If tCheckCardRang <> "" Then
                        If tExpDate <> "" Then
                            orbYear.Text = Mid(tExpDate, 1, 2)
                            orbMonth.Text = Mid(tExpDate, 3, 2)
                        End If
                    End If
                End If
                .ClearInput
                .DataEventEnabled = True
            End With
            
            bW_ChkCrd = True
            
            If tVB_CNMnuCurrent = "1831" Then '*Bump 56-04-02 เพิ่ม Code การเช็ค TC Card Comsheet TK-ISS3000-023
                tVB_CrdCardNo = tCheckCardRang
            Else
                tVB_CrdCardNo = Trim(otbAcc.Text)
            End If
            'ตรวจสอบความถูกต้องของเลขที่
            If Not W_PRCxFindCrdByAccNum(tVB_CrdCardNo) Then
'                Call SP_MSGnShowing("tms_CN005065", nCS_Error)
                Call W_CHKxCaseERROR(nW_CaseERROR)                      '*Ao 2013/02/22   CommSheet TK-ISS3000-011
                 bW_FoundAcc = False
                 otbAcc.Text = ""
                 orbMonth.Text = ""
                 orbYear.Text = ""
                 otbCardType.Text = ""
                 otbApvCode.Text = ""
                 otbNum.Text = ""
                 Call SP_OTBxSelText(otbAcc)
'                 If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True          '*Ao 10-02-09
                 Exit Sub
            Else
                dVB_CrdDateExp = DateSerial(2000 + Val(Format(orbYear.Text)), Val(Format(orbMonth.Text)) + 1, 1)
                If Format(dVB_CrdDateExp, "YYYY/MM/DD") <= Format(dVB_CurrentDate, "YYYY/MM/DD") Then '*KT 51-12-09 V 1.0.3 Phase 2 ข้อ 26  เปรียบเทียบน้อยกว่าเพราะdVB_CrdDateExp + ขึ้นไป 1 วันจากที่ระบุ Communication Sheet PH2-ML-SF-027 Credit Expire.xls
                    Call SP_MSGnShowing("tms_CN005143", nCS_Warning)
                    otbAcc.Text = ""
                    orbMonth.Text = ""
                    orbYear.Text = ""
                    otbCardType.Text = ""
                    otbApvCode.Text = ""
                    otbNum.Text = ""
                    Call SP_OTBxSelText(otbAcc)
'                    If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True          '*Ao 10-02-09
                    Exit Sub
                End If
                bW_FoundAcc = True
            End If
'            ocbType.AddItem tVB_CrdCardType, 1 '*Eaw 56-09-03 CommSheet TK-ISS3000-341
'            ocbType.Text = tVB_CrdCardType                  '*Ao 2013/03/07     CommSheet TK-ISS3000-012-00.00.05.xlsx
            With oDisc
                 If nVB_OnTopCount >= 5 Then
                    Call SP_MSGnShowing("tms_CN005154", nCS_Warning)
                    bVB_IsOnTop = False
                Else
                    oDisc.DISxClearAll
                    oDisc.dTranDate = tUT_SaleDate
                    oDisc.tTerminalNo = tVB_CNTerminalNum
                    oDisc.tTransNo = tVB_Running
                    oDisc.tCardID = tVB_CrdCardNo
                    oDisc.tCardType = tVB_CrdType
                    oDisc.tCardKey = tVB_CrdKey
                    oDisc.tStaChkExpire = "Y"
                    oDisc.cDiscValue = SP_VALcText2Double(wSub.olaBalance.Caption)  '*KT 52-02-23 640 or 800
                    dVB_CrdDateExp = DateSerial(2000 + Val(Format(orbYear.Text)), Val(Format(orbMonth.Text)), 1)
                    oDisc.dCardExpDate = dVB_CrdDateExp
                    
                    '*KT 54-06-01 V.4.0.0  Req.19 OnTop Specify POSNO
                    oDisc.bPOSNoOnTopFlag = bVB_CrdPOSNoOnTopFlag
                    oDisc.tPOSNoOnTopAlwType = tVB_CrdPOSNoOnTopAlwType
                    oDisc.tFloorAlwOnTop = tVB_CrdFloorAlwOnTop
                    oDisc.cSubTotalAll = cUT_SubTotal
                    oDisc.cSubTotal = cVB_NetSale
                    oDisc.tStaOnTopAlwTierAmtLast = tVB_CrdStaAlwTierAmtLast
                    
                    'สำหรับให้ทำการ Recal ส่วนลด Mall Card ใหม่ให้เหมือน Special Gift Voucher,Special Coupon
                    Select Case SP_VALnText2Double(tVB_TransactionSub)
                        Case tEN_TrnRedeem
                            oDisc.cBalance = SP_FMTcCurPoint(cVB_DepossitAmtRedeem, False, True)
                        Case tEN_TrnAddDeposit
                            oDisc.cBalance = SP_FMTcCurPoint(cVB_DepossitAmtBAddDpt, False, True)
                    End Select
                    
        '            DoEvents            '*KT 53-01-14
                    cVB_CNOntopMaxDisc = SP_FMTcCurPoint(oDisc.DIScCalOntop, True, False) 'Ontop ที่ได้
                    cVB_CNOntopMaxCal = SP_FMTcCurPoint(oDisc.cDiscValue, True, False) 'ยอดเงินที่หัก Ontop แล้ว
                    cVB_CNOntopMaxAlw = SP_FMTcCurPoint(oDisc.cDiscOntopAlw, True, False) 'ยอดเต็ม '*KT 51-06-20 Comsheet ML-SF-270
                    If cVB_CNOntopMaxDisc > 0 Then
                         Screen.MousePointer = 0
                        'Communication Sheet ML-SF-463 Must not use same credit card number within 5 ontop.xls
                        '--------------------------------------------------------------------------------------------------------------------
                        bTime5 = True
                        For nI = LBound(aVB_CreditCrd) To UBound(aVB_CreditCrd) - 1
                            If tVB_CrdCardNo = aVB_CreditCrd(nI) Then
                                bTime5 = False
                            End If
                        Next nI
                        '----------------------------------------------------------------------------------------------------------------------
                         If bTime5 Then
                            bVB_IsOnTop = True
                            wCNPayCreditAmt.Show vbModal
                            If cVB_CrdAmt > 0 And bVB_IsCrdCard Then
                                otbNum.Locked = True
                                otbNum.BackColor = nVB_CNClrBLock
                                otbNum.Text = SP_FMTcCurPoint(cVB_CrdAmt, False, True)
                                Call SP_OBJxSetFocus(otbApvCode)
                                
                                'กรณีที่รูดเครดิตแล้วได้ Ontop
                                '---------------------------------------------------------------------------------------
                                Select Case Val(Format(tVB_TransactionSub))
                                   Case tEN_TrnDeposit, tEN_TrnRedeem, tEN_TrnAddDeposit, tEN_TrnDepositCancel, tEN_TrnAddDepositCancel
                                        cVB_AdjustMalCrdOfOntop = SP_FMTcCurPoint(cVB_CNOntopMaxDisc, True, False)
                                        cVB_AdjustDiscMalCrdOfCrd = SP_FMTcCurPoint(oDisc.cAmtAllMall, True, False)
                                        
                                        cNetSale = SP_FMTcCurPoint(cVB_NetSale - cVB_AdjustMalCrdOfOntop, True, False)
                                        cTotalDiscAft = SP_FMTcCurPoint(cVB_TotalDiscAft + cVB_AdjustMalCrdOfOntop - cVB_AdjustDiscMalCrdOfCrd, True, False)
                                        cBalance = SP_FMTcCurPoint(cVB_Balance - cVB_RoundRcv - cVB_AdjustMalCrdOfOntop, True, False)
                                         '*KT 52-02-23 640 or 800
                                        wSub.olaDiscTotal.Caption = SP_FMTcCurPoint(cTotalDiscAft, False, True)
                                        wSub.olaNetSale.Caption = SP_FMTcCurPoint(cVB_Total - cTotalDiscAft, False, True)
                                        wSub.lblDepositAmt.Caption = SP_FMTcCurPoint(cVB_DepositAmt - cVB_AdjustMalCrdOfOntop, False, True)
                                        wSub.olaBalance.Caption = SP_FMTcCurPoint(cBalance, False, True)
                                        wSub.Refresh '*KT 53-02-03 แก้ไขตาม POSFront 2.0.33
        '                                DoEvents            '*KT 53-01-14
                                   Case Else
                                        cVB_AdjustMalCrdOfOntop = SP_FMTcCurPoint(cVB_CNOntopMaxDisc, True, False)
                                        cVB_AdjustDiscMalCrdOfCrd = SP_FMTcCurPoint(oDisc.cAmtAllMall, True, False)
                                        
                                        cNetSale = SP_FMTcCurPoint(cVB_NetSale - cVB_AdjustMalCrdOfOntop + cVB_AdjustDiscMalCrdOfCrd, True, False)
'                                        cRoundRcv = IIf(bVB_CNAlwRnd, SP_FMTcCurPoint(UT_GetnRoundDiff(cNetSale + cVB_NetDutyFreeDep), True, False), 0)  '*Eaw 56-01-29   แก้ไขการ Rounding สำหรับ TK Project
                                        cNetSaleRound = SP_FMTcCurPoint(cNetSale + cRoundRcv, True, False)
                                        cTotalDiscAft = SP_FMTcCurPoint(cVB_TotalDiscAft + cVB_AdjustMalCrdOfOntop - cVB_AdjustDiscMalCrdOfCrd, True, False)
                                        cBalance = SP_FMTcCurPoint(cVB_Balance - cVB_AdjustMalCrdOfOntop + cVB_AdjustDiscMalCrdOfCrd - cVB_RoundRcv, True, False)
                                         '*KT 52-02-23 640 or 800
                                        wSub.olaDiscTotal.Caption = SP_FMTcCurPoint(cTotalDiscAft, False, True)
                                        wSub.olaNetSale.Caption = SP_FMTcCurPoint(cNetSale, False, True)
                                        wSub.olaBalance.Caption = SP_FMTcCurPoint(cBalance, False, True)
                                        wSub.Refresh '*KT 53-02-03 แก้ไขตาม POSFront 2.0.33
        '                                DoEvents            '*KT 53-01-14
                                End Select
                                '--------------------------------------------------------------------------------------
                            Else
                                otbNum.Text = ""
                                otbNum.Locked = False
                                otbNum.BackColor = nVB_CNClrBUnLock
                                'กรณีที่รูดเครดิตแล้วได้ Ontop แล้วยกเลิก
                                bVB_IsOnTop = False
                                '*KT 51-08-16 ทำการคำนวณค่าใหม่กรณีที่มีการชำระด้วยบัตรเครดิตและมีการเฟกหรือไม่เฟก Mall Card ที่ Credit Card
                                '----------------------------------------------------------------------------------------------------
                                Select Case Val(Format(tVB_TransactionSub))
                                   Case tEN_TrnDeposit, tEN_TrnRedeem, tEN_TrnAddDeposit, tEN_TrnDepositCancel, tEN_TrnAddDepositCancel
                                        cVB_AdjustMalCrdOfOntop = 0
                                        cVB_AdjustDiscMalCrdOfCrd = SP_FMTcCurPoint(oDisc.DIScCalCredit, True, False)
                                        
                                        cNetSale = SP_FMTcCurPoint(cVB_NetSale - cVB_AdjustMalCrdOfOntop, True, False)
                                        cTotalDiscAft = SP_FMTcCurPoint(cVB_TotalDiscAft + cVB_AdjustMalCrdOfOntop - cVB_AdjustDiscMalCrdOfCrd, True, False)
                                        cBalance = SP_FMTcCurPoint(cVB_Balance - cVB_RoundRcv - cVB_AdjustMalCrdOfOntop, True, False)
                                         '*KT 52-02-23 640 or 800
                                        wSub.olaDiscTotal.Caption = SP_FMTcCurPoint(cTotalDiscAft, False, True)
                                        wSub.olaNetSale.Caption = SP_FMTcCurPoint(cVB_Total - cTotalDiscAft, False, True)
                                        wSub.lblDepositAmt.Caption = SP_FMTcCurPoint(cVB_DepositAmt - cVB_AdjustMalCrdOfOntop, False, True)
                                        wSub.olaBalance.Caption = SP_FMTcCurPoint(cBalance, False, True)
                                        wSub.Refresh '*KT 53-02-03 แก้ไขตาม POSFront 2.0.33
        '                                DoEvents                '*KT 53-01-14
                                        
                                    Case Else
                                        cVB_AdjustMalCrdOfOntop = 0
                                        cVB_AdjustDiscMalCrdOfCrd = SP_FMTcCurPoint(oDisc.DIScCalCredit, True, False)
                                        
                                        cNetSale = cVB_NetSale - cVB_AdjustMalCrdOfOntop + cVB_AdjustDiscMalCrdOfCrd
'                                        cRoundRcv = IIf(bVB_CNAlwRnd, SP_FMTcCurPoint(UT_GetnRoundDiff(cNetSale + cVB_NetDutyFreeDep), True, False), 0)  '*Eaw 56-01-29   แก้ไขการ Rounding สำหรับ TK Project
                                        cNetSaleRound = SP_FMTcCurPoint(cNetSale + cRoundRcv, True, False)
                                        cTotalDiscAft = SP_FMTcCurPoint(cVB_TotalDiscAft + cVB_AdjustMalCrdOfOntop - cVB_AdjustDiscMalCrdOfCrd, True, False)
                                        cBalance = SP_FMTcCurPoint(cVB_Balance - cVB_AdjustMalCrdOfOntop + cVB_AdjustDiscMalCrdOfCrd - cVB_RoundRcv, True, False)
                                         '*KT 52-02-23 640 or 800
                                        wSub.olaDiscTotal.Caption = SP_FMTcCurPoint(cTotalDiscAft, False, True)
                                        wSub.olaNetSale.Caption = SP_FMTcCurPoint(cNetSale, False, True)
                                        wSub.olaBalance.Caption = SP_FMTcCurPoint(cBalance, False, True)
                                        wSub.Refresh '*KT 53-02-03 แก้ไขตาม POSFront 2.0.33
        '                                DoEvents                '*KT 53-01-14
                                End Select
                                '---------------------------------------------------------------------------------------------------------
                            End If
                        Else
                            otbNum.Text = ""
                            otbNum.Locked = False
                            otbNum.BackColor = nVB_CNClrBUnLock
                            'กรณีที่รูดเครดิตแล้วได้ Ontop มาแล้ว 5 ครั้ง
                            bVB_IsOnTop = False
                            '*KT 51-08-16 ทำการคำนวณค่าใหม่กรณีที่มีการชำระด้วยบัตรเครดิตและมีการเฟกหรือไม่เฟก Mall Card ที่ Credit Card
                            '----------------------------------------------------------------------------------------------------
                            Select Case Val(Format(tVB_TransactionSub))
                                   Case tEN_TrnDeposit, tEN_TrnRedeem, tEN_TrnAddDeposit, tEN_TrnDepositCancel, tEN_TrnAddDepositCancel
                                        cVB_AdjustMalCrdOfOntop = 0
                                        cVB_AdjustDiscMalCrdOfCrd = SP_FMTcCurPoint(oDisc.DIScCalCredit, True, False)
                                        
                                        cNetSale = SP_FMTcCurPoint(cVB_NetSale - cVB_AdjustMalCrdOfOntop, True, False)
                                        cTotalDiscAft = SP_FMTcCurPoint(cVB_TotalDiscAft + cVB_AdjustMalCrdOfOntop - cVB_AdjustDiscMalCrdOfCrd, True, False)
                                        cBalance = SP_FMTcCurPoint(cVB_Balance - cVB_RoundRcv - cVB_AdjustMalCrdOfOntop, True, False)
                                         '*KT 52-02-23 640 or 800
                                        wSub.olaDiscTotal.Caption = SP_FMTcCurPoint(cTotalDiscAft, False, True)
                                        wSub.olaNetSale.Caption = SP_FMTcCurPoint(cVB_Total - cTotalDiscAft, False, True)
                                        wSub.lblDepositAmt.Caption = SP_FMTcCurPoint(cVB_DepositAmt - cVB_AdjustMalCrdOfOntop, False, True)
                                        wSub.olaBalance.Caption = SP_FMTcCurPoint(cBalance, False, True)
                                        wSub.Refresh '*KT 53-02-03 แก้ไขตาม POSFront 2.0.33
        '                                DoEvents                '*KT 53-01-14
                                    Case Else
                                        cVB_AdjustMalCrdOfOntop = 0
                                        cVB_AdjustDiscMalCrdOfCrd = SP_FMTcCurPoint(oDisc.DIScCalCredit, True, False)
                                        
                                        cNetSale = SP_FMTcCurPoint(cVB_NetSale - cVB_AdjustMalCrdOfOntop + cVB_AdjustDiscMalCrdOfCrd, True, False)
'                                        cRoundRcv = IIf(bVB_CNAlwRnd, SP_FMTcCurPoint(UT_GetnRoundDiff(cNetSale + cVB_NetDutyFreeDep), True, False), 0)  '*Eaw 56-01-29   แก้ไขการ Rounding สำหรับ TK Project
                                        cNetSaleRound = SP_FMTcCurPoint(cNetSale + cRoundRcv, True, False)
                                        cTotalDiscAft = SP_FMTcCurPoint(cVB_TotalDiscAft + cVB_AdjustMalCrdOfOntop - cVB_AdjustDiscMalCrdOfCrd, True, False)
                                        cBalance = SP_FMTcCurPoint(cVB_Balance - cVB_AdjustMalCrdOfOntop + cVB_AdjustDiscMalCrdOfCrd - cVB_RoundRcv, True, False)
                                         '*KT 52-02-23 640 or 800
                                        wSub.olaDiscTotal.Caption = SP_FMTcCurPoint(cTotalDiscAft, False, True)
                                        wSub.olaNetSale.Caption = SP_FMTcCurPoint(cNetSale, False, True)
                                        wSub.olaBalance.Caption = SP_FMTcCurPoint(cBalance, False, True)
                                        wSub.Refresh '*KT 53-02-03 แก้ไขตาม POSFront 2.0.33
        '                                DoEvents                '*KT 53-01-14
                                End Select
                            '---------------------------------------------------------------------------------------------------------
                            
                        End If
                    Else
                        otbNum.Text = ""
                        otbNum.Locked = False
                        otbNum.BackColor = nVB_CNClrBUnLock
                        'กรณีที่รูดเครดิตแล้วไม่ได้ Ontop
                        bVB_IsOnTop = False
                        '*KT 51-08-16 ทำการคำนวณค่าใหม่กรณีที่มีการชำระด้วยบัตรเครดิตและมีการเฟกหรือไม่เฟก Mall Card ที่ Credit Card
                        '----------------------------------------------------------------------------------------------------
                        Select Case Val(Format(tVB_TransactionSub))
                            Case tEN_TrnDeposit, tEN_TrnRedeem, tEN_TrnAddDeposit, tEN_TrnDepositCancel, tEN_TrnAddDepositCancel
                                cVB_AdjustMalCrdOfOntop = 0
                                cVB_AdjustDiscMalCrdOfCrd = SP_FMTcCurPoint(oDisc.DIScCalCredit, True, False)
                                
                                cNetSale = SP_FMTcCurPoint(cVB_NetSale - cVB_AdjustMalCrdOfOntop, True, False)
                                cTotalDiscAft = SP_FMTcCurPoint(cVB_TotalDiscAft + cVB_AdjustMalCrdOfOntop - cVB_AdjustDiscMalCrdOfCrd, True, False)
                                cBalance = SP_FMTcCurPoint(cVB_Balance - cVB_RoundRcv - cVB_AdjustMalCrdOfOntop, True, False)
                                '*KT 52-02-23 640 or 800
                                wSub.olaDiscTotal.Caption = SP_FMTcCurPoint(cTotalDiscAft, False, True)
                                wSub.olaNetSale.Caption = SP_FMTcCurPoint(cVB_Total - cTotalDiscAft, False, True)
                                wSub.lblDepositAmt.Caption = SP_FMTcCurPoint(cVB_DepositAmt - cVB_AdjustMalCrdOfOntop, False, True)
                                wSub.olaBalance.Caption = SP_FMTcCurPoint(cBalance, False, True)
                                wSub.Refresh '*KT 53-02-03 แก้ไขตาม POSFront 2.0.33
        '                        DoEvents                '*KT 53-01-14
                            Case Else
                                cVB_AdjustMalCrdOfOntop = 0
                                cVB_AdjustDiscMalCrdOfCrd = SP_FMTcCurPoint(oDisc.DIScCalCredit, True, False)
                                
                                cNetSale = SP_FMTcCurPoint(cVB_NetSale - cVB_AdjustMalCrdOfOntop + cVB_AdjustDiscMalCrdOfCrd, True, False)
'                                cRoundRcv = IIf(bVB_CNAlwRnd, SP_FMTcCurPoint(UT_GetnRoundDiff(cNetSale + cVB_NetDutyFreeDep), True, False), 0)    '*Eaw 56-01-29   แก้ไขการ Rounding สำหรับ TK Project
                                cNetSaleRound = cNetSale + cRoundRcv
                                cTotalDiscAft = SP_FMTcCurPoint(cVB_TotalDiscAft - cVB_AdjustMalCrdOfOntop - cVB_AdjustDiscMalCrdOfCrd, True, False)
                                cBalance = SP_FMTcCurPoint(cVB_Balance - cVB_AdjustMalCrdOfOntop + cVB_AdjustDiscMalCrdOfCrd - cVB_RoundRcv, True, False)
                                 '*KT 52-02-23 640 or 800
                                wSub.olaDiscTotal.Caption = SP_FMTcCurPoint(cTotalDiscAft, False, True)
                                wSub.olaNetSale.Caption = SP_FMTcCurPoint(cNetSale, False, True)
                                wSub.olaBalance.Caption = SP_FMTcCurPoint(cBalance, False, True)
                                wSub.Refresh '*KT 53-02-03 แก้ไขตาม POSFront 2.0.33
        '                        DoEvents            '*KT 53-01-14
                        End Select
                        '---------------------------------------------------------------------------------------------------------
                    End If
                End If
            End With
            Call SP_OBJxSetFocus(otbNum)
            '*Eaw 56-05-13  CommSheet TK-ISS3000-182-ST 00 00 01.xls
'             otbNum.Text = SP_FMTcCurPoint(SP_VALcText2Double(wSub.olaBalance.Caption), False, True)    '*Ao 2013/02/27
            If SP_EDCbCheckEDC Then
'                otbNum.Text = SP_FMTcCurPoint(SP_VALcText2Double(cVB_PayCredit), False, True)
                '*Eaw 56-07-10 CommSheet TK-ISS3000-272-TR.xls
                Select Case tVB_TenderCode
                    Case "T002", "T003"
                        otbNum.Text = SP_FMTcCurPoint(SP_VALcText2Double(cVB_PayCredit), False, True)
                    Case "T017"
                        otbNum.Text = SP_FMTcCurPoint(SP_VALcText2Double(wSub.olaBalance.Caption), False, True)
                End Select
            Else
                otbNum.Text = SP_FMTcCurPoint(SP_VALcText2Double(wSub.olaBalance.Caption), False, True)
            End If
            '------------------------------------------------------
'            If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True          '*Ao 10-02-09
    Else
           'none
    End If
    
 
    Exit Sub
ErrHandle:
    Call SP_MSGxShowError(Err.Number, Err.Description)
    MsgBox "กรุณาติดต่อผู้ดูแลระบบ ก่อนกดปุ่ม OK"
    Call SP_OBJxSetFocus(otbAcc)
'    If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True          '*Ao 10-02-09
End Sub

Private Sub orbMonth_GotFocus()
     orbMonth.SelStart = 0
    orbMonth.SelLength = Len(orbMonth.Text)
End Sub

Private Sub orbMonth_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDelete         'delete
            orbMonth.Text = ""
    End Select
End Sub

Private Sub orbMonth_KeyPress(KeyAscii As Integer)
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

Private Sub orbYear_GotFocus()
    orbYear.SelStart = 0
    orbYear.SelLength = Len(orbYear.Text)
    '
    If otbNum.Locked And otbNum.Text <> "" Then
        Select Case Val(Format(tVB_TransactionSub))
                Case tEN_TrnDeposit, tEN_TrnRedeem, tEN_TrnAddDeposit, tEN_TrnDepositCancel, tEN_TrnAddDepositCancel
                     '*KT 52-02-23 640 or 800
                    wSub.olaDiscTotal.Caption = SP_FMTcCurPoint(cVB_TotalDiscAft, False, True)
                    wSub.olaNetSale.Caption = SP_FMTcCurPoint(cVB_Total - cVB_TotalDiscAft, False, True)
                    wSub.olaBalance.Caption = SP_FMTcCurPoint(cVB_Balance, False, True)
                    wSub.lblDepositAmt.Caption = SP_FMTcCurPoint(cVB_DepositAmt, False, True)
                    wSub.Refresh '*KT 53-02-03 แก้ไขตาม POSFront 2.0.33
'                    DoEvents            '*KT 53-01-14
                Case Else
                     '*KT 52-02-23 640 or 800
                    wSub.olaDiscTotal.Caption = SP_FMTcCurPoint(cVB_TotalDiscAft, False, True)
                    wSub.olaNetSale.Caption = SP_FMTcCurPoint(cVB_NetSale, False, True)
                    wSub.olaBalance.Caption = SP_FMTcCurPoint(cVB_Balance, False, True)
                    wSub.Refresh '*KT 53-02-03 แก้ไขตาม POSFront 2.0.33
'                    DoEvents            '*KT 53-01-14
        End Select
        otbNum.Text = ""
        otbNum.Locked = False
        otbNum.BackColor = nVB_CNClrBUnLock
        otbApvCode.Text = ""
    End If
    '
End Sub

Private Sub orbYear_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDelete         'delete
            orbYear.Text = ""
    End Select
End Sub

Private Sub orbYear_KeyPress(KeyAscii As Integer)
    Call UT_ChkKeyAscii(KeyAscii)
    Select Case KeyAscii
        Case 8, 6           'backspace, delete
            'do nothing
        Case 13
            sendkeys "{Tab}"
'            DoEvents                '*KT 53-01-14
        Case Else
            If Not IsNumeric(Chr(KeyAscii)) Then
                KeyAscii = 0
            End If
    End Select
End Sub

Private Sub orbYear_LostFocus()
'*KT 51-11-11

    Dim tWhe
    Dim cDiscOnTopAmt As Double
    Dim cAdjust As Double
    Dim cNetSale  As Double
    Dim cRoundRcv As Double
    Dim cNetSaleRound As Double
    Dim cTotalDiscAft As Double
    Dim cBalance As Double
    Dim bTime5 As Boolean
    Dim nI As Long
    
    Screen.MousePointer = 11
    If Format(orbMonth.Text, "00") = "00" Then    '*Ao 52-02-05
        Call SP_MSGnShowing("tms_CN005143", nCS_Warning)
        orbMonth.Text = ""
        orbYear.Text = ""
        Call SP_OBJxSetFocus(orbMonth)
        Exit Sub
    End If
    
    '__________________________________________
    'Check Expire  Date Comsheet ML-SF-170
    dVB_CrdDateExp = DateSerial(2000 + Val(Format(orbYear.Text)), Val(Format(orbMonth.Text)) + 1, 1)
    If Format(dVB_CrdDateExp, "YYYY/MM/DD") <= Format(dVB_CurrentDate, "YYYY/MM/DD") Then '*KT 51-12-09 V 1.0.3 Phase 2 ข้อ 26  เปรียบเทียบน้อยกว่าเพราะdVB_CrdDateExp + ขึ้นไป 1 วันจากที่ระบุ Communication Sheet PH2-ML-SF-027 Credit Expire.xls
        Call SP_MSGnShowing("tms_CN005143", nCS_Warning)
        orbMonth.Text = ""
        orbYear.Text = ""
'        otbCardType.Text = ""
'        otbApvCode.Text = ""
        'otbNum.Text = ""
        Call SP_OBJxSetFocus(orbMonth)
        Exit Sub
    End If
    
    With oDisc
         If nVB_OnTopCount >= 5 Then
            Call SP_MSGnShowing("tms_CN005154", nCS_Warning)
            bVB_IsOnTop = False
        Else
            oDisc.DISxClearAll
            oDisc.dTranDate = tUT_SaleDate
            oDisc.tTerminalNo = tVB_CNTerminalNum
            oDisc.tTransNo = tVB_Running
            oDisc.tCardID = tVB_CrdCardNo
            oDisc.tCardType = tVB_CrdType
            oDisc.tCardKey = tVB_CrdKey
            oDisc.tStaChkExpire = "Y"
            oDisc.cDiscValue = SP_VALcText2Double(wSub.olaBalance.Caption)  '*KT 52-02-23 640 or 800
            dVB_CrdDateExp = DateSerial(2000 + Val(Format(orbYear.Text)), Val(Format(orbMonth.Text)), 1)
            oDisc.dCardExpDate = dVB_CrdDateExp
            
            '*KT 54-06-01 V.4.0.0  Req.19 OnTop Specify POSNO
            oDisc.bPOSNoOnTopFlag = bVB_CrdPOSNoOnTopFlag
            oDisc.tPOSNoOnTopAlwType = tVB_CrdPOSNoOnTopAlwType
            oDisc.tFloorAlwOnTop = tVB_CrdFloorAlwOnTop
            oDisc.cSubTotalAll = cUT_SubTotal
            oDisc.cSubTotal = cVB_NetSale
            oDisc.tStaOnTopAlwTierAmtLast = tVB_CrdStaAlwTierAmtLast
            
            'สำหรับให้ทำการ Recal ส่วนลด Mall Card ใหม่ให้เหมือน Special Gift Voucher,Special Coupon
            Select Case SP_VALnText2Double(tVB_TransactionSub)
                Case tEN_TrnRedeem
                    oDisc.cBalance = SP_FMTcCurPoint(cVB_DepossitAmtRedeem, False, True)
                Case tEN_TrnAddDeposit
                    oDisc.cBalance = SP_FMTcCurPoint(cVB_DepossitAmtBAddDpt, False, True)
            End Select
            
'            DoEvents            '*KT 53-01-14
            cVB_CNOntopMaxDisc = SP_FMTcCurPoint(oDisc.DIScCalOntop, True, False) 'Ontop ที่ได้
            cVB_CNOntopMaxCal = SP_FMTcCurPoint(oDisc.cDiscValue, True, False) 'ยอดเงินที่หัก Ontop แล้ว
            cVB_CNOntopMaxAlw = SP_FMTcCurPoint(oDisc.cDiscOntopAlw, True, False) 'ยอดเต็ม '*KT 51-06-20 Comsheet ML-SF-270
            If cVB_CNOntopMaxDisc > 0 Then
                 Screen.MousePointer = 0
                'Communication Sheet ML-SF-463 Must not use same credit card number within 5 ontop.xls
                '--------------------------------------------------------------------------------------------------------------------
                bTime5 = True
                For nI = LBound(aVB_CreditCrd) To UBound(aVB_CreditCrd) - 1
                    If tVB_CrdCardNo = aVB_CreditCrd(nI) Then
                        bTime5 = False
                    End If
                Next nI
                '----------------------------------------------------------------------------------------------------------------------
                 If bTime5 Then
                    bVB_IsOnTop = True
                    wCNPayCreditAmt.Show vbModal
                    If cVB_CrdAmt > 0 And bVB_IsCrdCard Then
                        otbNum.Locked = True
                        otbNum.BackColor = nVB_CNClrBLock
                        otbNum.Text = SP_FMTcCurPoint(cVB_CrdAmt, False, True)
                        Call SP_OBJxSetFocus(otbApvCode)
                        
                        'กรณีที่รูดเครดิตแล้วได้ Ontop
                        '---------------------------------------------------------------------------------------
                        Select Case Val(Format(tVB_TransactionSub))
                           Case tEN_TrnDeposit, tEN_TrnRedeem, tEN_TrnAddDeposit, tEN_TrnDepositCancel, tEN_TrnAddDepositCancel
                                cVB_AdjustMalCrdOfOntop = SP_FMTcCurPoint(cVB_CNOntopMaxDisc, True, False)
                                cVB_AdjustDiscMalCrdOfCrd = SP_FMTcCurPoint(oDisc.cAmtAllMall, True, False)
                                
                                cNetSale = SP_FMTcCurPoint(cVB_NetSale - cVB_AdjustMalCrdOfOntop, True, False)
                                cTotalDiscAft = SP_FMTcCurPoint(cVB_TotalDiscAft + cVB_AdjustMalCrdOfOntop - cVB_AdjustDiscMalCrdOfCrd, True, False)
                                cBalance = SP_FMTcCurPoint(cVB_Balance - cVB_RoundRcv - cVB_AdjustMalCrdOfOntop, True, False)
                                 '*KT 52-02-23 640 or 800
                                wSub.olaDiscTotal.Caption = SP_FMTcCurPoint(cTotalDiscAft, False, True)
                                wSub.olaNetSale.Caption = SP_FMTcCurPoint(cVB_Total - cTotalDiscAft, False, True)
                                wSub.lblDepositAmt.Caption = SP_FMTcCurPoint(cVB_DepositAmt - cVB_AdjustMalCrdOfOntop, False, True)
                                wSub.olaBalance.Caption = SP_FMTcCurPoint(cBalance, False, True)
                                wSub.Refresh '*KT 53-02-03 แก้ไขตาม POSFront 2.0.33
'                                DoEvents            '*KT 53-01-14
                           Case Else
                                cVB_AdjustMalCrdOfOntop = SP_FMTcCurPoint(cVB_CNOntopMaxDisc, True, False)
                                cVB_AdjustDiscMalCrdOfCrd = SP_FMTcCurPoint(oDisc.cAmtAllMall, True, False)
                                
                                cNetSale = SP_FMTcCurPoint(cVB_NetSale - cVB_AdjustMalCrdOfOntop + cVB_AdjustDiscMalCrdOfCrd, True, False)
'                                cRoundRcv = IIf(bVB_CNAlwRnd, SP_FMTcCurPoint(UT_GetnRoundDiff(cNetSale + cVB_NetDutyFreeDep), True, False), 0)   '*Eaw 56-01-29   แก้ไขการ Rounding สำหรับ TK Project
                                cNetSaleRound = SP_FMTcCurPoint(cNetSale + cRoundRcv, True, False)
                                cTotalDiscAft = SP_FMTcCurPoint(cVB_TotalDiscAft + cVB_AdjustMalCrdOfOntop - cVB_AdjustDiscMalCrdOfCrd, True, False)
                                cBalance = SP_FMTcCurPoint(cVB_Balance - cVB_AdjustMalCrdOfOntop + cVB_AdjustDiscMalCrdOfCrd - cVB_RoundRcv, True, False)
                                 '*KT 52-02-23 640 or 800
                                wSub.olaDiscTotal.Caption = SP_FMTcCurPoint(cTotalDiscAft, False, True)
                                wSub.olaNetSale.Caption = SP_FMTcCurPoint(cNetSale, False, True)
                                wSub.olaBalance.Caption = SP_FMTcCurPoint(cBalance, False, True)
                                wSub.Refresh '*KT 53-02-03 แก้ไขตาม POSFront 2.0.33
'                                DoEvents            '*KT 53-01-14
                        End Select
                        '--------------------------------------------------------------------------------------
                    Else
                        otbNum.Text = ""
                        otbNum.Locked = False
                        otbNum.BackColor = nVB_CNClrBUnLock
                        'กรณีที่รูดเครดิตแล้วได้ Ontop แล้วยกเลิก
                        bVB_IsOnTop = False
                        '*KT 51-08-16 ทำการคำนวณค่าใหม่กรณีที่มีการชำระด้วยบัตรเครดิตและมีการเฟกหรือไม่เฟก Mall Card ที่ Credit Card
                        '----------------------------------------------------------------------------------------------------
                        Select Case Val(Format(tVB_TransactionSub))
                           Case tEN_TrnDeposit, tEN_TrnRedeem, tEN_TrnAddDeposit, tEN_TrnDepositCancel, tEN_TrnAddDepositCancel
                                cVB_AdjustMalCrdOfOntop = 0
                                cVB_AdjustDiscMalCrdOfCrd = SP_FMTcCurPoint(oDisc.DIScCalCredit, True, False)
                                
                                cNetSale = SP_FMTcCurPoint(cVB_NetSale - cVB_AdjustMalCrdOfOntop, True, False)
                                cTotalDiscAft = SP_FMTcCurPoint(cVB_TotalDiscAft + cVB_AdjustMalCrdOfOntop - cVB_AdjustDiscMalCrdOfCrd, True, False)
                                cBalance = SP_FMTcCurPoint(cVB_Balance - cVB_RoundRcv - cVB_AdjustMalCrdOfOntop, True, False)
                                 '*KT 52-02-23 640 or 800
                                wSub.olaDiscTotal.Caption = SP_FMTcCurPoint(cTotalDiscAft, False, True)
                                wSub.olaNetSale.Caption = SP_FMTcCurPoint(cVB_Total - cTotalDiscAft, False, True)
                                wSub.lblDepositAmt.Caption = SP_FMTcCurPoint(cVB_DepositAmt - cVB_AdjustMalCrdOfOntop, False, True)
                                wSub.olaBalance.Caption = SP_FMTcCurPoint(cBalance, False, True)
                                wSub.Refresh '*KT 53-02-03 แก้ไขตาม POSFront 2.0.33
'                                DoEvents            '*KT 53-01-14
                                
                            Case Else
                                cVB_AdjustMalCrdOfOntop = 0
                                cVB_AdjustDiscMalCrdOfCrd = SP_FMTcCurPoint(oDisc.DIScCalCredit, True, False)
                                
                                cNetSale = cVB_NetSale - cVB_AdjustMalCrdOfOntop + cVB_AdjustDiscMalCrdOfCrd
                                cTotalDiscAft = SP_FMTcCurPoint(cVB_TotalDiscAft + cVB_AdjustMalCrdOfOntop - cVB_AdjustDiscMalCrdOfCrd, True, False)
                                cBalance = SP_FMTcCurPoint(cVB_Balance - cVB_AdjustMalCrdOfOntop + cVB_AdjustDiscMalCrdOfCrd - cVB_RoundRcv, True, False)
                                 '*KT 52-02-23 640 or 800
                                wSub.olaDiscTotal.Caption = SP_FMTcCurPoint(cTotalDiscAft, False, True)
                                wSub.olaNetSale.Caption = SP_FMTcCurPoint(cNetSale, False, True)
                                wSub.olaBalance.Caption = SP_FMTcCurPoint(cBalance, False, True)
                                wSub.Refresh '*KT 53-02-03 แก้ไขตาม POSFront 2.0.33
'                                DoEvents            '*KT 53-01-14
                        End Select
                        '---------------------------------------------------------------------------------------------------------
                    End If
                Else
   '                otbNum.Text = ""  '*Eaw 56-02-02
'                    otbNum.Text = SP_FMTcCurPoint(SP_VALcText2Double(cVB_PayCredit), False, True)    '*Eaw 56-02-02
                    '*Eaw 56-05-13  CommSheet TK-ISS3000-182-ST 00 00 01.xls
'                    otbNum.Text = SP_FMTcCurPoint(SP_VALcText2Double(wSub.olaBalance.Caption), False, True)    '*Ao 2013/02/27
                    If SP_EDCbCheckEDC Then
                    '*Eaw 56-06-11
'                        otbNum.Text = SP_FMTcCurPoint(SP_VALcText2Double(cVB_PayCredit), False, True)
                        Select Case Val(Format(tVB_TransactionSub))
                            Case tEN_TrnReturnManualTax, tEN_TrnReturnTaxInc, tEN_TrnSaleManualTax, tEN_TrnSaleTaxInc
                                otbNum.Text = SP_FMTcCurPoint(SP_VALcText2Double(wSub.olaBalance.Caption), False, True)
                            Case Else
                                If tVB_CNMnuCurrent = "1831" Then  '*Eaw 56-06-13 กรณีเป็น TC Card ไม่ใช้ EDC  CommSheet TK-ISS3000-229-ST.xls
                                    otbNum.Text = SP_FMTcCurPoint(SP_VALcText2Double(wSub.olaBalance.Caption), False, True)
                                Else
                                    otbNum.Text = SP_FMTcCurPoint(SP_VALcText2Double(cVB_PayCredit), False, True)
                                End If
                        End Select
                    Else
                        otbNum.Text = SP_FMTcCurPoint(SP_VALcText2Double(wSub.olaBalance.Caption), False, True)
                    End If
                    '------------------------------------------------------
                    Call SP_OTBxSelText(otbNum)   '*Eaw 56-02-02
                    otbNum.Locked = False
                    otbNum.BackColor = nVB_CNClrBUnLock
                    'กรณีที่รูดเครดิตแล้วได้ Ontop มาแล้ว 5 ครั้ง
                    bVB_IsOnTop = False
                    '*KT 51-08-16 ทำการคำนวณค่าใหม่กรณีที่มีการชำระด้วยบัตรเครดิตและมีการเฟกหรือไม่เฟก Mall Card ที่ Credit Card
                    '----------------------------------------------------------------------------------------------------
                    Select Case Val(Format(tVB_TransactionSub))
                           Case tEN_TrnDeposit, tEN_TrnRedeem, tEN_TrnAddDeposit, tEN_TrnDepositCancel, tEN_TrnAddDepositCancel
                                cVB_AdjustMalCrdOfOntop = 0
                                cVB_AdjustDiscMalCrdOfCrd = SP_FMTcCurPoint(oDisc.DIScCalCredit, True, False)
                                
                                cNetSale = SP_FMTcCurPoint(cVB_NetSale - cVB_AdjustMalCrdOfOntop, True, False)
                                cTotalDiscAft = SP_FMTcCurPoint(cVB_TotalDiscAft + cVB_AdjustMalCrdOfOntop - cVB_AdjustDiscMalCrdOfCrd, True, False)
                                cBalance = SP_FMTcCurPoint(cVB_Balance - cVB_RoundRcv - cVB_AdjustMalCrdOfOntop, True, False)
                                 '*KT 52-02-23 640 or 800
                                wSub.olaDiscTotal.Caption = SP_FMTcCurPoint(cTotalDiscAft, False, True)
                                wSub.olaNetSale.Caption = SP_FMTcCurPoint(cVB_Total - cTotalDiscAft, False, True)
                                wSub.lblDepositAmt.Caption = SP_FMTcCurPoint(cVB_DepositAmt - cVB_AdjustMalCrdOfOntop, False, True)
                                wSub.olaBalance.Caption = SP_FMTcCurPoint(cBalance, False, True)
                                wSub.Refresh '*KT 53-02-03 แก้ไขตาม POSFront 2.0.33
'                                DoEvents                '*KT 53-01-14
                            Case Else
                                cVB_AdjustMalCrdOfOntop = 0
                                cVB_AdjustDiscMalCrdOfCrd = SP_FMTcCurPoint(oDisc.DIScCalCredit, True, False)
                                
                                cNetSale = SP_FMTcCurPoint(cVB_NetSale - cVB_AdjustMalCrdOfOntop + cVB_AdjustDiscMalCrdOfCrd, True, False)
'                                cRoundRcv = IIf(bVB_CNAlwRnd, SP_FMTcCurPoint(UT_GetnRoundDiff(cNetSale + cVB_NetDutyFreeDep), True, False), 0)  '*Eaw 56-01-29   แก้ไขการ Rounding สำหรับ TK Project
                                cNetSaleRound = SP_FMTcCurPoint(cNetSale + cRoundRcv, True, False)
                                cTotalDiscAft = SP_FMTcCurPoint(cVB_TotalDiscAft + cVB_AdjustMalCrdOfOntop - cVB_AdjustDiscMalCrdOfCrd, True, False)
                                cBalance = SP_FMTcCurPoint(cVB_Balance - cVB_AdjustMalCrdOfOntop + cVB_AdjustDiscMalCrdOfCrd - cVB_RoundRcv, True, False)
                                 '*KT 52-02-23 640 or 800
                                wSub.olaDiscTotal.Caption = SP_FMTcCurPoint(cTotalDiscAft, False, True)
                                wSub.olaNetSale.Caption = SP_FMTcCurPoint(cNetSale, False, True)
                                wSub.olaBalance.Caption = SP_FMTcCurPoint(cBalance, False, True)
                                wSub.Refresh '*KT 53-02-03 แก้ไขตาม POSFront 2.0.33
'                                DoEvents                '*KT 53-01-14
                        End Select
                    '---------------------------------------------------------------------------------------------------------
                    
                End If
            Else
'                otbNum.Text = ""  '*Eaw 56-02-02
'                otbNum.Text = SP_FMTcCurPoint(SP_VALcText2Double(cVB_PayCredit), False, True)    '*Eaw 56-02-02
                '*Eaw 56-05-13  CommSheet TK-ISS3000-182-ST 00 00 01.xls
'                otbNum.Text = SP_FMTcCurPoint(SP_VALcText2Double(wSub.olaBalance.Caption), False, True)    '*Ao 2013/02/27
                If SP_EDCbCheckEDC Then
                '*Eaw 56-06-11
'                    otbNum.Text = SP_FMTcCurPoint(SP_VALcText2Double(cVB_PayCredit), False, True)
                    Select Case Val(Format(tVB_TransactionSub))
                        Case tEN_TrnReturnManualTax, tEN_TrnReturnTaxInc, tEN_TrnSaleManualTax, tEN_TrnSaleTaxInc
                            otbNum.Text = SP_FMTcCurPoint(SP_VALcText2Double(wSub.olaBalance.Caption), False, True)
                        Case Else
                            If tVB_CNMnuCurrent = "1831" Then  '*Eaw 56-06-13 กรณีเป็น TC Card ไม่ใช้ EDC  CommSheet TK-ISS3000-229-ST.xls
                                otbNum.Text = SP_FMTcCurPoint(SP_VALcText2Double(wSub.olaBalance.Caption), False, True)
                            Else
                                If otbNum.Text = "" Then otbNum.Text = SP_FMTcCurPoint(SP_VALcText2Double(cVB_PayCredit), False, True)
                                                                
                            End If
                    End Select
                Else
                    otbNum.Text = SP_FMTcCurPoint(SP_VALcText2Double(wSub.olaBalance.Caption), False, True)
                End If
                '------------------------------------------------------
                Call SP_OTBxSelText(otbNum)   '*Eaw 56-02-02
                otbNum.Locked = False
                otbNum.BackColor = nVB_CNClrBUnLock
                'กรณีที่รูดเครดิตแล้วไม่ได้ Ontop
                bVB_IsOnTop = False
                '*KT 51-08-16 ทำการคำนวณค่าใหม่กรณีที่มีการชำระด้วยบัตรเครดิตและมีการเฟกหรือไม่เฟก Mall Card ที่ Credit Card
                '----------------------------------------------------------------------------------------------------
                Select Case Val(Format(tVB_TransactionSub))
                    Case tEN_TrnDeposit, tEN_TrnRedeem, tEN_TrnAddDeposit, tEN_TrnDepositCancel, tEN_TrnAddDepositCancel
                        cVB_AdjustMalCrdOfOntop = 0
                        cVB_AdjustDiscMalCrdOfCrd = SP_FMTcCurPoint(oDisc.DIScCalCredit, True, False)
                        
                        cNetSale = SP_FMTcCurPoint(cVB_NetSale - cVB_AdjustMalCrdOfOntop, True, False)
                        cTotalDiscAft = SP_FMTcCurPoint(cVB_TotalDiscAft + cVB_AdjustMalCrdOfOntop - cVB_AdjustDiscMalCrdOfCrd, True, False)
                        cBalance = SP_FMTcCurPoint(cVB_Balance - cVB_RoundRcv - cVB_AdjustMalCrdOfOntop, True, False)
                       '*KT 52-02-23 640 or 800
                        wSub.olaDiscTotal.Caption = SP_FMTcCurPoint(cTotalDiscAft, False, True)
                        wSub.olaNetSale.Caption = SP_FMTcCurPoint(cVB_Total - cTotalDiscAft, False, True)
                        wSub.lblDepositAmt.Caption = SP_FMTcCurPoint(cVB_DepositAmt - cVB_AdjustMalCrdOfOntop, False, True)
                        wSub.olaBalance.Caption = SP_FMTcCurPoint(cBalance, False, True)
                        wSub.Refresh '*KT 53-02-03 แก้ไขตาม POSFront 2.0.33
'                        DoEvents                '*KT 53-01-14
                    Case Else
                        cVB_AdjustMalCrdOfOntop = 0
                        cVB_AdjustDiscMalCrdOfCrd = SP_FMTcCurPoint(oDisc.DIScCalCredit, True, False)
                        
                        cNetSale = SP_FMTcCurPoint(cVB_NetSale - cVB_AdjustMalCrdOfOntop + cVB_AdjustDiscMalCrdOfCrd, True, False)
'                        cRoundRcv = IIf(bVB_CNAlwRnd, SP_FMTcCurPoint(UT_GetnRoundDiff(cNetSale + cVB_NetDutyFreeDep), True, False), 0)  '*Eaw 56-01-29   แก้ไขการ Rounding สำหรับ TK Project
                        cNetSaleRound = cNetSale + cRoundRcv
                        cTotalDiscAft = SP_FMTcCurPoint(cVB_TotalDiscAft - cVB_AdjustMalCrdOfOntop - cVB_AdjustDiscMalCrdOfCrd, True, False)
                        cBalance = SP_FMTcCurPoint(cVB_Balance - cVB_AdjustMalCrdOfOntop + cVB_AdjustDiscMalCrdOfCrd - cVB_RoundRcv, True, False)
                         '*KT 52-02-23 640 or 800
                        wSub.olaDiscTotal.Caption = SP_FMTcCurPoint(cTotalDiscAft, False, True)
                        wSub.olaNetSale.Caption = SP_FMTcCurPoint(cNetSale, False, True)
                        wSub.olaBalance.Caption = SP_FMTcCurPoint(cBalance, False, True)
                        wSub.Refresh '*KT 53-02-03 แก้ไขตาม POSFront 2.0.33
'                        DoEvents                '*KT 53-01-14
                End Select
                '---------------------------------------------------------------------------------------------------------
            End If
        End If
    End With
    Screen.MousePointer = 0
End Sub

'Private Sub oScanner_DataEvent(ByVal Status As Long)
'On Error GoTo ErrHandle
'
'    With OPOSMSR1                                   '*Ao 10-02-09
'        If bVB_OPOSMSREnable Then
'            'Enable the event.
'            .DataEventEnabled = False
'        End If
'    End With
'
'    With oScanner
'        otbAcc.Text = .ScanData
''        DoEvents                '*KT 53-01-14
'        If otbAcc.Text <> "" Then
'            '*KT 52-05-06 PH 1.8 ML-SF-056 TP 2000 ไม่สามารถใช้ Scanner  Scan สินค้าได้ โดยเมื่อ Scan สินค้าแล้วมีรหัสสินค้าปรากฏขึ้นมาถูกต้องแต่ จะขึ้นข้อความไม่พบสินค้านั้น กรณีที่ OPOS บางตัวมีการอ่านค่า Enter ต่อท้าย ต้องตัดออก
'            '*KT 52-05-15 PH 1.8 ML-SF-061 Price Check ไม่สามารถใช้ Scanner ได้ จะขึ้นข้อความไม่พบรหัสสินค้า
'            If Asc(Right(otbAcc.Text, 1)) = 13 Then
'                otbAcc.Text = Mid(otbAcc.Text, 1, Len(otbAcc.Text) - 1)
'            End If
'            Call otbAcc_KeyDown(13, 0)
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
'
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
'    If bVB_OPOSMSREnable Then OPOSMSR1.DataEventEnabled = True      '*Ao 10-02-09
'    Exit Sub
'ErrHandle:
'    MsgBox "Scaner error" & vbCrLf & "Runtime : " & Err.Number & vbCrLf & Err.Description
'    If bVB_OPOSMSREnable Then OPOSMSR1.DataEventEnabled = True      '*Ao 10-02-09
'End Sub

Private Sub otbAcc_Change()
    bW_ChkCrd = False
    '***KT 51-11-20 ป้องกันการใช้เมาส์คลิก
    orbMonth.Text = ""
    orbYear.Text = ""
    otbCardType.Text = ""
    otbApvCode.Text = ""
'    otbNum.Text = ""
End Sub

Private Sub otbAcc_GotFocus()
    otbAcc.SelStart = 0
    otbAcc.SelLength = Len(otbAcc.Text)
    
   If otbNum.Locked And otbNum.Text <> "" Then
        Select Case Val(Format(tVB_TransactionSub))
            Case tEN_TrnDeposit, tEN_TrnRedeem, tEN_TrnAddDeposit, tEN_TrnDepositCancel, tEN_TrnAddDepositCancel
                 '*KT 52-02-23 640 or 800
                wSub.olaDiscTotal.Caption = SP_FMTcCurPoint(cVB_TotalDiscAft, False, True)
                wSub.olaNetSale.Caption = SP_FMTcCurPoint(cVB_Total - cVB_TotalDiscAft, False, True)
                wSub.olaBalance.Caption = SP_FMTcCurPoint(cVB_Balance, False, True)
                wSub.lblDepositAmt.Caption = SP_FMTcCurPoint(cVB_DepositAmt, False, True)
                wSub.Refresh '*KT 53-02-03 แก้ไขตาม POSFront 2.0.33
'                    DoEvents            '*KT 53-01-14
            Case Else
                 '*KT 52-02-23 640 or 800
                wSub.olaDiscTotal.Caption = SP_FMTcCurPoint(cVB_TotalDiscAft, False, True)
                wSub.olaNetSale.Caption = SP_FMTcCurPoint(cVB_NetSale, False, True)
                wSub.olaBalance.Caption = SP_FMTcCurPoint(cVB_Balance, False, True)
                wSub.Refresh '*KT 53-02-03 แก้ไขตาม POSFront 2.0.33
'                    DoEvents            '*KT 53-01-14
        End Select
        otbNum.Text = ""
        otbNum.Locked = False
        otbNum.BackColor = nVB_CNClrBUnLock
        otbApvCode.Text = ""
    End If
End Sub

Private Sub otbAcc_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim nI As Long
    
On Error GoTo ErrHandle
    Select Case KeyCode
'        Case vbKeyEscape                           '*Ao 2013/01/09
'            If tVB_CNMnuCurrent = "1831" Then     '*Eaw 56-01-03               '*Ao 2013/01/09
'                Call ocmCancel_Click
'            End If
'            If Not bW_HandleESC Then    '*Eaw 55-11-29  TEST   '*Ao 56-01-03
'                    Set oEdc = Nothing
'                    bW_ManaulMode = True
'                    ocmCancel.Cancel = True
'                    otbAcc.Locked = False
'                    otbAcc.BackColor = nVB_CNClrBUnLock
'                    Call SP_OBJxSetFocus(otbAcc)
'            End If
        Case vbKeyDelete         'delete
            otbAcc.Text = ""
        Case 13
            
            bW_ChkCrd = True
            
            tVB_CrdCardNo = Trim(otbAcc.Text)
'            If Trim(ocbType.Text) <> "-" Then
'                If W_CHKbCupsCard(Trim(ocbType.Text)) Then
''                    ocbType.Text = tVB_CrdCardType
'                    otbNum.MaxLength = 13 '
'                    If SP_EDCbCheckEDC Then
'                        otbNum.Text = SP_FMTcCurPoint(SP_VALcText2Double(cVB_PayCredit), False, True)
'                    Else
'                        otbNum.Text = SP_FMTcCurPoint(SP_VALcText2Double(wSub.olaBalance.Caption), False, True)
'                    End If
'
'                    bW_FoundAcc = True
'                    SendKeys "{Tab}"
'                Else
'                    Call W_CHKxCaseERROR(nW_CaseERROR)                      '*Ao 2013/02/22   CommSheet TK-ISS3000-011
'                    bW_FoundAcc = False
'                    otbAcc.Text = ""
'                    orbMonth.Text = ""
'                    orbYear.Text = ""
'                    otbCardType.Text = ""
'                    otbApvCode.Text = ""
'                    otbNum.Text = ""
'                    ocbType.ListIndex = 0  '*Eaw 56-05-07 CommSheet TK-ISS3000-159-IT2 00.00.02.xls
'                    Call SP_OTBxSelText(otbAcc)
'                End If
'                Exit Sub
'            End If
            '*Eaw 56-01-29 ตรวจสอบข้อมูลจาก Credit card no. เท่านั้น
            '---------------------------------------------------------------------------------------------------
            If Not W_PRCxFindCrdByAccNum(tVB_CrdCardNo) Then
'                    Call SP_MSGnShowing("tms_CN005065", nCS_Error)
                    Call W_CHKxCaseERROR(nW_CaseERROR)                      '*Ao 2013/02/22   CommSheet TK-ISS3000-011
                    bW_FoundAcc = False
                    otbAcc.Text = ""
                    orbMonth.Text = ""
                    orbYear.Text = ""
                    otbCardType.Text = ""
                    otbApvCode.Text = ""
                    otbNum.Text = SP_FMTcCurPoint(SP_VALcText2Double(cVB_PayCredit), False, True)
                    If ocbType.Enabled Then
                        ocbType.ListIndex = 0  '*Eaw 56-05-07 CommSheet TK-ISS3000-159-IT2 00.00.02.xls
                    End If
'                    ocbType.ListIndex = 0  '*Eaw 56-05-07 CommSheet TK-ISS3000-159-IT2 00.00.02.xls
                    Call SP_OTBxSelText(otbAcc)
            Else
                '*Eaw 56-05-07 CommSheet TK-ISS3000-159-IT2 00.00.02.xls
'                    If Not W_CHKbCupsCard(Trim(ocbType.Text)) Then
'                        ocbType.Text = tVB_CrdCardType
'                    End If
'                    ocbType.Text = tVB_CrdCardType
                    '---------------------------------------------------------------------------------------------
'                    otbNum.Text = SP_FMTcCurPoint(SP_VALcText2Double(cVB_PayCredit), False, True)    '*Eaw 56-02-02
                    otbNum.MaxLength = 13 '*Eaw 56-03-02  TK-ISS3000-043 กรณี Auto fill balance
                    '*Eaw 56-05-13  CommSheet TK-ISS3000-182-ST 00 00 01.xls
'                    otbNum.Text = SP_FMTcCurPoint(SP_VALcText2Double(wSub.olaBalance.Caption), False, True)    '*Ao 2013/02/27
                    If tVB_CNMnuCurrent = "1831" Then  '*Eaw 56-06-13 กรณีเป็น TC Card ไม่ใช้ EDC  CommSheet TK-ISS3000-229-ST.xls
                            otbNum.Text = SP_FMTcCurPoint(SP_VALcText2Double(wSub.olaBalance.Caption), False, True)
                            bW_FoundAcc = True
                            sendkeys "{Tab}"
                            Exit Sub
                    End If
                    If SP_EDCbCheckEDC Then
                        '*Eaw 56-06-11
'                        otbNum.Text = SP_FMTcCurPoint(SP_VALcText2Double(cVB_PayCredit), False, True)
                        Select Case Val(Format(tVB_TransactionSub))
                            Case tEN_TrnReturnManualTax, tEN_TrnReturnTaxInc, tEN_TrnSaleManualTax, tEN_TrnSaleTaxInc
                                otbNum.Text = SP_FMTcCurPoint(SP_VALcText2Double(wSub.olaBalance.Caption), False, True)
                            Case Else
                                otbNum.Text = SP_FMTcCurPoint(SP_VALcText2Double(cVB_PayCredit), False, True)
                                'otbNum.Text = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance), False, True)
                        End Select
                    Else
                        'otbNum.Text = SP_FMTcCurPoint(SP_VALcText2Double(wSub.olaBalance.Caption), False, True)
                        otbNum.Text = SP_FMTcCurPoint(cVB_PayCredit, True, True) '*Bump 58-05-18
                    End If
                    '------------------------------------------------------
                    bW_FoundAcc = True
                    sendkeys "{Tab}"
            End If
'            If tVB_CNMnuCurrent <> "1831" Then     '*Eaw 56-01-03
'                SendKeys "{Tab}"
'            Else
'                If Not W_PRCxFindCrdByAccNum(tVB_CrdCardNo) Then     '*Eaw 56-01-03
'                    Call SP_MSGnShowing("tms_CN005065", nCS_Error)
'                    bW_FoundAcc = False
'                    otbAcc.Text = ""
'                    orbMonth.Text = ""
'                    orbYear.Text = ""
'                    otbCardType.Text = ""
'                    otbApvCode.Text = ""
'                    otbNum.Text = ""
'                    Call SP_OTBxSelText(otbAcc)
'                Else
'                    bW_FoundAcc = True
'                    SendKeys "{Tab}"
'                End If
'            End If
            
            '-------------------------------------------------------------------------------------------------------------
'            'Communication Sheet ML-SF-463 Must not use same credit card number within 5 ontop.xls
'            '--------------------------------------------------------------------------------------------------------------------
'            For nI = LBound(aVB_CreditCrd) To UBound(aVB_CreditCrd) - 1
'                If tVB_CrdCardNo = aVB_CreditCrd(nI) Then
'                    Call SP_MSGnShowing("tms_CN005188", nCS_Error)
'                    Exit Sub
'                End If
'            Next nI
'            '----------------------------------------------------------------------------------------------------------------------
'            If bW_ManaulMode And tVB_CNMnuCurrent <> "1831" Then   '*Eaw 56-01-03
'                 Call SP_OBJxSetFocus(ocbType)
'                Exit Sub
'            End If
          
'            If Not W_PRCxFindCrdByAccNum(tVB_CrdCardNo) Then     '*Eaw 56-01-03
'                Call SP_MSGnShowing("tms_CN005065", nCS_Error)
'                bW_FoundAcc = False
'                otbAcc.Text = ""
'                orbMonth.Text = ""
'                orbYear.Text = ""
'                otbCardType.Text = ""
'                otbApvCode.Text = ""
'                otbNum.Text = ""
'                Call SP_OTBxSelText(otbAcc)
'            Else
'                bW_FoundAcc = True
'                SendKeys "{Tab}"
'            End If
            
    End Select
    Exit Sub
ErrHandle:
    
End Sub

Private Sub otbAcc_KeyPress(KeyAscii As Integer)
    Call UT_ChkKeyAscii(KeyAscii)
    Call SP_ChkKeyAsciiNumber(KeyAscii)
End Sub

Private Sub otbAcc_LostFocus()
    otbAcc.Text = SP_TXTtTrimRightLeft(otbAcc.Text)
'    If bW_ManaulMode And tVB_CNMnuCurrent <> "1831" Then    '*Eaw 56-01-03
'        Call SP_OBJxSetFocus(ocbType)
'        Exit Sub
'    End If
'    If bW_ChkCrd = False Then
'            tVB_CrdCardNo = Trim(otbAcc.Text)
'            If Not W_PRCxFindCrdByAccNum(tVB_CrdCardNo) Then   '*Eaw 56-01-03
'                Call SP_MSGnShowing("tms_CN005065", nCS_Error)
'                bW_FoundAcc = False
'                otbAcc.Text = ""
'                orbMonth.Text = ""
'                orbYear.Text = ""
'                otbCardType.Text = ""
'                otbApvCode.Text = ""
'                otbNum.Text = ""
'                Call SP_OTBxSelText(otbAcc)
'            Else
'                bW_FoundAcc = True
'            End If
'    End If
    bW_ChkCrd = True
End Sub

Private Sub otbApvCode_GotFocus()
    otbApvCode.SelStart = 0
    otbApvCode.SelLength = Len(otbApvCode.Text)
End Sub

Private Sub otbApvCode_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDelete         'delete
            otbApvCode.Text = ""
    End Select
End Sub

Private Sub otbApvCode_KeyPress(KeyAscii As Integer)
    Call UT_ChkKeyAscii(KeyAscii)
    If KeyAscii = 13 Then
        If bW_KeyEnter Then
            KeyAscii = 0
        Else
            bW_KeyEnter = True
            Call ocmOK_Click
        End If
    End If
End Sub

Private Sub otbNum_Change()     '*AO51-03-12    Comm.Sheet SF- 070
    If otbNum.Locked = False Then
        otbNum.Text = SP_DEFtChangeNum(Trim(otbNum.Text))
        otbNum.SelStart = Len(otbNum.Text)
    End If
End Sub

Private Sub otbNum_GotFocus()
    otbNum.SelStart = 0   '*Eaw 56-02-02
'    otbNum.SelStart = cVB_PayCredit
    otbNum.SelLength = Len(otbNum.Text)
End Sub

Private Sub otbNum_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDelete         'delete
            If Not otbNum.Locked Then otbNum.Text = ""
    End Select
End Sub

Private Sub otbNum_KeyPress(KeyAscii As Integer)
    Call UT_ChkKeyAscii(KeyAscii)
    Select Case KeyAscii
        Case 8, 6           'backspace, delete
            'do nothing
        Case 13
            sendkeys "{Tab}"
        Case 46             'ทศนิยม
            KeyAscii = 0
        Case Else
            If Not IsNumeric(Chr(KeyAscii)) Then
                KeyAscii = 0
            Else
                otbNum.MaxLength = 11  '*Eaw 56-03-02  TK-ISS3000-043 กรณีให้คีย์มูลค่าเองมีขนาด 8+2
            End If
    End Select
End Sub

Private Sub W_FRMxSetInitial()
'------------------------------------
'Call:
'Cmt: Set data start initial
'--------------------------------------
    'set value
    cVB_CNOntopMaxCal = 0
    cVB_CNOntopMaxDisc = 0
    otbNum.Text = 0
    'lock control
    otbCardType.Locked = True
'    ocbType.Enabled = False   '*Eaw 55-11-30
    otbCardType.BackColor = nVB_CNClrBLock
End Sub

'อย่าลบ[EDC]
'Private Sub W_EDCxCreditCard()
'    With oHyperCom
'        .SetNetAmout = otbNum.Text
'        If .EDC1CreateMessage Then
'            If W_EDC2HyperComSendMsg(tVB_CNEdcPos) Then                             'Create New Function in Module SP     ส่งพอร์ตที่ต้องการติดต่อไป
'                If .EDC3VerifyCreditCard Then    'IN DLL
'                    otbAcc.Text = .GetCardNumber
'                    otbCardType.Text = .GetCardBankCode
'                    'otbTrace.Text = .GetTraceCode
'                    otbApvCode.Text = .GetApproveCode
'                    ocmCancel.Enabled = False
'                Else
'                    MsgBox .GetMessage, vbInformation
'                End If
'           Else
'                MsgBox .GetMessage, vbInformation
'           End If
'        Else
'           MsgBox .GetMessage, vbInformation
'        End If
'    End With
'End Sub

'อย่าลบ[EDC]
'Private Sub W_EDCxPosNetCreditCard()    '*Gon 50-03-21
''-----------------------------------------------------------
''   Cmt : PosNet
''   Create by :
''-----------------------------------------------------------
'    With oPosNet
'        .SetNetAmout = otbNum.Text
'        If .EDC1CreateMessage Then
'           If W_EDC2PosNetSendMsg(tVB_CNEdcPos) Then
'                If .EDC3VerifyCreditCard Then
'                    otbAcc.Text = .GetCardNumber
'                    otbCardType.Text = .GetCardBankCode
'                    'otbTrace.Text = .GetTraceCode
'                    otbApvCode.Text = .GetApproveCode
'                    ocmCancel.Enabled = False
'                Else
'                    MsgBox .GetMessage, vbInformation
'                End If
'           Else
'                MsgBox .GetMessage, vbInformation
'           End If
'        Else
'           MsgBox .GetMessage, vbInformation
'        End If
'    End With
'End Sub

'อย่าลบ[EDC]
'Public Function W_EDC2HyperComSendMsg(ByVal pnComport As Integer, Optional ByVal pnTimeSec% = 10, Optional ByVal ptSetting = "") As Boolean
''-----------------------------------------------------------------------------
''   Call :
''   Cmt :
''   Developer : SOMJAI
''   Date Create : 10-04-2005 :  14:58:11
''   Date Modified : 10-05-2007 : 11:45 BY JIRAWAT
''   '*Gon 50-09-06
''----------------------------------------------------------------------------
'    On Error GoTo ErrHandle
'    Dim occPort As MSComm ' MSCommLib.MSComm
'    Dim nWait As Long
'    Dim tRet$
'    Dim bAck As Boolean
'    Dim bApv As Boolean
'
'    W_EDC2HyperComSendMsg = False
'    oHyperCom.W_SETxMsgInput = ""                                     '---------------------Call DLL 1
'
'    With wMain.ocoForm
'        If .PortOpen Then .PortOpen = False         'if open, close
'        .CommPort = pnComport
'
'        .Settings = IIf(ptSetting = "", "9600,N,8,1", ptSetting)
'        .PortOpen = True
'        .Output = oHyperCom.W_GETtMsgOutput                   '---------------------Call DLL 2
'
'        nWait = 0       'reset
'        Do
'            tRet = .Input
'            If tRet = Chr(6) Then bAck = True          'ถ้าเครื่องตอบรับจะส่งค่า 06H หรือ Chr(6) กลับมา จะตอบกลับทันทีที่เครื่อง available
'            nWait = nWait + 1
'        Loop Until tRet = Chr(6) Or nWait = pnTimeSec * 1000     '10000(ตั้งได้ เป็น timer ก็ได้ ไม่ต้องตั้งนาน) waiting ถ้าไม่ Ack จริงๆ ก็จะได้ออกจาก loop เหมือนกัน(ป้องกัน Infinity loop)
'
'        'อาจจะเกิด error อื่นๆ ลองทดสอบดูอีกที
'        'ถ้าไม่ตอบกลับ
'        If Not bAck Then
'            oHyperCom.W_SETxMsgError = "Please check terminal"            '---------------------Call DLL 3
'            Exit Function
'        End If
'
'        'ตอบกลับ, ช่วงนี้เครื่องกำลังจะรอรูดบัตร เมื่อรูดเสร็จเครื่องจะส่งค่ากลับอีกครั้ง จึงต้อง loop รออีกรอบ
'        nWait = 0       'reset ค่าใหม่
'        Do
'            DoEvents
'            tRet = tRet & .Input
'
'            '*Gon 50-11-07
'            If bIsUnload Then
'                Exit Do
'            End If
'            '****
'
'            If InStr(tRet, "NA") Then       'ถ้าเครื่องรอนานแล้วไม่ยอดรูด จะส่งค่ากลับมาเป็น ข้อความยาว โดยจะมีค่า "NA" อยู่ด้วยแต่ไม่มี "D4" ประมาณไมเกิน 1 นาที (จริงๆ ตำแหน่งนี้จะมีในตาราง Code อยู่)
'                bApv = False
'                Exit Do
'            End If
'            If InStr(tRet, "D4") Then bApv = True      'ถ้าเครื่องถูกรูดแล้ว ตอบรับจะส่งค่า เป็นข้อความยาวๆ กลับมา จากการสังเกตแล้ว ตอนท้ายของข้อความจะมีค่า "D4" อยู่เพียงที่เดียว ก็เลยเช็คด้วยค่านี้(อาจมีวิธีอื่น)
'            nWait = nWait + 1
'        Loop Until InStr(tRet, "D4") Or nWait = pnTimeSec * 1000000      '(ตั้งได้ เป็น timer ก็ได้ แต่ควรจะมากกว่า ret ค่า "NA") waiting ถ้าไม่ Apv จริงๆ(เช่นไม่ยอมรูดบัตร) ก็จะได้ออกจาก loop เหมือนกัน(ป้องกัน Infinity loop)
'
'        oHyperCom.W_SETxMsgInput = tRet                                                           '---------------------Call DLL 4
'
'        W_EDC2HyperComSendMsg = True
'    End With
'    Exit Function
'ErrHandle:
'    oHyperCom.W_SETxMsgError = Err.Description                                        '---------------------Call DLL 5
'End Function

'อย่าลบ[EDC]
'Public Function W_EDC2PosNetSendMsg(ByVal pnComport As Integer, Optional ByVal pnTimeSec% = 10, Optional ByVal ptSetting = "") As Boolean
''-----------------------------------------------------------
''   Call :
''   Cmt :
''   Developer : SOMJAI
''   Date Create : 10-04-2005 :  14:58:11
''   Date Modified : 06-09-2007 : 17:30 BY JIRAWAT
''   '*Gon 50-09-06
''-----------------------------------------------------------
'    On Error GoTo ErrHandle
'    Dim occPort As MSComm ' MSCommLib.MSComm
'    Dim nWait As Long
'    Dim tRet$
'    Dim tBegin As String
'    Dim bAck As Boolean
'    Dim bApv As Boolean
'
'    W_EDC2PosNetSendMsg = False
'    oPosNet.W_SETxMsgInput = ""                                     '---------------------Call DLL 1
'    With wMain.ocoForm
'        If .PortOpen Then .PortOpen = False         'if open, close
'        .CommPort = pnComport
'
'        .Settings = IIf(ptSetting = "", "9600,N,8,1", ptSetting)
'        .PortOpen = True
'        .Output = oPosNet.W_GETtMsgOutput                   '---------------------Call DLL 2
'
'        nWait = 0       'reset
'        Do
'            tRet = .Input
'            If tRet = Chr(6) Then bAck = True          'ถ้าเครื่องตอบรับจะส่งค่า 06H หรือ Chr(6) กลับมา จะตอบกลับทันทีที่เครื่อง available
'            nWait = nWait + 1
'        Loop Until tRet = Chr(6) Or nWait = 100000     '10000(ตั้งได้ เป็น timer ก็ได้ ไม่ต้องตั้งนาน) waiting ถ้าไม่ Ack จริงๆ ก็จะได้ออกจาก loop เหมือนกัน(ป้องกัน Infinity loop)
'
'        'อาจจะเกิด error อื่นๆ ลองทดสอบดูอีกที
'        'ถ้าไม่ตอบกลับ
'        If Not bAck Then
'            oPosNet.W_SETxMsgError = "Please check terminal"            '---------------------Call DLL 3
'            Exit Function
'        End If
'
'        '*Gon 50-03-21*************************************************
'        'ตอบกลับ, ช่วงนี้เครื่องกำลังจะรอรูดบัตร เมื่อรูดเสร็จเครื่องจะส่งค่ากลับอีกครั้ง จึงต้อง loop รออีกรอบ
''        nWait = 0       'reset ค่าใหม่
''        Do
''            DoEvents
''            tRet = tRet & .Input
''            If InStr(tRet, "NA") Then       'ถ้าเครื่องรอนานแล้วไม่ยอดรูด จะส่งค่ากลับมาเป็น ข้อความยาว โดยจะมีค่า "NA" อยู่ด้วยแต่ไม่มี "D4" ประมาณไมเกิน 1 นาที (จริงๆ ตำแหน่งนี้จะมีในตาราง Code อยู่)
''                bApv = False
''                Exit Do
''            End If
''            If InStr(tRet, "D4") Then bApv = True      'ถ้าเครื่องถูกรูดแล้ว ตอบรับจะส่งค่า เป็นข้อความยาวๆ กลับมา จากการสังเกตแล้ว ตอนท้ายของข้อความจะมีค่า "D4" อยู่เพียงที่เดียว ก็เลยเช็คด้วยค่านี้(อาจมีวิธีอื่น)
''            nWait = nWait + 1
''        Loop Until InStr(tRet, "D4") Or nWait = pnTimeSec * 1000000      '(ตั้งได้ เป็น timer ก็ได้ แต่ควรจะมากกว่า ret ค่า "NA") waiting ถ้าไม่ Apv จริงๆ(เช่นไม่ยอมรูดบัตร) ก็จะได้ออกจาก loop เหมือนกัน(ป้องกัน Infinity loop)
'        nWait = 0        'reset ค่าใหม่
'        Do
'            tRet = .Input
'            tBegin = tBegin & tRet
'            If InStr(1, tBegin, Chr(2) & Chr(83) & Chr(65) & Chr(76) & Chr(69), vbTextCompare) Then bApv = True 'ถ้าเจอข้อมูลเริ่มต้น
'            nWait = nWait + 1
'            DoEvents
'        Loop Until InStr(1, tBegin, Chr(2) & Chr(83) & Chr(65) & Chr(76) & Chr(69), vbTextCompare) Or nWait >= 1600000  'ตั้งค่า ลูบไว้ ป้องกัน Infinity loop ในกรณีเครื่องตอบรับแล้วไม่ได้รูดบัตร
'
'        'ถ้าหาตำแหน่งเริ่มต้นไม่พบ
'        If Not bApv Then
'            W_EDC2PosNetSendMsg = False
'            oPosNet.W_SETxMsgError = "Please Pay agian" 'tW_MsgError    '---------------------Call DLL 4
'            Exit Function
'        End If
'
'        nWait = 0       'reset ค่าใหม่
'         'เก็บข้อมูล ถึง ตำแน่งสุดท้ายข้อมูล (chr(3)) ข้อมูลที่ได้มีความยาว 228
'        Do
'            DoEvents
'            tRet = tRet & .Input
'            If InStr(1, tRet, Chr(3), vbTextCompare) Then bApv = True
'            nWait = nWait + 1
'        Loop Until InStr(1, tRet, Chr(3), vbTextCompare)
'        '-*********************************************************************
'
'        'ในกรณีที่ User กดปุ่ม Cancel ที่ตัว PosNet
'        If InStr(1, tRet, Chr(2) & Chr(83) & Chr(65) & Chr(76) & Chr(69) & Chr(48) & Chr(48) & Chr(50) & Chr(67) & Chr(65), vbTextCompare) Then
'            W_EDC2PosNetSendMsg = False
'            oPosNet.W_SETxMsgError = "Cancel"                                             '---------------------Call DLL 5
'            Exit Function
'        End If
'
'        oPosNet.W_SETxMsgInput = tRet                                                           '---------------------Call DLL 6
'
'        W_EDC2PosNetSendMsg = True
'    End With
'    Exit Function
'ErrHandle:
'    oPosNet.W_SETxMsgError = Err.Description                                        '---------------------Call DLL  7
'End Function

''อย่าลบ[EDC]
'Private Sub otmForm_Timer()
'    If otbNum.Text = "" Then
'        Unload Me
'        Exit Sub
'    End If
'    If otbAcc.Text = "" Then
'        If Me.ForeColor = vbRed Then
'            Me.ForeColor = vbBlue
'            Me.Caption = tVB_TenderType & "  " & "Please Swipe a Card"
'        Else
'            Me.ForeColor = vbRed
'            Me.Caption = tVB_TenderType & "  " & "Please Swipe a Card"
'        End If
'        If bVB_CNEdcPos Then
'            If tVB_TypeOfEDC = "1" Then
'                Set oHyperCom = CreateObject("AdaEDC.cHyperCom")
'                Call W_EDCxCreditCard
'            Else
'                Set oPosNet = CreateObject("AdaEDC.cPosNet")
'                Call W_EDCxPosNetCreditCard
'            End If
'        End If
'    Else
'          Me.Caption = tVB_TenderType
'          otmForm.Enabled = False
'          otbNum.SetFocus
'    End If
'End Sub

Private Sub W_EDCxHyperComCreditCard()
    'Dim oEdc As New HyperCom '  CreditCard.Verify
    With oEdc
        .SetNetAmout = otbAmt.Text
        .bShowThai = True
        '*Tee 53-04-22
        If tVB_StepAckOfEDC = "3" Then
            .SetAckAfApprove = True
        Else
            .SetAckAfApprove = False
        End If
        Call SP_INIbWriteFile(App.Path & "\History\EDCLog" & Format(Date, "yy-mm-dd") & ".log", "HyperCom", tVB_Running, "START :" & Format(Time, "HH:mm:ss"))
        If .EDC1CreateMessage Then
           Call SP_INIbWriteFile(App.Path & "\History\EDCLog" & Format(Date, "yy-mm-dd") & ".log", "HyperCom", tVB_Running, "EDC1CreateMessage :" & Format(Time, "HH:mm:ss"))
           If .EDC2SendMessage(tVB_CNEdcPort) Then
                Call SP_INIbWriteFile(App.Path & "\History\EDCLog" & Format(Date, "yy-mm-dd") & ".log", "HyperCom", tVB_Running, "EDC2SendMessage :" & Format(Time, "HH:mm:ss"))
                If .EDC3VerifyCreditCard Then
                    Call SP_INIbWriteFile(App.Path & "\History\EDCLog" & Format(Date, "yy-mm-dd") & ".log", "HyperCom", tVB_Running, "EDC3VerifyCreditCard :" & Format(Time, "HH:mm:ss"))
                    otbAcc.Text = .GetCardNumber
                    'otbBank.Text = .GetCardBankCode
                    'otbTrace.Text = .GetTraceCode
                    otbApvCode.Text = .GetApproveCode
                Else
                    MsgBox .GetMessage, vbInformation
                End If
                Call SP_INIbWriteFile(App.Path & "\History\EDCLog" & Format(Date, "yy-mm-dd") & ".log", "HyperCom", tVB_Running, "STOP :" & Format(Time, "HH:mm:ss"))
           Else
                MsgBox .GetMessage, vbInformation
           End If
        Else
           MsgBox .GetMessage, vbInformation
        End If
    End With
    
End Sub

Private Sub W_EDCxPosNetCreditCard()    '*Gon 50-03-21
'-----------------------------------------------------------
'   Cmt : PosNet
'   Create by : Jirawat 21-03-50
'-----------------------------------------------------------
    'Dim oEdc As New PosNet '  CreditCard.Verify
    With oEdc
        .SetNetAmout = otbAmt.Text
        .bShowThai = True
        Call SP_INIbWriteFile(App.Path & "\History\EDCLog" & Format(Date, "yy-mm-dd") & ".log", "PosNet", tVB_Running, "START :" & Format(Time, "HH:mm:ss"))
        If .EDC1CreateMessage Then
            Call SP_INIbWriteFile(App.Path & "\History\EDCLog" & Format(Date, "yy-mm-dd") & ".log", "PosNet", tVB_Running, "EDC1CreateMessage :" & Format(Time, "HH:mm:ss"))
           If .EDC2SendMessage(tVB_CNEdcPort) Then
                Call SP_INIbWriteFile(App.Path & "\History\EDCLog" & Format(Date, "yy-mm-dd") & ".log", "PosNet", tVB_Running, "EDC2SendMessage :" & Format(Time, "HH:mm:ss"))
                If .EDC3VerifyCreditCard Then
                    Call SP_INIbWriteFile(App.Path & "\History\EDCLog" & Format(Date, "yy-mm-dd") & ".log", "PosNet", tVB_Running, "EDC3VerifyCreditCard :" & Format(Time, "HH:mm:ss"))
                    otbAcc.Text = .GetCardNumber
                    'otbBank.Text = .GetCardBankCode
                    'otbTrace.Text = .GetTraceCode
                    otbApvCode.Text = .GetApproveCode
                Else
                    MsgBox .GetMessage, vbInformation
                End If
                Call SP_INIbWriteFile(App.Path & "\History\EDCLog" & Format(Date, "yy-mm-dd") & ".log", "PosNet", tVB_Running, "STOP :" & Format(Time, "HH:mm:ss"))
           Else
                MsgBox .GetMessage, vbInformation
           End If
        Else
           MsgBox .GetMessage, vbInformation
        End If
    End With
End Sub

Private Sub W_EDCxVerifoneCreditCard()    '*Tee 51-04-28  RQ0803-031
'-----------------------------------------------------------
'   Cmt : PosNet
'   Create by : Jirawat 21-03-50
'-----------------------------------------------------------
    'Dim oEdc As New Verifone  '  CreditCard.Verify
    With oEdc
        .SetNetAmout = otbAmt.Text
        .bShowThai = True
        Call SP_INIbWriteFile(App.Path & "\History\EDCLog" & Format(Date, "yy-mm-dd") & ".log", "Verifone", tVB_Running, "START :" & Format(Time, "HH:mm:ss"))
        If .EDC1CreateMessage And bW_ManaulMode = False Then
            Call SP_INIbWriteFile(App.Path & "\History\EDCLog" & Format(Date, "yy-mm-dd") & ".log", "Verifone", tVB_Running, "EDC1CreateMessage :" & Format(Time, "HH:mm:ss"))
           If .EDC2SendMessage(tVB_CNEdcPort) And bW_ManaulMode = False Then
                Call SP_INIbWriteFile(App.Path & "\History\EDCLog" & Format(Date, "yy-mm-dd") & ".log", "Verifone", tVB_Running, "EDC2SendMessage :" & Format(Time, "HH:mm:ss"))
                If .EDC3VerifyCreditCard And bW_ManaulMode = False Then
                    Call SP_INIbWriteFile(App.Path & "\History\EDCLog" & Format(Date, "yy-mm-dd") & ".log", "Verifone", tVB_Running, "EDC3VerifyCreditCard :" & Format(Time, "HH:mm:ss"))
                    otbAcc.Text = .GetCardNumber
                    'otbBank.Text = .GetCardBankCode
                    'otbTrace.Text = .GetTraceCode
                    otbApvCode.Text = .GetApproveCode
                Else
                    MsgBox .GetMessage, vbInformation
                End If
                Call SP_INIbWriteFile(App.Path & "\History\EDCLog" & Format(Date, "yy-mm-dd") & ".log", "Verifone", tVB_Running, "STOP :" & Format(Time, "HH:mm:ss"))
           Else
                MsgBox .GetMessage, vbInformation
           End If
        Else
           MsgBox .GetMessage, vbInformation
        End If
    End With
End Sub

Private Sub otmForm_Timer()
    nW_Timer = nW_Timer + 1
'    tW_EdcAction = "SALE"   '*Korn 55-10-10 Edit เพิ่ม 1 บรรทัด
    
    If bW_HandleESC Then
        If nW_Timer = nW_TimeESC Then
            bW_HandleESC = False
            '*Korn 55-10-10 Edit ปรับปรุง 8 บรรทัด
            ofa3Sec.Left = 120      'ซ่อน
            ofa3Sec.Top = 5400
            ofaSwipeCard.Left = 1600   'center
            ofaSwipeCard.Top = 600
            otcForm1(1).Enabled = False     'lock text box
            otcForm1(2).Enabled = False     'lock ปุ่ม
            otmForm.Enabled = False  'stop otmForm
            otmEDC.Enabled = True     'start otmEDC
            
        ElseIf nW_Timer < nW_TimeESC Then
            olaForm(8).Caption = nW_TimeESC - nW_Timer
            olaForm(9).Caption = nW_TimeESC - nW_Timer
        End If
    End If
End Sub

Private Sub otmEDC_Timer()
    If bW_ManaulMode = False Then
        'EDC/NETS job
        Select Case tVB_TypeOfEDC            '*Tee 51-04-28  RQ0803-031
            Case "1"        'Hypercom
                Call W_EDCxHyperComCreditCard
            Case "2"        'PosNet
                Call W_EDCxPosNetCreditCard
            Case "3"        'Verifone
                Call W_EDCxVerifoneCreditCard
            Case "4"        'EDC
'                Call W_EDCxVerifoneCreditCard
                '*Korn 55-10-08 Edit ตั้งชื่อ procedure ใหม่ (1 บรรทัด)
                '?? ค่า Comport จะต้องอ่านจาก TPOS, และ 40000 คือจำนวน 40 วินาทีที่รอเครื่อง อาจจะตั้งเป็น Option
'                Call W_EDCxMagic3CreditCard("COM2", 40000, tW_EdcAction)
                If tW_Comport = "" Then
                    otmEDC.Enabled = False
                    bW_HandleESC = True
                    Call Form_KeyDown(vbKeyEscape, 0)
                Else
                    Call SP_CRTxLogMallCardExtend(("Date:" & Date & " Time:" & Time$ & " tVB_Transaction:" & tVB_Transaction))  '*Bump 57-11-04 TK_Issue 27/10/2017 รับชำระด้วย Credit ทำให้ Transaction Type เปลี่ยนจาก 03 เป็น 31
                    Call W_EDCxMagic3CreditCard(tW_Comport, 100000, tW_EdcAction)
                End If
'                Call W_EDCxMagic3CreditCard(tW_Comport, 40000, tW_EdcAction)
            Case "5"        'NETS
                Call W_EDCxVerifoneCreditCard
        End Select
    End If
End Sub

Private Function W_PRCtRandom() As String
    Dim nRandomNumber As Integer
    Dim nMinValue As Integer
    Dim nMaxValue As Integer
    Dim tVal As String
    nMinValue = 0
    nMaxValue = 9
    
    'nRandomNumber = Int((nMaxValue - nMinValue + 1) * Rnd + nMinValue)
    nRandomNumber = Int(10 * Rnd)
    
    Select Case nRandomNumber
        Case 0: 'Sale $80,  redeem $20
            tVal = "R2000216XXXXXXXXXXXX03230412000000002000071011211248121404XXXX2201C3712542624604920380604063539020041084001300842121681689971825204DBSI53620440008000E800A000000004101001DBS MasterCard  66D7C0FC562A19745410MASTERCARD5501M5601L5706000005620600001363600000130000000020000000000020000000000000000001000000000000006406000001961011000000049826 /SWITCH UAT 851"
        Case 1: 'Sale $80, no redeem
            tVal = "R2000216XXXXXXXXXXXX03230412000000008000071011211249091404XXXX2201C3712998484643301380693736339020041084001300842121681689971825204DBSI53620440008000E800A000000004101001DBS MasterCard  99B120D0DCDF3E785410MASTERCARD5501M5601L5706000006620600001563600000150000000080000000000000000000000000000001000000000000006406000001961011000000049826 /SWITCH UAT 851"
        Case 2: 'Sale $80, redeem $80
             tVal = "R2000216XXXXXXXXXXXX03230412000000008000071011211250001404XXXX2201C3712750431560397380657992939020041084001300842121681689971825204DBSI53620440008000E800A000000004101001DBS MasterCard  86050DD67EA795985410MASTERCARD5501M5601L5706000007620600001763600000170000000080000000000080000000000000000001000000000000006406000001961011000000049826 /SWITCH UAT 851"
        Case 3:
            tVal = "R2000216XXXXXXXXXXXX03230412000000008000071011211249091404XXXX2201C3712998484643301380693736339020041084001300842121681689971825204DBSI53620440008000E800A000000004101001DBS MasterCard  99B120D0DCDF3E785410MASTERCARD5501M5601L57060000066206000015 "
'        Case 4:
'            tVal = "R2000216XXXXXXXXXXXX93200412000000000100071010151522551404XXXX2201M3712DEMODEMODEMO3806DEMO5139020041081234567842151681680000000005203DBS5410MASTERCARD5501M5601D57060000016206000047640600000296101100000004"
'        Case 5:
'            tVal = "R2000216XXXXXXXXXXXX93200412000000000100071010151522551404XXXX2201M3712DEMODEMODEMO3806DEMO5139020041081234567842151681680000000005203DBS5410MASTERCARD5501M5601D57060000016206000047640600000296101100000004"
'         Case 6:
'            tVal = "R20002165414963706119128041200000000010007100625102004140416012201C3712DEMODEMODEMO3806DEMO5139020041081234567842158006500000000005205TXTHB53620000008000E800A0000000041010  MASTERCARD      48CC4623D8BE7D5B5410MASTERCARD5501M5601I57060000016206000272634400000000233207645233247033500888888888888THB640600000396101101100004" 'HP Master"
'        Case 7:
'            tVal = "R20002165124920000045038041200000000010007100625102004140415032201C3712DEMODEMODEMO3806DEMO5139020041081234567842158006500000000005205TXTHB53620000008000E800A0000000041010  MASTERCARD      48CC4623D8BE7D5B5410MASTERCARD5501M5601T57060000016206000272634400000000233207645233247033500888888888888THB640600000396101101100004" 'TanachatMC
'        Case 8:
'             tVal = "R2000245#gTAjLBBtYd1K03zzhSNn+uqsofhcoaYUwNnFh2QjHX4=0412000000000001071009041054581425#0OjPrIAyg8sYFRNZLP/DUg==2201C371276593321626238063777433902004108160688884211168168188885203DBS53620000008000E800A000000004101001DBS MasterCard  56A6F475EE6006C65410MASTERCARD5501M5601D570600000162060000386406000008961011000000149826 /RAJESH CHAURASIA        "
'        Case 9:
'            tVal = "R20002165124920000045038041200000000010007100625102004140415032201C3712DEMODEMODEMO3806DEMO5139020041081234567842158006500000000005205TXTHB53620000008000E800A0000000041010  MASTERCARD      48CC4623D8BE7D5B5410MASTERCARD5501M5601I57060000016206000272634400000000233207645233247033500888888888888THB640600000396101101100004" 'HP Thanachat
        Case Else 'Sale $80, Loyalty host is down. (Exception scenario)
            tVal = "R2000216XXXXXXXXXXXX03230412000000002000071011211248121404XXXX2201C3712542624604920380604063539020041084001300842121681689971825204DBSI53620440008000E800A000000004101001DBS MasterCard  66D7C0FC562A19745410MASTERCARD5501M5601L5706000005620600001363600000130000000020000000000020000000000000000001000000000000006406000001961011000000049826 /SWITCH UAT 851"
    End Select
    W_PRCtRandom = tVal
End Function
