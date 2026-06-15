VERSION 5.00
Object = "{D6810AA0-45D7-11CF-88F3-02608CA3BE43}#1.0#0"; "Msr.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.1#0"; "Codejock.SuiteCtrls.Unicode.v11.1.3.ocx"
Object = "{9D6F2CC0-07CE-11CF-A800-0020AF9EBB73}#1.0#0"; "Scanner.ocx"
Begin VB.Form wCNDiscMall 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "The MALL Card;klg_Title"
   ClientHeight    =   4215
   ClientLeft      =   1155
   ClientTop       =   2940
   ClientWidth     =   7215
   ControlBox      =   0   'False
   Icon            =   "wCNDiscMall.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   7215
   Tag             =   "2;"
   Visible         =   0   'False
   Begin VB.Timer otmEDC 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   600
      Top             =   3720
   End
   Begin VB.Timer otmForm 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   120
      Top             =   3720
   End
   Begin VB.Frame ofa3Sec 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   2535
      Left            =   120
      TabIndex        =   19
      Top             =   4680
      Width           =   3495
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
         TabIndex        =   21
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label olaForm 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Press ESC for manual or  wait for swipe card."
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
         TabIndex        =   20
         Top             =   480
         Width           =   2775
      End
   End
   Begin VB.Frame ofaSwipeCard 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   2535
      Left            =   2904
      TabIndex        =   17
      Top             =   4680
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
         TabIndex        =   18
         Top             =   960
         Width           =   2535
      End
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
         Left            =   2520
         MaxLength       =   2
         TabIndex        =   1
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox orbYear 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3480
         MaxLength       =   2
         TabIndex        =   2
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox otbCusID 
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
         Width           =   4335
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
         Index           =   7
         Left            =   4200
         TabIndex        =   12
         Tag             =   "2;"
         Top             =   1320
         Width           =   2775
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Discount Rate %:;klg_DiscRate"
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
         TabIndex        =   10
         Tag             =   "2;รหัสพนักงานขาย;Cashier ID"
         Top             =   2520
         Width           =   2115
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Type:;klg_CusType"
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
         TabIndex        =   9
         Tag             =   "2;รหัสพนักงานขาย;Cashier ID"
         Top             =   1920
         Width           =   2115
      End
      Begin VB.Label olaRate 
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
         TabIndex        =   8
         Top             =   2520
         Width           =   4335
      End
      Begin VB.Label olaName 
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
         Top             =   1920
         Width           =   4335
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
         Index           =   1
         Left            =   240
         TabIndex        =   6
         Tag             =   "2;รหัสพนักงานขาย;Cashier ID"
         Top             =   1320
         Width           =   2130
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
         Index           =   0
         Left            =   240
         TabIndex        =   5
         Tag             =   "2;"
         Top             =   720
         Width           =   2085
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
         Index           =   2
         Left            =   3240
         TabIndex        =   4
         Top             =   1320
         Width           =   210
      End
   End
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   1116
      Index           =   2
      Left            =   0
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   3240
      Width           =   7140
      _Version        =   720897
      _ExtentX        =   12594
      _ExtentY        =   1968
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
      PaintManager.Position=   2
      Begin MSRLib.MSR OPOSMSR1 
         Left            =   0
         Top             =   0
         _Version        =   65536
         _ExtentX        =   741
         _ExtentY        =   741
         _StockProps     =   0
      End
      Begin SCANNERLib.Scanner oScanner 
         Left            =   600
         Top             =   120
         _Version        =   65536
         _ExtentX        =   741
         _ExtentY        =   741
         _StockProps     =   0
      End
      Begin VB.Timer otmTime 
         Enabled         =   0   'False
         Interval        =   5000
         Left            =   960
         Top             =   480
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
         Height          =   510
         Left            =   5640
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   150
         Width           =   1455
      End
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
         Left            =   2520
         TabIndex        =   15
         Tag             =   "2;"
         Top             =   150
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
         Left            =   4080
         TabIndex        =   14
         Tag             =   "2;"
         Top             =   150
         Width           =   1455
      End
      Begin VB.CommandButton ocmFind 
         Caption         =   "&Find;klg_Find"
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
         Left            =   960
         TabIndex        =   13
         Tag             =   "2;"
         Top             =   150
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
         Left            =   0
         TabIndex        =   22
         Top             =   0
         Visible         =   0   'False
         Width           =   735
      End
   End
End
Attribute VB_Name = "wCNDiscMall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private orsMall As ADODB.Recordset
Private bW_FoundMallCard  As Boolean
Private orsMem As ADODB.Recordset
Private tW_ChkExp  As String
Private tW_ChkDpt  As String
Private bW_KeyEnter As Boolean '***KT 51-12-01 ป้องกันการเคาะ Enter เบิ้ลทำใหเเกิดส่วนลดเบิ้ล

Private aW_PointDpt() As POINTDPT
Private aW_PointMall() As POINTDPT
Private aW_PointSKU() As POINTSKU

Private cW_AmtMax As Double
Private cW_AmtMin As Double
Private tW_MallRegAlw  As String
Private tW_MallProAlw  As String
Private dW_MallStartDate  As Date
Private dW_MallEndDate  As Date
Private tW_StaMallFrmAmt   As String
Private tW_StaDiscMallAlw  As String
Private tW_IsChkDigitMallCard As String '1: allow check 0 : not allow check other : not allow check
Private tW_CardName As String '*KT 52-11-26 PH2-ML-SF-170 ใน Ticket ไม่พิมพ์ชื่อพนักงานที่ใช้ส่วนลด Employee discount
Private tW_CardExp As String '*Tao 60-06-02 Card Expire (MMYY)
Public nW_TrackNo  As Long '*KT 55-07-10 Trackdata Number of MSR

Private bW_EnableMSR  As Boolean '*Bump 58-04-18 CommSheet TK-ISS3000-558.02_22042015.xls

Private bW_ManaulMode As Boolean
Private bW_HandleESC As Boolean
Private nW_TimeESC As Long
Private nW_TimeExpireEDC As Long
Private nW_Timer As Long
Private oEdc As Object
Private tW_Comport As String
Private tW_EdcAction As String

Private tW_EDCRun As String
Private nW_CaseERROR As Integer
Private bW_ChkESC As Boolean

Private bW_EdcCheck As Boolean
Private bW_ClickEDC  As Boolean '*Bump 58-04-30 กดปุ่ม EDC บนหน้าจอ

Private Function W_CHKbAccByID(ByRef ptRefCardType As String) As Boolean
'-------------------------------------------------------------------------------------------------
'Call:
'Cmt: หา Card Type ที่อยู่ใน master เมื่อรูดบัตรผ่าน EDC     '*Eaw 56-05-13 CommSheet TK-ISS3000-167-ST 00.00.02.xls
'-------------------------------------------------------------------------------------------------
Dim tSql As String
Dim orsTemp As ADODB.Recordset
On Error GoTo ErrorHandles
    W_CHKbAccByID = False
'    tSql = W_PRCxStrSQL(Trim(otbAcc.Text))
'    Call SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine, , , , , True, True) 'read master
'    If Not (orsTemp.BOF Or orsTemp.EOF) Then
'        W_CHKbAccByID = True
'        ptRefCardType = SP_FEDtChkString(orsTemp, "FTCdcCreditType")  '*Eaw 56-06-13  CommSheet TK-ISS3000-086 00.00.08.xls
'        tVB_CrdCardType = SP_FEDtChkString(orsTemp, "FTCdcType") '*Eaw 56-09-10
'        tVB_TenderName = SP_FEDtChkString(orsTemp, "FTCdcCreditType") '*Eaw 56-09-13 ให้เปลี่ยนเป็น FTCdcCreditType แทน FTCdcDescCredit
'        tVB_CreditGroup = SP_FEDtChkString(orsTemp, "FTCdcCreditGroup")         '*Ao 2013/11/05
'        tVB_CrdStaBPSale = SP_FEDtChkString(orsTemp, "FTCdcStaBPSale") '*Eaw 56-11-06
'        tVB_CrdStaBPReturn = SP_FEDtChkString(orsTemp, "FTCdcStaBPReturn")
'        Call SP_CRTxLogMallCardEvent("Credit", "EDCResponse", "W_CHKbAccByID=True=", tSql)
'    Else
        W_CHKbAccByID = False
        ptRefCardType = ""   '*Eaw 56-06-13  CommSheet TK-ISS3000-086 00.00.08.xls
        Call SP_CRTxLogMallCardEvent("Credit", "EDCResponse", "W_CHKbAccByID=False=", tSql)
'    End If
    Exit Function
ErrorHandles:
    W_CHKbAccByID = False
    ptRefCardType = ""  '*Eaw 56-06-13  CommSheet TK-ISS3000-086 00.00.08.xls
    Call SP_CRTxLogMallCardEvent("Credit", "EDCResponse", "W_CHKbAccByID=ErrorHandles=", tSql)
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
    tSql = "SELECT DISTINCT TOP 1 FTTdmCode,FTCdcCreditType,FTCdcDescCredit,FTCdcType, FTCdcStaBPSale,FTCdcStaBPReturn,FTRemark as FTFreeRangCard,FTCdcCreditGroup"
    tSql = tSql & vbCrLf & "FROM TCNMCreditCard WHERE ISNULL(FTRemark,'') = 'Y'"
    tSql = tSql & vbCrLf & "AND FTCdcType = '" & ptCardType & "'"
    tSql = tSql & vbCrLf & "AND FTTdmCode = '" & ptTenderCode & "'"
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
        'aTag[0/1,17]=>64-000002                             หมายเลข Batch
        'aTag[0/1,18]=>96-1100000004                    หมายเลขอ้างอิงในการพิมพ์
'____________________________________________________________________________________
    Dim tTemp$, tCut$, nLen%, n%, tMsg$
    Dim aTag() As String
    Dim tExpired As String
    Dim tCardType As String
    Dim bChk  As Boolean
    Dim tRefCardType As String  '*Eaw 56-06-13 เก็บค่า Card Type ที่เอามาจาก master เช่น TKAmex,CUP,Visa
    Dim tChkCardType As String '*Eaw 56-07-08 เก็บค่ากรณีเป็น ezLink เพื่อไม่เช็ควันหมดอายุบัตร
    Dim tTagDetail As String '*Eaw 56-07-24
    Dim tCardTag55 As String
    Dim tTag39 As String '*Bump 58-05-11
    Dim tTmpCard As String
    Dim tTmpDate As String
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
        tTemp = Right(tTemp, Len(tTemp) - nLen)   'ตัด temp ออก
        aTag(1, n) = tCut
        '*Tao 62-12-10 ปิดไว้ เนื่องจากได้ข้อมูลครบ (ไม่ได้ Trim ptRet จะไม่มี error5)
'        If tTag39 = "39" Then '*Bump 58-05-11 กรณีกดปุ่ม EDC เช็คถึง tag39 ก็พอป้องกัน error5
'            Exit Do
'        End If
    Loop

    Call SP_CRTxLogMallCardEvent("C209", "EDCResponse", "SplitResponseLastTagID=", n & ":" & tCut)
    Call SP_CRTxLogMallCardEvent("C209", "EDCResponse", "SplitResponseLastText=", tTemp)
    'เช็คว่ามีค่ารหัสบัตรหรือไม่ (02:มีรหัสบัตร)
    If aTag(0, 1) <> "02" Then      'ถ้าไม่มีรหัสบัตร แสดงว่าอาจมีปัญหา
        If aTag(0, 1) = "39" Then '*Bump 58-04-30
            tMsg = SP_GETtMsgErrEDC(tCut) '
            Call SP_MSGnShowing("tms_CN005065" & ";" & tMsg, nCS_Warning)
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
            
            Call SP_MSGnShowing("tms_CN005065" & ";" & tMsg, nCS_Warning)
            Unload Me
            Exit Sub
        End If
    End If
    
    'แสดงว่ารูดบัตร ผ่าน
    'ซ่อน Swipe card แต่ไม่ปลด lock ปุ่ม ok, cancel
    ofaSwipeCard.Left = 2900        'ซ่อน
    ofaSwipeCard.Top = 5400
    
    'นำค่าลงในหน้าจอ
    '*TU 2020-09-09 CR_Encrypted Creditcard
    '---------------------------------------------------------------------------------------------------------
    tTmpCard = ""
    If Left(SP_EDCxCheckTag(aTag, n, "02"), 1) = "#" Then
       tTmpCard = SP_EDCtDecrypt(Trim(SP_EDCxCheckTag(aTag, n, "02")))
        If tTmpCard <> "" Then
            otbCusID.Text = Right(tTmpCard, Len(tTmpCard) - 2)
        Else
            otbCusID.Text = ""
        End If
    Else
        otbCusID.Text = SP_EDCxCheckTag(aTag, n, "02")
    End If
    
    tTmpDate = ""
    If Left(SP_EDCxCheckTag(aTag, n, "14"), 1) = "#" Then
        tTmpDate = SP_EDCtDecrypt(Trim(SP_EDCxCheckTag(aTag, n, "14")))
        tExpired = Right(tTmpDate, Len(tTmpDate) - 2)
    Else
        tExpired = SP_EDCxCheckTag(aTag, n, "14")
    End If
    
'    otbCusID.text = SP_EDCxCheckTag(aTag, n, "02")
'    tExpired = SP_EDCxCheckTag(aTag, n, "14")
'---------------------------------------------------------------------------------------------------------
    orbMonth.Text = Right(tExpired, 2)
    orbYear.Text = Left(tExpired, 2)
    
    '02           card Number
    '14  Expire Date
    '22  Entry Mode
    '39  Response code
    '98  card 's Holder name
    'tTagDetail = "02" & SP_EDCxCheckTag(aTag, n, "02") '*Tao 61-09-18 ของเดิม
    tTagDetail = "02" & SP_PRCtStrMarkX(SP_EDCxCheckTag(aTag, n, "02")) '*Tao 61-09-18 Mark X ก่อน
    tTagDetail = tTagDetail & ",14" & SP_EDCxCheckTag(aTag, n, "14")
    tTagDetail = tTagDetail & ",14" & SP_EDCxCheckTag(aTag, n, "22")
    tTagDetail = tTagDetail & ",14" & SP_EDCxCheckTag(aTag, n, "39")
    tTagDetail = tTagDetail & ",14" & SP_EDCxCheckTag(aTag, n, "98")
'    tTagDetail = tTagDetail & ",38" & SP_EDCxCheckTag(aTag, n, "38")
'    tTagDetail = tTagDetail & ",56" & SP_EDCxCheckTag(aTag, n, "56")
'    tTagDetail = tTagDetail & ",54" & SP_EDCxCheckTag(aTag, n, "54")
'    tTagDetail = tTagDetail & ",55" & SP_EDCxCheckTag(aTag, n, "55")
    If Len(Trim(tTagDetail)) > 255 Then
        tVB_EDCDetail = Left(tTagDetail, 255)
    Else
        tVB_EDCDetail = tTagDetail
    End If

    Call SP_CRTxLogMallCardEvent("C209", "EDCResponse", "Data=", tVB_EDCDetail)
    tVB_EDCDetail = ""
    '-------------------------------------------------------------------------------------------
    'เรียก ok
    bW_EdcCheck = False  'เพื่อไม่ให้ทำการเช็คเลขที่บัตรอีก เมื่อกดปุ่ม Ok(แบบ Manual)
    tVB_DiscMallCardID = otbCusID.Text
    If Not W_PRCxFindMallCardByID(tVB_DiscMallCardID) Then
        SleepEx 1200, 0
        bW_FoundMallCard = False
        tW_CardName = "" '*KT 52-11-26 PH2-ML-SF-170 ใน Ticket ไม่พิมพ์ชื่อพนักงานที่ใช้ส่วนลด Employee discount
        otbCusID.Text = ""                      '*AO51-03-20    Comm.Sheet  SF-108
        orbMonth.Text = ""
        orbYear.Text = ""
        olaName.Caption = ""
        olaRate.Caption = ""
        Call SP_OTBxSelText(otbCusID)
        'Call SP_CRTxLogMallCardEvent("Keyboard", "CheckMember", "Card ID:", Trim(otbCusID.Text)) '*Tao 61-09-18 ของเดิม
        Call SP_CRTxLogMallCardEvent("Keyboard", "CheckMember", "Card ID:", SP_PRCtStrMarkX(otbCusID.Text)) '*Tao 61-09-18 Mark X ก่อน
        Unload Me
        Exit Sub
    End If
    
    '*Bump 58-04-29 ใช้แทนข้างล่าง
      dVB_CrdDateExp = DateSerial(2000 + Val(Format(orbYear.Text)), Val(Format(orbMonth.Text)) + 1, 1)  '*Eaw 56-06-21CommSheet TK-ISS3000-238-ST.xls เพิ่มการตรวจสอบวันหมดอายุ
        If CDate(Format(dVB_CrdDateExp, "yyyy/MM/dd") < CDate(Format(dVB_CurrentDate, "yyyy/MM/dd"))) Then
            'Call SP_MSGnShowing("Member card is expired", nCS_Warning)
            MsgBox "Member card is expired", vbExclamation
            SleepEx 1200, 0
            otbCusID.Text = ""
            orbMonth.Text = ""
            orbYear.Text = ""
            olaName.Caption = ""
            olaRate.Caption = ""
            Unload Me
            Exit Sub
        End If

    Call ocmOK_Click            'สมบูรณ์แล้ว จะ unload โดยอัติโนมัติ

    Exit Sub
ErrorHandles:
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
    Call SP_MSGxShowError(Err.Number, Err.Description & "NOT PROPERLY from EDC Response")
    Unload Me
End Sub

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
        Case Else: tMsg = "No Response from EDC"
    End Select
    
    W_EDCtReturnCase = tMsg
    
End Function

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
'            tSend = tSend & "5706"        'tag 57 ตาราง tag 1.6.4 จำนวน 06 หลัก, fix
'            tSend = tSend & tW_EDCRun
    End Select
    
    W_EDCtGetString2Send = tSend
    'Call SP_CRTxLogMallCardEvent("C20", "EDCSendString", "SendStringValue=", tSend) '*Eaw 56-10-07 เก็บ Log
    Call SP_CRTxLogMallCardExtend(("Date:" & Date & " Time:" & Time$ & " C209:POSRequest: " & tSend))
End Function

Private Sub W_PRCxloadEDC()
'_______________________________________________
'*NOT 58-04-16 เพิ่มปุม EDC
'
'
'_______________________________________________
tVB_TypeOfEDC = "4"
If tVB_TypeOfEDC = "4" Then
        Dim oEdc As Object
        Dim tTime As String
        olaForm(8).Caption = ""
        olaForm(9).Caption = ""
'        otbAcc.Locked = True
'        otbAcc.BackColor = nVB_CNClrBLock
        nW_Timer = 0
        bW_HandleESC = True
        bW_ManaulMode = False
        ocmCancel.Cancel = False
'        bW_ChkocmEDC = True
'        Call SP_OBJxSetFocus(otbAcc)

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
'                        otbAcc.Locked = False
'                        otbAcc.BackColor = nVB_CNClrBUnLock
                        bW_HandleESC = False
                        bW_ManaulMode = True
                        otmEDC.Enabled = False
                        otmForm.Enabled = False
'                        Call SP_OBJxSetFocus(otbAcc)
                        ocmEDC.Enabled = False
                    Case Else
                        olaForm(8).Caption = nW_TimeESC
                        olaForm(9).Caption = nW_TimeESC
                        ofa3Sec.Left = 1600     'center
                        ofa3Sec.Top = 600
                        ofaSwipeCard.Left = 2900    'ซ่อน
                        ofaSwipeCard.Top = 5400
                        otcForm1(0).Enabled = False     'lock text box
                        otcForm1(2).Enabled = False     'lock ปุ่ม
                        otmForm.Enabled = True
'                        Call otmForm_Timer
                End Select
        End If
    End If
End Sub


'Private Function W_PRCxStrSQL(ByVal ptAccNum As String) As String
''--------------------------------------------------------------------------------------------
''Call:
''Cmt: '*Eaw 56-09-04 Get string to check TC card
''----------------------------------------------------------------------------------------------
'Dim tSql As String
'Dim nLen As Long
'    nLen = Len(ptAccNum)
'    tSql = "SELECT TOP 1 TCNMCreditCard.*"
'    tSql = tSql & " ,TCNMCreditCardRange.FTCcrRangeFrom,TCNMCreditCardRange.FTCcrRangeTo"
'    tSql = tSql & " FROM TCNMCreditCard INNER JOIN"
'    tSql = tSql & " TCNMCreditCardRange ON TCNMCreditCard.FTTdmCode = TCNMCreditCardRange.FTTdmCode AND"
'    tSql = tSql & " TCNMCreditCard.FTCdcCreditType = TCNMCreditCardRange.FTCdcCreditType"
'    tSql = tSql & " WHERE TCNMCreditCard.FTTdmCode = '" & tVB_TenderCode & "'"
'    If tVB_TenderCode <> "T017" Then
'        tSql = tSql & " AND ISNULL(TCNMCreditCard.FTRemark,'N') = 'N'"
'    End If
'    Select Case (tVB_TenderCode)
'        Case "T017"
'            tSql = tSql & vbCrLf & " AND (LEFT( '" & ptAccNum & "' ,TCNMCreditCard.FNCdcDigitChk) BETWEEN LEFT(TCNMCreditCardRange.FTCcrRangeFrom,TCNMCreditCard.FNCdcDigitChk)  AND LEFT(TCNMCreditCardRange.FTCcrRangeTo,TCNMCreditCard.FNCdcDigitChk)"
'            tSql = tSql & vbCrLf & " OR (LEN(  '" & ptAccNum & "') = " & nLen & " AND '" & ptAccNum & "' BETWEEN RIGHT(TCNMCreditCardRange.FTCcrRangeFrom, " & nLen & ") AND RIGHT(TCNMCreditCardRange.FTCcrRangeTo, " & nLen & ")))"
'            tSql = tSql & vbCrLf & " AND (LEN(RIGHT(RTRIM(TCNMCreditCardRange.FTCcrRangeFrom), " & nLen & ")) = " & nLen & ")"
'        Case "T002", "T003"
'            tSql = tSql & " AND (LEFT('" & ptAccNum & "' ,TCNMCreditCard.FNCdcDigitChk) BETWEEN LEFT(TCNMCreditCardRange.FTCcrRangeFrom,TCNMCreditCard.FNCdcDigitChk)  AND LEFT(TCNMCreditCardRange.FTCcrRangeTo,TCNMCreditCard.FNCdcDigitChk) ) "
'            tSql = tSql & " AND LEN('" & ptAccNum & "') = TCNMCreditCard.FNCdcDigits"
'    End Select
'
'    tSql = tSql & " ORDER BY TCNMCreditCard.FNCdcDigitChk DESC"
'    W_PRCxStrSQL = tSql
'
'End Function
Private Function W_PRCbFindTCCardByNum(ByVal ptTCCardNo As String) As Boolean
'--------------------------------------------------------------------------------------------
'Call:
'Cmt: '*Eaw 56-09-04 ตรวจสอบ TC Card ว่ามีอยู่ใน Master หรือไม่
'----------------------------------------------------------------------------------------------
Dim tSql As String
Dim orsTemp As ADODB.Recordset
On Error GoTo ErrorHandles
    W_PRCbFindTCCardByNum = False
'    tSql = W_PRCxStrSQL(ptTCCardNo)
    tSql = "STP_PRCxFindMallCardByID @ptAccNum= '" & ptTCCardNo & "'" '*Eaw 56-09-27
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

Private Function W_PRCxSpcMem91()
'______________________________________
'   Call:
'   Cmt: '*Eaw 56-01-05 ใช้เรียก  Hot Key ของ Discount 91
'______________________________________
    Dim tDate$
    Dim tMonth$
    Dim tYear$
    Const tCusID = "91"
    Const KeyAscii = 13
    tDate = Format(Year(Date), "0000")
    tYear = Right(tDate, 2)
    tMonth = Format(Month(Date), "00")
    otbCusID.Text = tCusID
    Call otbCusID_KeyPress(KeyAscii)
    orbMonth.Text = tMonth
    orbYear.Text = tYear
    ocmFind.Enabled = False '*TON 58-09-30
End Function

Private Sub W_PRCxNOTFoundMallCard()
'    otbCusID.Text = ""
'    orbMonth.Text = ""
'    orbYear.Text = ""
    olaName.Caption = ""
    olaRate.Caption = ""
    tVB_StaMemPointAlw = "N" '*KT 52-02-17 member point
    tVB_MallCardGroup = ""
    cVB_DiscMallGrp = 0
    cVB_DiscMallRanking = 0
    tVB_DiscMallCardKey = ""
    tVB_DiscMallCardType = ""
    cVB_DiscMallRate = 0
    cW_AmtMax = 0
    cW_AmtMin = 0
    tW_MallRegAlw = "N"
    tW_MallProAlw = "N"
    dW_MallStartDate = Date
    dW_MallEndDate = Date
    tW_StaMallFrmAmt = "N"
    tW_StaDiscMallAlw = "N"
    tW_IsChkDigitMallCard = "0"
    tVB_MemEvnCode = "" '*KT 54-11-18 V.4.0.5  DCR_5411003 PRJ09010TheMall-R4-SE-External Spec.Upload Bonus Points-00.04.00-FSBT.doc
    tVB_IsMemJIS = "N"
    tW_CardExp = "" '*Tao 60-06-02
End Sub

Private Function W_PRCxFindMallCardByID(ByVal ptCusID$) As Boolean
'--------------------------------------
'Call:
'Cmt: ค้นหา Format ของเลขที่บัตร Mall Card อยู่ในช่วงใด
'*KT 52-02-12 New Requirment
'1.1 เพิ่มการตรวจสอบหมายแลขบัตร โดยเพิ่ม การระบุจำนวน Digits และ จำนวน Digits  Check
'Range และแต่ละ Card Type ยังสามารถระบุ Range ของ Card ได้หลาย Range
'*KT 52-09-30 PH2-ML-SF-149 หมายเลขบัตร Mall Card ไม่ตรงกัน ระหว่าง M-Price กับ รูด Mall Card (Mall8) เพิ่ม FTMcdIsChkDigit สำหรับตรวจสอบ MOD 10
'--------------------------------------
    Dim tSql$
On Error GoTo ErrHandle
     Call SP_STAxMouse(Me, True)
    W_PRCxFindMallCardByID = False
    
'    tSql = "SELECT  TOP 1 MM.FNMcdDigits,MM.FNMcdDigitChk,MM.FTMcdCode,MM.FTMcdCardType,MM.FTMcdDesc1"
'    tSql = tSql & vbCrLf & ",MM.FTMcdDesc2,MM.FCMcdAmtRank,MM.FCMcdRate,MM.FTMcdStaRang,MM.FTMcdFrom,MM.FTMcdTo"
'    tSql = tSql & vbCrLf & ",MM.FTMcdStaChkExp,MM.FTMcdStaDeptCode,MM.FTMcdGroup,MM.FTMcdAlwPoint,MM.FCMcdAmtPoint"
'    tSql = tSql & vbCrLf & ",MM.FDMcdStartDatePoint,MM.FDMcdEndDatePoint,MM.FTMcdRegPoint,MM.FTMcdPromoPoint"
'    tSql = tSql & vbCrLf & ",MM.FCMcdAmtMaxPoint,MM.FCMcdAmtMinPoint,MM.FTMcdStaDeptCodePoint ,MM.FCMcdAmtMax"
'    tSql = tSql & vbCrLf & ",MM.FCMcdAmtMin,MM.FTMcdReg,MM.FTMcdPromo,MM.FTMcdStaAmtDisc,MM.FDMcdStartDate"
'    tSql = tSql & vbCrLf & ",MM.FDMcdEndDate,MM.FTMcdAlwPointAll,MM.FTMcdStaSubAllPoint,MM.FTMcdStaTier"
'    tSql = tSql & vbCrLf & ",MM.FTMcdStaTime,MM.FTMcdStaAddSku,MM.FTMcdIsChkDigit,MM.FTCdcStaNonCard"
'    tSql = tSql & vbCrLf & ",MM.FTEvnCode,MM.FTMcdIsJIS FROM TCNMMallCardMtn MM INNER JOIN TCNMMallCardRange MR"
'    tSql = tSql & vbCrLf & "ON MM.FTMcdCode = MR.FTMcdCode"
'    tSql = tSql & vbCrLf & "WHERE LEFT(ptCusID,MM.FNMcdDigitChk) BETWEEN LEFT(MR.FTMcrRangeFrom,MM.FNMcdDigitChk) AND LEFT(MR.FTMcrRangeTo,MM.FNMcdDigitChk)"
'    tSql = tSql & vbCrLf & "AND LEN(ptCusID) = MM.FNMcdDigits"
'    tSql = tSql & vbCrLf & "AND ISNULL(MM.FTCdcStaNonCard,'N') <> 'Y'"
'    tSql = tSql & vbCrLf & "ORDER BY MM.FNMcdDigitChk DESC"
    
    tSql = "STP_PRCxFindMallCardByID @ptAccNum= '" & ptCusID & "'"
    Call SP_SQLvExecute(tSql, orsMall, oDB.ocnOnLine, , , , , True, True) 'read master
    '*Eaw 56-10-02 CommSheet TK-ISS3000-368.xls ถ้าระบบ Offline อยู่ถ้ารูด Member ระบบจะช้า เลยแยกเคสกันระหว่าง Online กับ Offline
'    If oDB.SQLbAvailableServer(oDB.OffServerName, oDB.OnComIP) Then
'            Call SP_SQLvExecute(tSql, orsMall, oDB.ocnOnLine, , , , , True, True) 'Online
'    Else
'            Call SP_SQLvExecute(tSql, orsMall, oDB.ocnOffLine, , , , , False, False) 'Offline
'    End If
    If Not (orsMall.BOF Or orsMall.EOF) Then
        tVB_IsMemJIS = SP_FEDtChkString(orsMall, "FTMcdIsJIS")
        If nW_TrackNo = 4 Then
            If UCase(tVB_IsMemJIS) = UCase("N") Then
                Call SP_MSGnShowing("tms_CN005238", nCS_Warning)              '*Ao 2013/02/04
                Call W_PRCxNOTFoundMallCard
                Exit Function
            End If
        Else
            If (nW_TrackNo = 2) And (tVB_CNTrnDiscType = "1809") Then
                If UCase(tVB_IsMemJIS) = UCase("Y") Then
                    Call SP_MSGnShowing("tms_CN005068", nCS_Warning)
                    Call W_PRCxNOTFoundMallCard
                    Exit Function
                End If
            End If
        End If
        W_PRCxFindMallCardByID = True
        olaName.Caption = SP_FEDtChkString(orsMall, "FTMcdDesc1")
        tW_StaDiscMallAlw = SP_FEDtChkString(orsMall, "FTMcdStaRang")
        If tW_StaDiscMallAlw = "Y" Then
            cVB_DiscMallRate = SP_FEDvChkNumeric(orsMall, "FCMcdRate")              '*Ao 2013/03/12
            olaRate.Caption = cVB_DiscMallRate
        Else
            cVB_DiscMallRate = 0
            olaRate.Caption = cVB_DiscMallRate
        End If
'        cVB_DiscMallRate = SP_FEDvChkNumeric(orsMall, "FCMcdRate")
'        olaRate.Caption = cVB_DiscMallRate
        tVB_DiscMallCardType = SP_FEDtChkString(orsMall, "FTMcdCardType")
        tVB_DiscMallCardKey = SP_FEDtChkString(orsMall, "FTMcdCode")
        tW_ChkExp = SP_FEDtChkString(orsMall, "FTMcdStaChkExp")
        tW_ChkDpt = SP_FEDtChkString(orsMall, "FTMcdStaDeptCode")
        cVB_DiscMallRanking = SP_FEDvChkNumeric(orsMall, "FCMcdAmtRank")
        cVB_DiscMallGrp = SP_FEDvChkNumeric(orsMall, "FCMcdAmtRank")
        tVB_MallCardGroup = SP_FEDtChkString(orsMall, "FTMcdGroup")
        tVB_StaMemPointAlw = SP_FEDtChkString(orsMall, "FTMcdAlwPointAll") '*KT 52-02-17 member point
        cW_AmtMax = SP_FEDvChkNumeric(orsMall, "FCMcdAmtMax")
        cW_AmtMin = SP_FEDvChkNumeric(orsMall, "FCMcdAmtMin")
        tW_MallRegAlw = SP_FEDtChkString(orsMall, "FTMcdReg")
        tW_MallProAlw = SP_FEDtChkString(orsMall, "FTMcdPromo")
        dW_MallStartDate = SP_FEDdChkDate(orsMall, "FDMcdStartDate")
        dW_MallEndDate = SP_FEDdChkDate(orsMall, "FDMcdEndDate")
        tW_StaMallFrmAmt = SP_FEDtChkString(orsMall, "FTMcdStaAmtDisc")
'        tW_StaDiscMallAlw = SP_FEDtChkString(orsMall, "FTMcdStaRang")
        tW_IsChkDigitMallCard = SP_FEDtChkString(orsMall, "FTMcdIsChkDigit")
        tVB_MemEvnCode = SP_FEDtChkString(orsMall, "FTEvnCode") '*KT 54-11-18 V.4.0.5  DCR_5411003 PRJ09010TheMall-R4-SE-External Spec.Upload Bonus Points-00.04.00-FSBT.doc
        '*Tao 60-06-02 เก็บ Expire Date (ไม่ใช้ SP_FEDdChkDate เพราะอาจจะได้ System Date ออกมา)
        If IsNull(orsMall.Fields("FDMcdEndDate")) Then
            tW_CardExp = ""
        Else
            If CDate(orsMall.Fields("FDMcdEndDate")) >= Date Then
                tW_CardExp = Format(CDate(orsMall.Fields("FDMcdEndDate")), "MMYY")
            Else
                tW_CardExp = ""
            End If
        End If
        '=========================================================================
'        tVB_IsMemJIS = SP_FEDtChkString(orsMall, "FTMcdIsJIS")      '*Ao 2013/02/04
    Else
        Call W_PRCxNOTFoundMallCard    '*Ao 2013/02/04
        If bW_ClickEDC Then '*Bump 58-05-02
            '*กรณีกดปุ่ม EDC บนหน้าจอที่ส่ง C209
            'MsgBox "Invalid Card Data from EDC", tCS_CNMsgWarn, tVB_CNAppTitle
            Call SP_MSGnShowing("tms_CN008042", nCS_Warning) '*Bump 59-05-03 CommSheet TK-ISS3000-617 Wrong warning message.xls
        Else
            Call SP_MSGnShowing("tms_CN005068", nCS_Warning)
        End If
        'Call SP_MSGnShowing("tms_CN005068", nCS_Error)
        otbCusID.Text = ""
        orbMonth.Text = ""
        orbYear.Text = ""
'        olaName.Caption = ""
'        olaRate.Caption = ""
'        tVB_StaMemPointAlw = "N" '*KT 52-02-17 member point
'        tVB_MallCardGroup = ""
'        cVB_DiscMallGrp = 0
'        cVB_DiscMallRanking = 0
'        tVB_DiscMallCardKey = ""
'        tVB_DiscMallCardType = ""
'        cVB_DiscMallRate = 0
'        cW_AmtMax = 0
'        cW_AmtMin = 0
'        tW_MallRegAlw = "N"
'        tW_MallProAlw = "N"
'        dW_MallStartDate = Date
'        dW_MallEndDate = Date
'        tW_StaMallFrmAmt = "N"
'        tW_StaDiscMallAlw = "N"
'        tW_IsChkDigitMallCard = "0"
'        tVB_MemEvnCode = "" '*KT 54-11-18 V.4.0.5  DCR_5411003 PRJ09010TheMall-R4-SE-External Spec.Upload Bonus Points-00.04.00-FSBT.doc
'        tVB_IsMemJIS = "N"
    End If
    Call SP_STAxMouse(Me, False)
    Exit Function
ErrHandle:
    W_PRCxFindMallCardByID = False
    Call W_PRCxNOTFoundMallCard    '*Ao 2013/02/04
'    tVB_StaMemPointAlw = "N" '*KT 52-02-17 member point
'    tVB_MallCardGroup = ""
'    cVB_DiscMallGrp = 0
'    cVB_DiscMallRanking = 0
'    tVB_DiscMallCardKey = ""
'    tVB_DiscMallCardType = ""
'    cVB_DiscMallRate = 0
'    cW_AmtMax = 0
'    cW_AmtMin = 0
'    tW_MallRegAlw = "N"
'    tW_MallProAlw = "N"
'    dW_MallStartDate = Date
'    dW_MallEndDate = Date
'    tW_StaMallFrmAmt = "N"
'    tW_StaDiscMallAlw = "N"
'    tW_IsChkDigitMallCard = "0"
'    tVB_MemEvnCode = "" '*KT 54-11-18 V.4.0.5  DCR_5411003 PRJ09010TheMall-R4-SE-External Spec.Upload Bonus Points-00.04.00-FSBT.doc
'    tVB_IsMemJIS = "N"
    Call SP_STAxMouse(Me, False)
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Function


Private Sub Command1_Click()
'    Dim tTrack1Data  As String
'    Dim tTrack2Data  As String
'    Dim tExpDate, tTrack, tMsg  As String
'    Dim nPos As Long
'
'    tTrack1Data = "s60000120353349200940274190000000000000140200000000000000000000000060"
'    tTrack2Data = "5334920094027419=14021011387700000947"
'
'    If (tTrack1Data <> "") Or (tTrack2Data <> "") Then
'            If tTrack1Data <> "" Then
'                tTrack = Mid(tTrack1Data, 11, 16)
''                If Right(tTrack, 1) = "0" Then tTrack = Left(tTrack, 15)  '*Eaw 56-07-10 CommSheet TK-ISS3000-272-TR.xls
'                If Not W_PRCbFindTCCardByNum(tTrack) Then tTrack = Left(tTrack, 15) '*Eaw 56-09-04
'                tExpDate = Mid(tTrack1Data, 40, 4)
'            Else
'                If tTrack2Data <> "" Then
'                    If Len(tTrack2Data) > 20 Then
'                        nPos = InStrRev(tTrack2Data, "=")
'                        If nPos > 1 Then
'                            tTrack = Trim(Mid(tTrack2Data, 1, InStr(1, tTrack2Data, "=") - 1))
'                            tExpDate = Trim(Mid(tTrack2Data, InStr(1, tTrack2Data, "=") + 1, Len(tTrack2Data) - InStr(1, tTrack2Data, "=")))
'                        End If
'                    Else
'                        tTrack = tTrack2Data
'                    End If
'                Else
'                    tTrack = ""
'                End If
'            End If
'            If tTrack <> "" Then
'                otbCusID.Text = tTrack
'                If tExpDate <> "" Then
'                    orbYear.Text = Mid(tExpDate, 1, 2)
'                    orbMonth.Text = Mid(tExpDate, 3, 2)
'                End If
'            End If
'        Else
'            tMsg = "Not found Track2Data. swiped card try again."
'            Call SP_MSGnShowing("tms_CN005068" & ";" & tMsg, nCS_Error)
'            otbCusID.Text = ""
'            orbYear.Text = ""
'            orbMonth.Text = ""
'            olaName.Caption = ""
'            olaRate.Caption = ""
'            Call SP_OTBxSelText(otbCusID)
'            If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True          '*Ao 10-02-09
'            Exit Sub
'        End If
'
'        tTrack = ""
'        If nW_TrackNo = 4 Then                                  'nW_TrackNo = 4  :TC Card
'            If (tTrack1Data <> "") Then
''                tTrack = Mid(tTrack1Data, 1, 26)
'''                If Right(tTrack, 1) = "0" Then tTrack = Left(tTrack, 25)  '*Eaw 56-07-10 CommSheet TK-ISS3000-272-TR.xls
''                If Not W_PRCbFindTCCardByNum(tTrack) Then tTrack = Left(tTrack, 25) '*Eaw 56-09-27
'
''                '* Bump 57-03-07 comsheet. TK-ISS3000-465  V.4.0.120 ใช้แทน 3 line ข้างบน
'                tTrack = Mid(tTrack1Data, 11, 16)
'                If Not W_PRCbFindTCCardByNum(tTrack) Then tTrack = Mid(tTrack1Data, 11, 15)
'                tExpDate = Mid(tTrack1Data, 40, 4)
'            Else
'                If (tTrack2Data <> "") Then
'                    nPos = InStrRev(tTrack2Data, "=")
'                    If nPos > 1 Then
'                        tTrack = Trim(Mid(tTrack2Data, 1, InStr(1, tTrack2Data, "=") - 1))
'                    End If
'                End If
'            End If
'        Else
'            If (tTrack2Data <> "") Then
'                nPos = InStrRev(tTrack2Data, "=")
'                If nPos > 1 Then
'                    tTrack = Trim(Mid(tTrack2Data, 1, InStr(1, tTrack2Data, "=") - 1))
'                End If
'            Else
'                If (tTrack1Data <> "") Then
'                    tTrack = Mid(tTrack1Data, 11, 16)
''                    If Right(tTrack, 1) = "0" Then tTrack = Left(tTrack, 15)  '*Eaw 56-07-10 CommSheet TK-ISS3000-272-TR.xls
'                    If Not W_PRCbFindTCCardByNum(tTrack) Then tTrack = Left(tTrack, 15) '*Eaw 56-09-27
'                End If
'            End If
'        End If
'        tVB_DiscMallCardID = tTrack
'        '---------------------------------------------------------------------------------------------
''        If tTranc2 <> "" Then
''             If Len(Trim(Mid(tTranc2, 1, InStr(1, tTranc2, "=") - 1))) <= 20 Then    '*KT 51-06-18 Communication Sheet ML-SF-349
''                otbCusID.Text = Trim(Mid(tTranc2, 1, InStr(1, tTranc2, "=") - 1))
''                tExpDate = Trim(Mid(tTranc2, InStr(1, tTranc2, "=") + 1, Len(tTranc2) - InStr(1, tTranc2, "=")))
''                orbYear.Text = Mid(tExpDate, 1, 2)
''                orbMonth.Text = Mid(tExpDate, 3, 2)
''                Call SP_CRTxLogMallCardEvent("MSR", "DataEvent[1]", IIf(tVB_CNTrnDiscType <> "1850", "Discount", "Mall Card Register"), Trim(otbCusID.Text))
''            Else
''                Call SP_MSGnShowing("tms_CN005068", nCS_Error)
''                bW_FoundMallCard = False
''                otbCusID.Text = ""
''                orbMonth.Text = ""
''                orbYear.Text = ""
''                olaName.Caption = ""
''                olaRate.Caption = ""
''                Call SP_OTBxSelText(otbCusID)
''                Call SP_CRTxLogMallCardEvent("MSR", "DataEvent[2]", IIf(tVB_CNTrnDiscType <> "1850", "Discount", "Mall Card Register"), Trim(otbCusID.Text))
''                If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True          '*Ao 10-02-09
''                Exit Sub
''            End If
''        Else
''            tMsg = "Not found Track2Data. swiped card try again."                               '*Ao 2013/05/09
''            Call SP_MSGnShowing("tms_CN005068" & ";" & tMsg, nCS_Error)
'''            Call SP_MSGnShowing("tms_CN005068", nCS_Error)
''            bW_FoundMallCard = False
''            otbCusID.Text = ""
''            orbMonth.Text = ""
''            orbYear.Text = ""
''            olaName.Caption = ""
''            olaRate.Caption = ""
''            Call SP_OTBxSelText(otbCusID)
''            Call SP_CRTxLogMallCardEvent("MSR", "DataEvent[3]", IIf(tVB_CNTrnDiscType <> "1850", "Discount", "Mall Card Register"), Trim(otbCusID.Text))
''            If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True          '*Ao 10-02-09
''            Exit Sub
''        End If
'
'        '*Bump 56-04-02 เพิ่มการเช็ค nW_TrackNo = 4 Comsheet TK-ISS3000-023
''        If nW_TrackNo = 4 Then
''            tVB_DiscMallCardID = tCheckCardRang
''        Else
''            tVB_DiscMallCardID = Trim(otbCusID.Text)
''        End If
'
'        If Not W_PRCxFindMallCardByID(tVB_DiscMallCardID) Then
''            Call SP_MSGnShowing("tms_CN005068", nCS_Error)                     '*Ao 2013/02/04
'            bW_FoundMallCard = False
'            otbCusID.Text = ""                      '*AO51-03-20    Comm.Sheet  SF-108
'            orbMonth.Text = ""
'            orbYear.Text = ""
'            olaName.Caption = ""
'            olaRate.Caption = ""
'            Call SP_OTBxSelText(otbCusID)
'            Call SP_CRTxLogMallCardEvent("MSR", "DataEvent[4]", IIf(tVB_CNTrnDiscType <> "1850", "Discount", "Mall Card Register"), Trim(otbCusID.Text))
'            If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True          '*Ao 10-02-09
'            Exit Sub
'        Else
'
'        End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If bVB_CNSplash Then KeyCode = 0 '*KT 51-09-17 ป้องการ Key กรณีที่ Splash ขึ้น
    '*KT 52-06-05 ในหน้าจอ Tender ถ้ากดปุ่ม Price Check จะปรากฏ Screen Find Customer แล้วถ้ารูดบัตร โปรแกรมจะ Hang PH 2.0 ML-SF-012
     '*KT 52-06-05 PH 2.0 ML-SF-013 ในหน้าจอ Discount --> Mall Card --> Price check --> Enter เครื่อง POS จะHang โดยต้อง Restart อย่างเดียว
    
    
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
                        otcForm1(0).Enabled = True     'unlock text box
                        otcForm1(2).Enabled = True     'unlock ปุ่ม
                        
                        Set oEdc = Nothing
                        bW_ManaulMode = True
                        ocmCancel.Cancel = True
'                        otbAcc.Locked = False
'                        otbAcc.BackColor = nVB_CNClrBUnLock
'                        Call SP_OBJxSetFocus(otbAcc)
                        
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
    
    If Shift = 0 Then
        Select Case KeyCode
'            Case 13
'                SendKeys "{Tab}"               '*Ao 53-01-14
            Case vbKeyF4
                Call ocmFind_Click
        End Select
    End If
    End If
End Sub

Private Sub Form_Load()
On Error GoTo ErrHandle
    bVB_DiscAccept = False
    bW_FoundMallCard = False
    bW_KeyEnter = False '***KT 51-12-01 V 1.0.1
    tVB_MallCardGroup = ""
    tVB_StaMemPointAlw = "N" '*KT 52-02-17 member point
    olaRate.Caption = ""
    If (tVB_CNTrnDiscType = "1850") Then tVB_MallCardIDRef = "" '*Tao 60-04-27
    Call SP_FRMxChildCenter(Me)
    Call SP_FRMxGetCaptionFrmDB2Tag(Me)
    Call SP_FRMxSetCapForm(Me)
    bW_ChkESC = True
    Call SP_EDCbCheckEDC
    tW_Comport = tVB_EDCComport
    If bVB_Res640X480 Then
        Me.Top = 2100
    End If
    
    '1809 : Member card   1887 :TC Card 1850:Bonus Point Award
    bW_EnableMSR = True
    If tVB_CNTrnDiscType = "1850" Then '*Bump 58-04-18 'CommSheet TK-ISS3000-558.02_22042015.xls
        bW_EnableMSR = False
        ocmEDC.Enabled = True
    Else
        ocmEDC.Enabled = False
    End If

    If bVB_OPOSMSREnable And bW_EnableMSR Then '*Bump 58-04-18 เพิ่ม bW_EnableMSR
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
                Call SP_MSGnShowing("tms_CN005092", nCS_Warning)
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
              '*Bump 56-02-28 เพิ่ม code 3 บรรทัด
            If .DeviceName = tSNR_TOUCH_SCANNER Then
                .DecodeData = True
            End If
        End With
    End If
    '*KT 55-08-07 PRJ10001TK-External Specification Bonus Point Award-00.01.00.doc
    '*KT 55-10-10 ยกเลิก reqirement เนื่องจากเกรงว่าอุปกรณ์เสียจะทำงานไม่ได้
'     If tVB_CNTrnDiscType = "1850" Then
'        otbCusID.Locked = True
'    Else
'        otbCusID.Locked = False
'    End If
    If bVB_Discount91 Then     '*Eaw 56-01-05
        Call W_PRCxSpcMem91
    End If
    
    Exit Sub
ErrHandle:
    
End Sub

Private Sub W_EDCxHyperComCreditCard()
    With oEdc
''''''        .SetNetAmout = otbAmt.Text
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
'''''                    otbAcc.Text = .GetCardNumber
                    'otbBank.Text = .GetCardBankCode
                    'otbTrace.Text = .GetTraceCode
'''''                    otbApvCode.Text = .GetApproveCode
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
'''''        .SetNetAmout = otbAmt.Text
        .bShowThai = True
        Call SP_INIbWriteFile(App.Path & "\History\EDCLog" & Format(Date, "yy-mm-dd") & ".log", "PosNet", tVB_Running, "START :" & Format(Time, "HH:mm:ss"))
        If .EDC1CreateMessage Then
            Call SP_INIbWriteFile(App.Path & "\History\EDCLog" & Format(Date, "yy-mm-dd") & ".log", "PosNet", tVB_Running, "EDC1CreateMessage :" & Format(Time, "HH:mm:ss"))
           If .EDC2SendMessage(tVB_CNEdcPort) Then
                Call SP_INIbWriteFile(App.Path & "\History\EDCLog" & Format(Date, "yy-mm-dd") & ".log", "PosNet", tVB_Running, "EDC2SendMessage :" & Format(Time, "HH:mm:ss"))
                If .EDC3VerifyCreditCard Then
                    Call SP_INIbWriteFile(App.Path & "\History\EDCLog" & Format(Date, "yy-mm-dd") & ".log", "PosNet", tVB_Running, "EDC3VerifyCreditCard :" & Format(Time, "HH:mm:ss"))
''''''                    otbAcc.Text = .GetCardNumber
                    'otbBank.Text = .GetCardBankCode
                    'otbTrace.Text = .GetTraceCode
''''''                    otbApvCode.Text = .GetApproveCode
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
''''        .SetNetAmout = otbAmt.Text
        .bShowThai = True
        Call SP_INIbWriteFile(App.Path & "\History\EDCLog" & Format(Date, "yy-mm-dd") & ".log", "Verifone", tVB_Running, "START :" & Format(Time, "HH:mm:ss"))
        If .EDC1CreateMessage And bW_ManaulMode = False Then
            Call SP_INIbWriteFile(App.Path & "\History\EDCLog" & Format(Date, "yy-mm-dd") & ".log", "Verifone", tVB_Running, "EDC1CreateMessage :" & Format(Time, "HH:mm:ss"))
           If .EDC2SendMessage(tVB_CNEdcPort) And bW_ManaulMode = False Then
                Call SP_INIbWriteFile(App.Path & "\History\EDCLog" & Format(Date, "yy-mm-dd") & ".log", "Verifone", tVB_Running, "EDC2SendMessage :" & Format(Time, "HH:mm:ss"))
                If .EDC3VerifyCreditCard And bW_ManaulMode = False Then
                    Call SP_INIbWriteFile(App.Path & "\History\EDCLog" & Format(Date, "yy-mm-dd") & ".log", "Verifone", tVB_Running, "EDC3VerifyCreditCard :" & Format(Time, "HH:mm:ss"))
''''''                    otbAcc.Text = .GetCardNumber
                    'otbBank.Text = .GetCardBankCode
                    'otbTrace.Text = .GetTraceCode
''''''                    otbApvCode.Text = .GetApproveCode
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

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrHandle
    With OPOSMSR1
        If bVB_OPOSMSREnable And bW_EnableMSR Then
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
'    Set wCNDiscMall = Nothing
    Call SP_DATxRsNothing(orsMall)
    Call SP_DATxRsNothing(orsMem)
    Call SP_OBJxSetNothing(Me)
ErrHandle:
    
End Sub

Private Sub ocmCancel_Click()
On Error GoTo ErrHandle
    bW_ClickEDC = False
    If bVB_OPOSMSREnable And bW_EnableMSR Then OPOSMSR1.DataEventEnabled = False          '*Ao 10-02-09
    If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = False
    Call SP_CRTxLogMallCardEvent("Keyboard", "Cancel/Esc", IIf(tVB_CNTrnDiscType <> "1850", "Discount", "Mall Card Register"), Trim(otbCusID.Text))
    bVB_DiscAccept = False
    bVB_Discount91 = False
    If (tVB_CNTrnDiscType = "1850") Then tVB_MallCardIDRef = "" '*Tao 60-04-27
    Unload Me
    Exit Sub
ErrHandle:
    If bVB_OPOSMSREnable And bW_EnableMSR Then OPOSMSR1.DataEventEnabled = True           '*Ao 10-02-09
    If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True
End Sub

Private Sub ocmEDC_Click()
    If (tVB_CNTrnDiscType = "1850") Then tVB_MallCardIDRef = "" '*Tao 60-04-20 EDC ถือว่าไม่ใช่ Manual (ถ้าถือว่าเป้น Manual ให้ใส่ *)
    bW_ClickEDC = True
    Call W_PRCxloadEDC
End Sub

Private Sub ocmFind_Click()
On Error GoTo ErrHandle

    tVB_DiscMallCardIDByTel = ""
    wCNFindMem.Show vbModal
    If tVB_DiscMallCardIDByTel <> "" Then
        otbCusID.Text = tVB_DiscMallCardIDByTel
        If (tVB_CNTrnDiscType = "1850") Then tVB_MallCardIDRef = "*" '*Tao 60-04-20 Manual
        Call otbCusID_KeyPress(13)
        If olaName.Caption <> "" Then Call SP_OBJxSetFocus(orbMonth)
    Else
        Call SP_OBJxSetFocus(otbCusID)
    End If
    Exit Sub
ErrHandle:
    
End Sub

Private Sub ocmOK_Click()
    Dim nI As Long
    Dim tStaMemberPointSpecial As String
    Dim tSql As String
    
On Error GoTo ErrHandle
    
    If bVB_OPOSMSREnable And bW_EnableMSR Then OPOSMSR1.DataEventEnabled = False         '*Ao 10-02-09
    If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = False
    
    Call SP_CRTxLogMallCardEvent("Keyboard/Keyboard", "OK[1]", IIf(tVB_CNTrnDiscType <> "1850", "Discount", "Mall Card Register"), Trim(otbCusID.Text) & " Month:" & Trim(orbMonth.Text) & " Year:" & Trim(orbYear.Text))
    If bW_EnableMSR = True Then
        If Not bW_FoundMallCard Then
            bW_KeyEnter = False '***KT 51-12-01 V 1.0.1
            Call SP_CRTxLogMallCardEvent("MSR/Keyboard", "OK[2]", IIf(tVB_CNTrnDiscType <> "1850", "Discount", "Mall Card Register"), Trim(otbCusID.Text))
            If bVB_OPOSMSREnable And bW_EnableMSR Then OPOSMSR1.DataEventEnabled = True          '*Ao 10-02-09
            If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True
            If bW_ClickEDC Then '*Bump 58-05-02
                SleepEx 1000, 0
                Unload Me
            End If
            Exit Sub
        End If
    End If
    If UCase(tW_ChkExp) = "Y" Then 'Comsheet ML-SF-171
        If orbMonth.Text = "" Or orbYear.Text = "" Then
            Call SP_MSGnShowing("tms_CN005141", nCS_Warning)
            If orbMonth.Text = "" Then
                Call SP_OBJxSetFocus(orbMonth)
            Else
                Call SP_OBJxSetFocus(orbYear)
            End If
            bW_KeyEnter = False '***KT 51-11-01 V 1.0.1
            Call SP_CRTxLogMallCardEvent("MSR/Keyboard", "OK[3]", IIf(tVB_CNTrnDiscType <> "1850", "Discount", "Mall Card Register"), Trim(otbCusID.Text))
            If bVB_OPOSMSREnable And bW_EnableMSR Then OPOSMSR1.DataEventEnabled = True           '*Ao 10-02-09
            If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True
            If bW_ClickEDC Then '*Bump 58-05-02
                SleepEx 1000, 0
                Unload Me
            End If
            Exit Sub
        End If
        If Format(orbMonth.Text, "00") = "00" Then              '*Ao 52-02-05
            Call SP_MSGnShowing("tms_CN005141", nCS_Warning)
            orbMonth.Text = ""
            orbYear.Text = ""
            Call SP_OBJxSetFocus(orbMonth)
            bW_KeyEnter = False '***KT 51-11-01 V 1.0.1
            Call SP_CRTxLogMallCardEvent("MSR/Keyboard", "OK[4]", IIf(tVB_CNTrnDiscType <> "1850", "Discount", "Mall Card Register"), Trim(otbCusID.Text))
            If bVB_OPOSMSREnable And bW_EnableMSR Then OPOSMSR1.DataEventEnabled = True          '*Ao 10-02-09
            If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True
            If bW_ClickEDC Then '*Bump 58-05-02
                SleepEx 1000, 0
                Unload Me
            End If
            Exit Sub
        End If
        '_______________________________________
        dVB_DiscMallCardExp = DateSerial(2000 + Val(Format(orbYear.Text)), Val(Format(orbMonth.Text)) + 1, 1)
         If Format(dVB_DiscMallCardExp, "YYYY/MM/DD") <= Format(dVB_CurrentDate, "YYYY/MM/DD") Then 'ทาง FSBT ให้ทำการตรวจสอบวันที่ Expire Date ตาม System Date
            Call SP_MSGnShowing("tms_CN005141", nCS_Warning)
            orbMonth.Text = ""              '*AO51-03-20      Comm.Sheet SF-108
            orbYear.Text = ""
            Call SP_OBJxSetFocus(orbMonth)
            bW_KeyEnter = False '***KT 51-11-01 V 1.0.1
            Call SP_CRTxLogMallCardEvent("MSR/Keyboard", "OK[5]", IIf(tVB_CNTrnDiscType <> "1850", "Discount", "Mall Card Register"), Trim(otbCusID.Text))
            If bVB_OPOSMSREnable Then OPOSMSR1.DataEventEnabled = True          '*Ao 10-02-09
            If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True
            If bW_ClickEDC Then '*Bump 58-05-02
                SleepEx 1000, 0
                Unload Me
            End If
            Exit Sub
        End If
    Else
        If orbYear.Text = "" Or orbMonth.Text = "" Then
            Call SP_MSGnShowing("tms_CN005141", nCS_Warning)
            If orbMonth.Text = "" Then
                Call SP_OBJxSetFocus(orbMonth) 'orbMonth.SetFocus
            Else
                Call SP_OBJxSetFocus(orbYear) 'orbYear.SetFocus
            End If
            bW_KeyEnter = False '***KT 51-11-01 V 1.0.1
            Call SP_CRTxLogMallCardEvent("MSR/Keyboard", "OK[6]", IIf(tVB_CNTrnDiscType <> "1850", "Discount", "Mall Card Register"), Trim(otbCusID.Text))
            If bVB_OPOSMSREnable And bW_EnableMSR Then OPOSMSR1.DataEventEnabled = True          '*Ao 10-02-09
            If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True
            Exit Sub
        End If
    End If
    '*Bump 56-12-04 Comment ไว้ก่อน แล้ว Recompile เป็น POSFront[4.0.109]20131204_1732.exe      CommSheet TK-ISS3000-437-FAPL-162-00.00.01_NG.xls
    'เนื่องจาก กด StoreWide Discount ไม่ได้  เช่น
    ' Allow Discount  = Yes
    ' Start Date 01/11/2013     End Date    31/12/2013
    'ในหน้าจอนี้ ถ้าใส่ Expiration Date = 12/13(MMYY) จะต้องได้ส่วนลด เนื่องจาก ยังอยู่ในเดือน 12/2013  แต่โปรกรมแจ้งว่า Mall Card Expire Date.   ซึ่งไม่ถูกต้อง
    '
    'ในหน้าจอนี้ ถ้าใส่ Expiration Date = 01/14(MMYY) จะต้องไม่ได้ส่วนลด เนื่องจาก Discount กำหนด EndDate ไว้ เป็น    31/12/2013
    '
    '*Ao 2013/12/03      FDMcdStartDate,FDMcdEndDate
'    If Not (Format(dW_MallStartDate, "yyyy/MM/dd") <= Format(tUT_SaleDate, "yyyy/MM/dd") And Format(dW_MallEndDate, "yyyy/MM/dd") >= Format(tUT_SaleDate, "yyyy/MM/dd")) Then
'        Call SP_MSGnShowing("tms_CN005141", nCS_Error)
'        bVB_Discount91 = False
'        otbCusID.Text = ""
'        otbCusID.Enabled = True
'        otbCusID.Locked = False
'        orbMonth.Text = ""
'        orbYear.Text = ""
'        olaName.Caption = ""
'        olaRate.Caption = ""
'        Call SP_OTBxSelText(otbCusID)
'        Exit Sub
'    Else
'        bVB_Discount91 = False
'    End If
'*********************************************************************************************************
    tVB_DiscMallCardID = Trim(otbCusID.Text)
    If tVB_CNTrnDiscType <> "1850" Then
        'ตรวจสอบกลุ่มของ Mall Card ว่าก่อนหน้านีมีการรูด Mall Card กลุ่มเดียวกันหรือไม่
        '*KT 53-03-30 PH1.8-ML-SF-016
        Call SP_CRTxLogMallCardEvent("MSR/Keyboard", "OK[6]", IIf(tVB_CNTrnDiscType <> "1850", "Discount", "Mall Card Register"), Trim(otbCusID.Text))
        If W_CHKbMallCardGroup(tVB_MallCardGroup) Then
            Call SP_MSGnShowing("tms_CN005151", nCS_Warning)
            Call SP_OBJxSetFocus(otbCusID)
            bW_KeyEnter = False '***KT 51-11-01 V 1.0.1
            Call SP_CRTxLogMallCardEvent("MSR/Keyboard", "OK[7]", IIf(tVB_CNTrnDiscType <> "1850", "Discount", "Mall Card Register"), Trim(otbCusID.Text))
            If bVB_OPOSMSREnable And bW_EnableMSR Then OPOSMSR1.DataEventEnabled = True           '*Ao 10-02-09
            If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True
            Exit Sub
        End If
        tStaMemberPointSpecial = SP_FEDtChkString(orsMall, "FTMcdAlwPoint")
        If tW_ChkDpt = "N" And tW_StaDiscMallAlw = "N" Then
            cVB_BalMallDisc = 0
            bVB_DiscAccept = False
            If tVB_StaMemPointAlw = "N" And tStaMemberPointSpecial = "N" Then
                Call SP_MSGnShowing("tms_CN005178", nCS_Warning)
            Else
                Call SP_MSGnShowing("tms_CN005204", nCS_Warning)
            End If
            Call SP_CRTxLogMallCardEvent("MSR/Keyboard", "OK[8]", IIf(tVB_CNTrnDiscType <> "1850", "Discount", "Mall Card Register"), Trim(otbCusID.Text))
            Call SP_OBJxSetFocus(otbCusID)
            bW_KeyEnter = False
            If bVB_OPOSMSREnable And bW_EnableMSR Then OPOSMSR1.DataEventEnabled = True          '*Ao 10-02-09
            If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True
            Exit Sub
        Else
            Call SP_CRTxLogMallCardEvent("MSR/Keyboard", "OK[9]", IIf(tVB_CNTrnDiscType <> "1850", "Discount", "Mall Card Register"), Trim(otbCusID.Text))
            '*KT 52-05-26 PH 1.8 ML-SF-075 ในกรณีที่ POS ไม่สามารถติดต่อ Member Server ได้ ขอให้ POS แสดง Warning Message เพียงครั้งเดียว
            If ocnMember Is Nothing Then
                Call SP_MSGnShowing("tms_CN005134", nCS_Warning) '*tms_CN005134 = Can not connect the database of Member Server.
                oDisc.bReConnectMember = False
                Call SP_CRTxLogMallCardEvent("MSR/Keyboard", "OK[10]", IIf(tVB_CNTrnDiscType <> "1850", "Discount", "Mall Card Register"), Trim(otbCusID.Text))
            Else
                oDisc.bReConnectMember = False
                Call SP_CRTxLogMallCardEvent("MSR/Keyboard", "OK[11]", IIf(tVB_CNTrnDiscType <> "1850", "Discount", "Mall Card Register"), Trim(otbCusID.Text))
            End If
            oDisc.bRecalByItem = False
            oDisc.DISxClearAll
            oDisc.dTranDate = tUT_SaleDate
            oDisc.tTerminalNo = tVB_CNTerminalNum
            oDisc.tTransNo = tVB_Running
            oDisc.tCardID = tVB_DiscMallCardID
            oDisc.tCardType = tVB_DiscMallCardType
            oDisc.tCardKey = tVB_DiscMallCardKey
            oDisc.cDiscValue = SP_VALcText2Double(olaRate.Caption)
            dVB_DiscMallCardExp = DateSerial(2000 + Val(Format(orbYear.Text)), Val(Format(orbMonth.Text)), 1)
            oDisc.dCardExpDate = dVB_DiscMallCardExp
            oDisc.tStaChkExpire = tW_ChkExp
            oDisc.tStaChkDepart = tW_ChkDpt
            oDisc.cAlwDis = cVB_DiscMallRanking
            oDisc.cAmtMallMax = cW_AmtMax
            oDisc.cAmtMallMin = cW_AmtMin
            oDisc.tMallRegAlw = tW_MallRegAlw
            oDisc.tMallProAlw = tW_MallProAlw
            oDisc.dMallStartDate = dW_MallStartDate
            oDisc.dMallEndDate = dW_MallEndDate
            oDisc.tMallFrmAmt = tW_StaMallFrmAmt
            oDisc.tDiscMallAlw = tW_StaDiscMallAlw
            oDisc.tCardName = tW_CardName '*KT 52-11-26 PH2-ML-SF-170 ใน Ticket ไม่พิมพ์ชื่อพนักงานที่ใช้ส่วนลด Employee discount
            oDisc.tMallAlwPoint = tVB_StaMemPointAlw '*KT 54-06-20 V.4.0.0 Ph.4.0 Req.13 การให้ Special Point(SKU Point)และสิทธิส่วนลด
            oDisc.tEventID = ""                                 '*Ao 2013/07/31     CommSheet TK-ISS3000-148-IT2 00.00.02-Aoo.xls
            Call oDisc.DIScDiscount(nEN_Disc11MallCrd, eEN_VerifyDiscType.nEN_Type2ByDcn)
            cVB_BalMallDisc = oDisc.cAmtAllMall 'Sub Reg ที่เหลือ
            bVB_DiscAccept = True
            Call SP_CRTxLogMallCardEvent("MSR/Keyboard", "OK[12]", IIf(tVB_CNTrnDiscType <> "1850", "Discount", "Mall Card Register"), Trim(otbCusID.Text))
        End If
'*KT 52-02-17
'1 member point ไม่ต้องคำนวน Point ที่ POS แล้ว
'2 กรณีที่ member server ไม่ได้ให้ไม่เก็บ member ที่แสดง last point
'3 กรณีที่ mall card maintenance ไม่ flag สถานะเป็น member point ไม่เก็บ member ที่แสดง last point
'        If tVB_StaMemPointAlw = "Y" Then                                                               '*Ao 2013/03/12  TK POINT
'            nVB_DiscMallCount = nVB_DiscMallCount + 1
'            Call W_PRCxPointMallCardEx 'เก็บค่า member เพื่อแสดง last point
'            Call SP_CRTxLogMallCardEvent("MSR/Keyboard", "OK[13]", IIf(tVB_CNTrnDiscType <> "1850", "Discount", "Mall Card Register"), Trim(otbCusID.Text))
'        End If
    Else
        tStaMemberPointSpecial = SP_FEDtChkString(orsMall, "FTMcdAlwPoint")
        If (tVB_StaMemPointAlw = "Y") And (tStaMemberPointSpecial = "Y") Then
'            If Trim(tVB_MemPointID) = "" Then '*Eaw 56-07-06  CommSheet TK-ISS3000-264-TR.xlsx
                tVB_MemPointID = Trim(otbCusID.Text)        '*Ao 2013/02/27    TK POINT
                'wDisc.olaMEM.Caption = "BP Award ID:" & tVB_MemPointID '*Tao 60-04-21 ของเดิม
                wDisc.olaMEM.Caption = "BP Award ID:" & tVB_MemPointID & tVB_MallCardIDRef  '*Tao 60-04-21 ถ้าได้ MallCardID จาก Manual ให้ใส่ *
                wDisc.olaMEM.Refresh '*Tao 60-05-02 แสดงเลขสมาชิก
                tVB_MemExp = Trim(Format(orbMonth.Text, "00")) & Trim(Format(orbYear.Text, "00"))
                tVB_ChkMemExp = tW_ChkExp
                tSql = "UPDATE " & tVB_CNTblHD
                tSql = tSql & vbCrLf & "SET FNCstCardPoint = 0"
                tSql = tSql & vbCrLf & ",FTCstCardCode = '" & tVB_MemPointID & "'"
                tSql = tSql & vbCrLf & ",FTCstCardType = '" & tVB_MemExp & "'"
                tSql = tSql & vbCrLf & ",FTShdCshOrCrd = '" & tVB_MallCardIDRef & "'" '*Tao 60-04-21 ถ้าได้ MallCardID จาก Manual ให้ใส่ *
                tSql = tSql & vbCrLf & "WHERE FTTmnNum = '" & tVB_CNTerminalNum & "'"
                tSql = tSql & vbCrLf & "AND FTShdTransNo= '" & tVB_Running & "'"
                tSql = tSql & vbCrLf & "AND FDShdTransDate= '" & tUT_SaleDate & "'"
                Call SP_SQLvExecute(tSql)
'            End If
        Else
            If (tVB_CNTrnDiscType = "1850") Then tVB_MallCardIDRef = "" '*Tao 60-04-27 ถ้าไม่ให้ Point ไม่ต้องใส่ *
        End If
        Call SP_CRTxLogMallCardEvent("MSR/Keyboard", "OK[14]", IIf(tVB_CNTrnDiscType <> "1850", "Discount", "Mall Card Register"), Trim(otbCusID.Text))
'        If ocnMember Is Nothing Then
'            Call SP_MSGnShowing("tms_CN005134", nCS_Warning)
'            Call SP_CRTxLogMallCardEvent("MSR/Keyboard", "OK[15]", IIf(tVB_CNTrnDiscType <> "1850", "Discount", "Mall Card Register"), Trim(otbCusID.Text))
'        End If
'        If nVB_RegMallCount > 0 Then
'            For nI = LBound(aVB_MallCardReg) To UBound(aVB_MallCardReg)
'                If aVB_MallCardReg(nI).ID = tVB_DiscMallCardID Then
'                    Call SP_CRTxLogMallCardEvent("MSR/Keyboard", "OK[16]", IIf(tVB_CNTrnDiscType <> "1850", "Discount", "Mall Card Register"), Trim(otbCusID.Text))
'                    Unload Me
'                    Exit Sub
'                End If
'            Next nI
'        End If
'        nVB_RegMallCount = nVB_RegMallCount + 1
'        Call W_PRCxRegMallCard
'        Call SP_CRTxLogMallCardEvent("MSR/Keyboard", "OK[17]", IIf(tVB_CNTrnDiscType <> "1850", "Discount", "Mall Card Register"), Trim(otbCusID.Text))
    End If
    bVB_Discount91 = False          '*Ao 2013/12/06
'    If bW_EnableMSR = False Then
'        otmTime.Enabled = True
'    End If
    Unload Me
    Exit Sub
ErrHandle:
    bW_KeyEnter = False
    Call SP_MSGxShowError(Err.Number, Err.Description)
    Call SP_CRTxLogMallCardEvent("MSR/Keyboard", "OK[18]", IIf(tVB_CNTrnDiscType <> "1850", "Discount", "Mall Card Register"), Trim(otbCusID.Text))
    If bVB_OPOSMSREnable And bW_EnableMSR Then OPOSMSR1.DataEventEnabled = True           '*Ao 10-02-09
    If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True
End Sub

Private Sub OPOSMSR1_DataEvent(ByVal Status As Long)
    Dim tExpDate$
    Dim tTranc2$
    
    Dim tTrack1Data$ 'Bump 56-04-02 TC Card
    Dim tTrack2Data$ 'Bump 54-04-02 TC Card
'    Dim tCheckCardRang  As String   'Bump 54-04-02 TC Card
'    Dim tCardID  As String   'Bump 54-04-02 TC Card
    Dim nPos As Long
    Dim tMsg As String
    Dim tTrack As String
    
On Error GoTo ErrHandle '*KT 52-02-11  PH2-ML-SF-092 โปรแกรมหลุดมาที่ Desktop ตอนพิมพ์ receipt , ชำระด้วย Credit
    With oScanner                                                           '*Ao 10-02-09
     If bVB_OPOSScannerEnable Then
            'Enable the event.
            .DataEventEnabled = False
     End If
     End With
    With OPOSMSR1
'        Select Case nW_TrackNo                     '*Ao 2013/05/15         แก้ไขการใช้ MSR
'            Case 1
'                tTrack1Data = .Track1Data
'            Case 2
'                tTrack2Data = .Track2Data
'            Case 3
'                tTranc2 = .Track3Data
'            Case 4
'                    'tTranc2 = .Track4Data
'                    tTranc2 = .Track2Data   '*Bump เปลี่ยนเป็นอ่านหมายเลขบัตรตามปกติ
'                    '*Bump 56-04-02 เพิ่ม Code การเช็ค TC Card Comsheet TK-ISS3000-023
'                    tTrack1Data = .Track1Data '-->"s60000120345389706440548160000000000000040200000000000000000000000060"
'                    tTrack2Data = .Track2Data   '-->"4538970644054816=04021011387700000947"
''                    Call SP_CRTxLogMallCardEvent("MSR/TCCard", "MSR", "Track1Data:", tTrack1Data & ":")
''                    Call SP_CRTxLogMallCardEvent("MSR/TCCard", "MSR", "Track2Data:", tTrack2Data & ":")
''                    If tTrack1Data <> "" And tTrack2Data <> "" Then
''                        nPos = InStrRev(tTrack2Data, "=")
''                        If nPos > 1 Then
''                            tCardID = Trim(Mid(tTrack2Data, 1, nPos - 1))  '-->"4538970644054816" คือ หมายเลขบนบัตร TC Card
''                            nPos = InStrRev(tTrack1Data, tCardID)
''                                If nPos > 1 Then
''                                    tCheckCardRang = Mid(tTrack1Data, (nPos - 10), 26)
''                                    If Len(tCheckCardRang) < 26 Then tCheckCardRang = tCardID
''                                Else
''                                    tCheckCardRang = tCardID
''                                End If
''                                Call SP_CRTxLogMallCardEvent("MSR/TCCard", "MSR", "CardID:", tCardID & ":")
''                                Call SP_CRTxLogMallCardEvent("MSR/TCCard", "MSR", "IDMark->Validaty Code:", tCheckCardRang)
''                        End If
''                    End If
'        End Select
        '*Ao 2013/05/15         แก้ไขการใช้ MSR
'        tTrack1Data = "s60000120345389706440548160000000000000140200000000000000000000000060"
'        tTrack2Data = "4538970644054816=14021011387700000947"
        tTrack1Data = .Track1Data
        tTrack2Data = .Track2Data
        
        .ClearInput
        .DataEventEnabled = True
        Call SP_CRTxLogMallCardEvent("MSR/TCCard", "MSR", "Track1Data:", tTrack1Data & ":")
        Call SP_CRTxLogMallCardEvent("MSR/TCCard", "MSR", "Track2Data:", tTrack2Data & ":")
        '1809 : Member card   1887 :TC Card 1850:Bonus Point Award
        
        If (tTrack1Data <> "") Or (tTrack2Data <> "") Then
            If tTrack1Data <> "" Then
                tTrack = Mid(tTrack1Data, 11, 16)
'                If Right(tTrack, 1) = "0" Then tTrack = Left(tTrack, 15)  '*Eaw 56-07-10 CommSheet TK-ISS3000-272-TR.xls
                If Not W_PRCbFindTCCardByNum(tTrack) Then tTrack = Left(tTrack, 15) '*Eaw 56-09-04
                tExpDate = Mid(tTrack1Data, 40, 4)
            Else
                If tTrack2Data <> "" Then
                    If Len(tTrack2Data) > 20 Then
                        nPos = InStrRev(tTrack2Data, "=")
                        If nPos > 1 Then
                            tTrack = Trim(Mid(tTrack2Data, 1, InStr(1, tTrack2Data, "=") - 1))
                            tExpDate = Trim(Mid(tTrack2Data, InStr(1, tTrack2Data, "=") + 1, Len(tTrack2Data) - InStr(1, tTrack2Data, "=")))
                        End If
                    Else
                        tTrack = tTrack2Data
                    End If
                Else
                    tTrack = ""
                End If
            End If
            If tTrack <> "" Then
                otbCusID.Text = tTrack
                If tExpDate <> "" Then
                    orbYear.Text = Mid(tExpDate, 1, 2)
                    orbMonth.Text = Mid(tExpDate, 3, 2)
                End If
            End If
        Else
            tMsg = "Not found Track2Data. swiped card try again."
            Call SP_MSGnShowing("tms_CN005068" & ";" & tMsg, nCS_Warning)
            otbCusID.Text = ""
            orbYear.Text = ""
            orbMonth.Text = ""
            olaName.Caption = ""
            olaRate.Caption = ""
            Call SP_OTBxSelText(otbCusID)
            If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True          '*Ao 10-02-09
            Exit Sub
        End If
        
        tTrack = ""
        If nW_TrackNo = 4 Then                                  'nW_TrackNo = 4  :TC Card
            If (tTrack1Data <> "") Then
'                tTrack = Mid(tTrack1Data, 1, 26)
''                If Right(tTrack, 1) = "0" Then tTrack = Left(tTrack, 25)  '*Eaw 56-07-10 CommSheet TK-ISS3000-272-TR.xls
'                If Not W_PRCbFindTCCardByNum(tTrack) Then tTrack = Left(tTrack, 25) '*Eaw 56-09-27

'                '* Bump 57-03-07 comsheet. TK-ISS3000-465  V.4.0.121 ใช้แทน 3 line ข้างบน
'                '* ใช้คู่กับ STP_PRCxFindMallCardByID_21.0.SQL
                tTrack = Mid(tTrack1Data, 11, 16)
                If Not W_PRCbFindTCCardByNum(tTrack) Then tTrack = Mid(tTrack1Data, 11, 15)
                tExpDate = Mid(tTrack1Data, 40, 4)
            Else
                If (tTrack2Data <> "") Then
                    nPos = InStrRev(tTrack2Data, "=")
                    If nPos > 1 Then
                        tTrack = Trim(Mid(tTrack2Data, 1, InStr(1, tTrack2Data, "=") - 1))
                    End If
                End If
            End If
        Else
            If (tTrack2Data <> "") Then
                nPos = InStrRev(tTrack2Data, "=")
                If nPos > 1 Then
                    tTrack = Trim(Mid(tTrack2Data, 1, InStr(1, tTrack2Data, "=") - 1))
                End If
            Else
                If (tTrack1Data <> "") Then
                    tTrack = Mid(tTrack1Data, 11, 16)
'                    If Right(tTrack, 1) = "0" Then tTrack = Left(tTrack, 15)  '*Eaw 56-07-10 CommSheet TK-ISS3000-272-TR.xls
                    If Not W_PRCbFindTCCardByNum(tTrack) Then tTrack = Left(tTrack, 15) '*Eaw 56-09-27
                End If
            End If
        End If
        tVB_DiscMallCardID = tTrack
        '---------------------------------------------------------------------------------------------
'        If tTranc2 <> "" Then
'             If Len(Trim(Mid(tTranc2, 1, InStr(1, tTranc2, "=") - 1))) <= 20 Then    '*KT 51-06-18 Communication Sheet ML-SF-349
'                otbCusID.Text = Trim(Mid(tTranc2, 1, InStr(1, tTranc2, "=") - 1))
'                tExpDate = Trim(Mid(tTranc2, InStr(1, tTranc2, "=") + 1, Len(tTranc2) - InStr(1, tTranc2, "=")))
'                orbYear.Text = Mid(tExpDate, 1, 2)
'                orbMonth.Text = Mid(tExpDate, 3, 2)
'                Call SP_CRTxLogMallCardEvent("MSR", "DataEvent[1]", IIf(tVB_CNTrnDiscType <> "1850", "Discount", "Mall Card Register"), Trim(otbCusID.Text))
'            Else
'                Call SP_MSGnShowing("tms_CN005068", nCS_Error)
'                bW_FoundMallCard = False
'                otbCusID.Text = ""
'                orbMonth.Text = ""
'                orbYear.Text = ""
'                olaName.Caption = ""
'                olaRate.Caption = ""
'                Call SP_OTBxSelText(otbCusID)
'                Call SP_CRTxLogMallCardEvent("MSR", "DataEvent[2]", IIf(tVB_CNTrnDiscType <> "1850", "Discount", "Mall Card Register"), Trim(otbCusID.Text))
'                If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True          '*Ao 10-02-09
'                Exit Sub
'            End If
'        Else
'            tMsg = "Not found Track2Data. swiped card try again."                               '*Ao 2013/05/09
'            Call SP_MSGnShowing("tms_CN005068" & ";" & tMsg, nCS_Error)
''            Call SP_MSGnShowing("tms_CN005068", nCS_Error)
'            bW_FoundMallCard = False
'            otbCusID.Text = ""
'            orbMonth.Text = ""
'            orbYear.Text = ""
'            olaName.Caption = ""
'            olaRate.Caption = ""
'            Call SP_OTBxSelText(otbCusID)
'            Call SP_CRTxLogMallCardEvent("MSR", "DataEvent[3]", IIf(tVB_CNTrnDiscType <> "1850", "Discount", "Mall Card Register"), Trim(otbCusID.Text))
'            If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True          '*Ao 10-02-09
'            Exit Sub
'        End If
        
        '*Bump 56-04-02 เพิ่มการเช็ค nW_TrackNo = 4 Comsheet TK-ISS3000-023
'        If nW_TrackNo = 4 Then
'            tVB_DiscMallCardID = tCheckCardRang
'        Else
'            tVB_DiscMallCardID = Trim(otbCusID.Text)
'        End If
        
        If Not W_PRCxFindMallCardByID(tVB_DiscMallCardID) Then
'            Call SP_MSGnShowing("tms_CN005068", nCS_Error)                     '*Ao 2013/02/04
            bW_FoundMallCard = False
            otbCusID.Text = ""                      '*AO51-03-20    Comm.Sheet  SF-108
            orbMonth.Text = ""
            orbYear.Text = ""
            olaName.Caption = ""
            olaRate.Caption = ""
            Call SP_OTBxSelText(otbCusID)
            Call SP_CRTxLogMallCardEvent("MSR", "DataEvent[4]", IIf(tVB_CNTrnDiscType <> "1850", "Discount", "Mall Card Register"), Trim(otbCusID.Text))
            If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True          '*Ao 10-02-09
            Exit Sub
        Else
            '*KT 52-09-30 PH2-ML-SF-149 หมายเลขบัตร Mall Card ไม่ตรงกัน ระหว่าง M-Price กับ รูด Mall Card (Mall8)
            If tW_IsChkDigitMallCard = "1" Then
                If Not SP_PRCbIsChkDigitMallCard(Trim(otbCusID.Text)) Then
                    Call SP_MSGnShowing("tms_CN005216", nCS_Warning)
                    bW_FoundMallCard = False
                    otbCusID.Text = ""                      '*AO51-03-20    Comm.Sheet  SF-108
                    orbMonth.Text = ""
                    orbYear.Text = ""
                    olaName.Caption = ""
                    olaRate.Caption = ""
                    Call SP_OTBxSelText(otbCusID)
                    Call SP_CRTxLogMallCardEvent("MSR", "DataEvent[5]", IIf(tVB_CNTrnDiscType <> "1850", "Discount", "Mall Card Register"), Trim(otbCusID.Text))
                    If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True          '*Ao 10-02-09
                    Exit Sub
                End If
            End If
            bW_FoundMallCard = True
            Call SP_CRTxLogMallCardEvent("MSR", "DataEvent[6]", IIf(tVB_CNTrnDiscType <> "1850", "Discount", "Mall Card Register"), Trim(otbCusID.Text))
        End If
        
        If UCase(tW_ChkExp) = "Y" Then
            dVB_DiscMallCardExp = DateSerial(2000 + Val(Format(orbYear.Text)), Val(Format(orbMonth.Text)) + 1, 1)
            If Format(dVB_DiscMallCardExp, "YYYY/MM/DD") <= Format(dVB_CurrentDate, "YYYY/MM/DD") Then 'ทาง FSBT ให้ทำการตรวจสอบวันที่ Expire Date ตาม System Date
                Call SP_MSGnShowing("tms_CN005079", nCS_Warning)
                bW_FoundMallCard = False
                otbCusID.Text = ""
                orbMonth.Text = ""
                orbYear.Text = ""
                olaName.Caption = ""
                olaRate.Caption = ""
                Call SP_OTBxSelText(otbCusID)
                Call SP_CRTxLogMallCardEvent("MSR", "DataEvent[7]", IIf(tVB_CNTrnDiscType <> "1850", "Discount", "Mall Card Register"), Trim(otbCusID.Text))
                If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True          '*Ao 10-02-09
                Exit Sub
            End If
        Else
            If orbYear.Text = "" Or orbMonth.Text = "" Then
                Call SP_OBJxSetFocus(orbMonth)
                Call SP_CRTxLogMallCardEvent("MSR", "DataEvent[8]", IIf(tVB_CNTrnDiscType <> "1850", "Discount", "Mall Card Register"), Trim(otbCusID.Text))
                If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True          '*Ao 10-02-09
                Exit Sub
            End If
        End If
        
'        If nW_TrackNo = 4 Then 'JIS
'            If tVB_IsMemJIS = "N" Then
'                Call SP_MSGnShowing("tms_CN005238", nCS_Error)
'                bW_FoundMallCard = False
'                otbCusID.Text = ""                      '*AO51-03-20    Comm.Sheet  SF-108
'                orbMonth.Text = ""
'                orbYear.Text = ""
'                olaName.Caption = ""
'                olaRate.Caption = ""
'                Call SP_OTBxSelText(otbCusID)
'                If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True          '*Ao 10-02-09
'                Exit Sub
'            End If
'        End If
        Call SP_CRTxLogMallCardEvent("MSR", "DataEvent[9]", IIf(tVB_CNTrnDiscType <> "1850", "Discount", "Mall Card Register"), Trim(otbCusID.Text))
    End With
    If (tVB_CNTrnDiscType = "1850") Then tVB_MallCardIDRef = ""   '*Tao 60-04-20 Msr ถือว่าไม่ใช่ Manual (ถ้าถือว่าเป้น Manual ให้ใส่ *)
    Call otbCusID_KeyPress(13)
    Call SP_OBJxSetFocus(ocmOk)
    If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True          '*Ao 10-02-09
    Exit Sub
ErrHandle:
    bW_FoundMallCard = False
    Call SP_MSGxShowError(Err.Number, Err.Description)
    Call SP_CRTxLogMallCardEvent("MSR", "DataEvent[10]", IIf(tVB_CNTrnDiscType <> "1850", "Discount", "Mall Card Register"), Trim(otbCusID.Text))
    Call SP_OBJxSetFocus(otbCusID)
    If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True      '*Ao 10-02-09
End Sub

Private Sub orbMonth_GotFocus()
    'orbMonth.SelStart = 0
    'orbMonth.SelLength = Len(orbMonth.Text)
    Call SP_OTBxSelText(orbMonth)
End Sub

Private Sub orbMonth_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDelete         'delete
            orbMonth.Text = ""
        Case 13
            If Trim(orbMonth.Text) <> "" Then
                SendKeys "{TAB}"
            End If
    End Select
End Sub

Private Sub orbMonth_KeyPress(KeyAscii As Integer)
    Call UT_ChkKeyAscii(KeyAscii)
    Select Case KeyAscii
        Case 8, 6           'backspace, delete
            'do nothing
        Case 13
            'SendKeys "{Tab}"
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
    Call SP_OTBxSelText(orbYear)
End Sub

Private Sub orbYear_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDelete         'delete
            orbYear.Text = ""
        Case 13
            If Trim(orbYear.Text) <> "" Then
                SendKeys "{TAB}"
            End If
    End Select
End Sub

Private Sub orbYear_KeyPress(KeyAscii As Integer)
    Call UT_ChkKeyAscii(KeyAscii)
    Select Case KeyAscii
        Case 8, 6           'backspace, delete
            'do nothing
        Case 13
            '***KT 51-12-01 V 1.0.1
            '----------------------------------------------------------------
            If bW_KeyEnter Then
                KeyAscii = 0
            Else
                bW_KeyEnter = True
                Call ocmOK_Click
            End If
            '----------------------------------------------------------------
        Case Else
            If Not IsNumeric(Chr(KeyAscii)) Then
                KeyAscii = 0
            End If
    End Select
End Sub

Private Sub oScanner_DataEvent(ByVal Status As Long)
On Error GoTo ErrHandle
    With OPOSMSR1                                   '*Ao 10-02-09
        If bVB_OPOSMSREnable And bW_EnableMSR Then
            'Enable the event.
            .DataEventEnabled = False
        End If
    End With
    With oScanner
        'otbCusID.Text = .ScanData '*Bump 56-02-28 Comment บรรทัด นี้
        '*Bump 56-02-28 เพิ่ม code 8 บรรทัด
        If .DeviceName = tSNR_TOUCH_SCANNER And Not .DecodeData Then
            .DecodeData = True
        End If
        If .DeviceName = tSNR_TOUCH_SCANNER Then
            otbCusID.Text = SP_SplitBarcodeData(.ScanDataType, .ScanData, .ScanDataLabel) '*ต้องย้ายจาก mCNSTP.SP_SplitBarcodeData  ไปที่ mCNSP.SP_SplitBarcodeData
        Else
            otbCusID.Text = .ScanData
        End If
'        DoEvents        '*KT 53-01-14
        Call SP_CRTxLogMallCardEvent("SCANNER", "DataEvent[1]", IIf(tVB_CNTrnDiscType <> "1850", "Discount", "Mall Card Register"), Trim(otbCusID.Text))
        If otbCusID.Text <> "" Then
            If Asc(Right(otbCusID.Text, 1)) = 13 Then
                otbCusID.Text = Mid(otbCusID.Text, 1, Len(otbCusID.Text) - 1)
            End If
            Call SP_CRTxLogMallCardEvent("SCANNER", "DataEvent[2]", IIf(tVB_CNTrnDiscType <> "1850", "Discount", "Mall Card Register"), Trim(otbCusID.Text))
            Call otbCusID_KeyPress(13)
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
    Call SP_CRTxLogMallCardEvent("SCANNER", "DataEvent[3]", IIf(tVB_CNTrnDiscType <> "1850", "Discount", "Mall Card Register"), Trim(otbCusID.Text))
    If bVB_OPOSMSREnable And bW_EnableMSR Then OPOSMSR1.DataEventEnabled = True       '*Ao 10-02-09
    Exit Sub
ErrHandle:
    Call SP_CRTxLogMallCardEvent("SCANNER", "DataEvent[4]", IIf(tVB_CNTrnDiscType <> "1850", "Discount", "Mall Card Register"), Trim(otbCusID.Text))
    MsgBox "Scaner error" & vbCrLf & "Runtime : " & Err.Number & vbCrLf & Err.Description
    If bVB_OPOSMSREnable And bW_EnableMSR Then OPOSMSR1.DataEventEnabled = True      '*Ao 10-02-09
End Sub

Private Sub otbCusID_Change()
    bW_FoundMallCard = False
End Sub

Private Sub otbCusID_KeyPress(KeyAscii As Integer)
   Dim tSta As String
   Dim tSql As String
   Dim tSplash As String
   Dim nErr As Long
   Dim tTokenCardNo As String
   Dim tCardNo As String
   Dim tdbToken As String
On Error GoTo ErrHandle
    Call UT_ChkKeyAscii(KeyAscii)
    Call SP_ChkKeyAsciiNumber(KeyAscii)
    If otbCusID.Text <> "" Then
        If KeyAscii = 13 Or KeyAscii = vbKeyTab Then
             Call SP_CRTxLogMallCardEvent(IIf(bW_FoundMallCard, "MSR", "Keyboard"), "KeyPress[1]", IIf(tVB_CNTrnDiscType <> "1850", "Discount", "Mall Card Register"), Trim(otbCusID.Text))
             tVB_DiscMallCardID = Trim(otbCusID.Text)
            '===================================
            '*Bmp 61-02-15 Token
            tTokenCardNo = Trim(otbCusID.Text)
            tdbToken = ""
            '*Tao 61-10-05 ไม่ต้องค้นหาจาก MemServer
'            If SP_PRCbTokenMemSrv(IIf(tVB_CNTrnDiscType <> "1850", "Discount", "BPAward"), Trim(otbCusID.Text), tdbToken) Then
'                tCardNo = tdbToken
'            Else
                If SP_PRCbToken(IIf(tVB_CNTrnDiscType <> "1850", "Discount", "BPAward"), Trim(otbCusID.Text), tTokenCardNo) Then
                    tCardNo = tTokenCardNo
                Else
                    tCardNo = Trim(otbCusID.Text)
                End If
'            End If
            '===================================
             If bW_FoundMallCard Then
                    'ทาง FSBT กำหนดว่ากรณีที่ Flag ใน Mall Card Maintanance นั้น Tik Check Expiredate ให้ตรวจสอบวันที่หมดอายุเฉพาะวันที่ที่ได้จากบัตรและตรวจสอบกับ System Date
                    'ส่วนกรณีที่เป็น N ไม่ต้องตรวจสอบวันที่ Expire ที่บัตร แต่ให้ไปตรวจสถานะใน MemberServer เท่านั้น ไม่สนใจวันที่หมดอายุ
                    'ยืนยันการแก้ไขจากพี่บัม วันที่ 30/06/2009
                     '*KT 52-03-16
                    Call SP_CRTxLogMallCardEvent("MSR", "KeyPress[2]", IIf(tVB_CNTrnDiscType <> "1850", "Discount", "Mall Card Register"), Trim(otbCusID.Text))
                    If ocnMember Is Nothing And bVB_ReConnectMem Then
                        Call SP_STAxMouse(Me, True)
                        tSplash = "Can not connect the database member, System is  reconnect database member again "
                        '*KT 52-05-26 PH 1.8 ML-SF-075 ในกรณีที่ POS ไม่สามารถติดต่อ Member Server ได้ ขอให้ POS แสดง Warning Message เพียงครั้งเดียว
                        'Call SP_CONbMemberServer(True, tSplash, False) '*Bump 56-10-11 V.4.0.79
                        '*Bump 56-10-11 V.4.0.79  ถ้า Online จึงให้ทำ Call SP_CONbMemberServer(True, tSplash, False) CommSheet TK-ISS3000-368
                        If oDB.SQLbAvailableServer(oDB.OffServerName, oDB.OnComIP) Then Call SP_CONbMemberServer(True, tSplash, False)
                        Call SP_STAxMouse(Me, False)
                        bVB_ReConnectMem = False
                        Call SP_CRTxLogMallCardEvent("MSR", "KeyPress[3]", IIf(tVB_CNTrnDiscType <> "1850", "Discount", "Mall Card Register"), Trim(otbCusID.Text))
                    End If
                    '*KT 51-08-07 Communication Sheet ML-SF-436 Member Card Status.xls
                      '----------------------------------------------------------------------------
                    If Not ocnMember Is Nothing And oDB.SQLbAvailableServer(oDB.OffServerName, oDB.OnComIP) Then
On Error GoTo ErrHandle2
                           Call SP_STAxMouse(Me, True)
                           Call SP_CRTxLogMallCardEvent("MSR", "KeyPress[4]", IIf(tVB_CNTrnDiscType <> "1850", "Discount", "Mall Card Register"), Trim(otbCusID.Text))
'                           tSql = "SELECT Top 1  FTMcdCode,FTMcdSta,FDMcdExpDate,FCMcdTotalPoint,FTMcdName,FTMcdSurName "
                           tSql = "SELECT *"                                                                    '*Ao 2013/02/12    CommSheet TK-ISS3000-017 00.00.03 Replied by JESLIN_Aoo.xls
                           tSql = tSql & vbCrLf & "FROM TCNMMallCard"
                           tSql = tSql & vbCrLf & "WHERE FTMcdCode ='" & tCardNo & "'" & " ORDER BY FDMcdExpDate DESC" '*Bmp 61-02-15 Token '*Bump 55-12-26 เพิ่ม ORDER BY FDMcdExpDate DESC ในกรณีที่มีบัตรสมาชิกมากกว่า 1 ใบ ให้เอารายการที่วันหมดอายุมากที่สุดมาใช้งาน
                           ''tSql = tSql & vbCrLf & "WHERE FTMcdCode ='" & Trim(otbCusID.Text) & "'" & " ORDER BY FDMcdExpDate DESC" '*Bump 55-12-26 เพิ่ม ORDER BY FDMcdExpDate DESC ในกรณีที่มีบัตรสมาชิกมากกว่า 1 ใบ ให้เอารายการที่วันหมดอายุมากที่สุดมาใช้งาน
                           nErr = SP_SQLvExecute(tSql, orsMem, ocnMember, , , , True)
                            Select Case nErr
                                Case -2147467259, 91, -2147217871, 3704 '*KT 54-04-07 ป้องกันโปรแกรมหลุด 3.0.56
                                    '*KT 52-05-26 PH 1.8 ML-SF-075 ในกรณีที่ POS ไม่สามารถติดต่อ Member Server ได้ ขอให้ POS แสดง Warning Message เพียงครั้งเดียว
                                    '*KT 52-05-26 ให้ทำการแจ้ง message เตือนเมื่อทำการ Connect Member Server
                                    'Call SP_MSGnShowing("tms_CN005134", tCS_CNMsgErr)
                                    Call SP_CRTxLogMallCardEvent("MSR", "KeyPress[5]", IIf(tVB_CNTrnDiscType <> "1850", "Discount", "Mall Card Register"), Trim(otbCusID.Text))
                            End Select
                           If Not (orsMem.BOF Or orsMem.EOF) Then
                                Call SP_CRTxLogMallCardEvent("MSR", "KeyPress[6]", IIf(tVB_CNTrnDiscType <> "1850", "Discount", "Mall Card Register"), Trim(otbCusID.Text))
'                                If UCase(tW_ChkExp) = "N" Then
                                   Call SP_CRTxLogMallCardEvent("MSR", "KeyPress[7]", IIf(tVB_CNTrnDiscType <> "1850", "Discount", "Mall Card Register"), Trim(otbCusID.Text))
                                   tSta = SP_FEDtChkString(orsMem, "FTMcdSta")
                                   If UCase(tSta) = "A" Or UCase(tSta) = "1" Then
                                       bW_FoundMallCard = True
                                       '*KT 52-11-26 PH2-ML-SF-170 ใน Ticket ไม่พิมพ์ชื่อพนักงานที่ใช้ส่วนลด Employee discount
                                       tW_CardName = SP_FEDtChkString(orsMem, "FTMcdName") & " " & SP_FEDtChkString(orsMem, "FTMcdSurName")
                                       If Trim(tW_CardName) = "" Then tW_CardName = "New Member"
                                       '*Tao 60-06-02 ใส่ Expire Date
                                        If Not IsNull(orsMem.Fields("FDMcdExpDate")) Then
                                            If CDate(orsMem.Fields("FDMcdExpDate")) >= Date Then
                                                tW_CardExp = Format(CDate(orsMem.Fields("FDMcdExpDate")), "MMYY")
                                                'orbMonth.Text = Left(tW_CardExp, 2)
                                                'orbYear.Text = Mid(tW_CardExp, 3, 2)
                                            End If
                                        End If
                                        '========================
                                       Call SP_CRTxLogMallCardEvent("MSR", "KeyPress[8]", IIf(tVB_CNTrnDiscType <> "1850", "Discount", "Mall Card Register"), Trim(otbCusID.Text))
                                   Else
                                       bW_FoundMallCard = False
                                       Select Case UCase(tSta)
                                           Case "E", "2"
                                               Call SP_MSGnShowing("tms_CN005172", nCS_Information)
                                           Case "L", "3"
                                               Call SP_MSGnShowing("tms_CN005173", nCS_Information)
                                           Case "H", "4"
                                               Call SP_MSGnShowing("tms_CN005174", nCS_Information)
                                       End Select
                                       bW_FoundMallCard = False
                                       tW_CardName = "" '*KT 52-11-26 PH2-ML-SF-170 ใน Ticket ไม่พิมพ์ชื่อพนักงานที่ใช้ส่วนลด Employee discount
                                       otbCusID.Text = ""                      '*AO51-03-20    Comm.Sheet  SF-108
                                       orbMonth.Text = ""
                                       orbYear.Text = ""
                                       olaName.Caption = ""
                                       olaRate.Caption = ""
                                       Call SP_OTBxSelText(otbCusID)
                                       Call SP_CRTxLogMallCardEvent("MSR", "KeyPress[9]", IIf(tVB_CNTrnDiscType <> "1850", "Discount", "Mall Card Register"), Trim(otbCusID.Text))
                                   End If
'                                End If
                           Else
                               bW_FoundMallCard = True
                               tW_CardName = "New Member" ''*KT 52-11-26 PH2-ML-SF-170 ใน Ticket ไม่พิมพ์ชื่อพนักงานที่ใช้ส่วนลด Employee discount
                               Call SP_CRTxLogMallCardEvent("MSR", "KeyPress[10]", IIf(tVB_CNTrnDiscType <> "1850", "Discount", "Mall Card Register"), Trim(otbCusID.Text))
                           End If
                           Call SP_STAxMouse(Me, False)
                   Else
                      Set ocnMember = Nothing
                      bVB_ReConnectMem = False
                      bW_FoundMallCard = True
                      Call SP_CRTxLogMallCardEvent("MSR", "KeyPress[11]", IIf(tVB_CNTrnDiscType <> "1850", "Discount", "Mall Card Register"), Trim(otbCusID.Text))
                   End If
                   '---------------------------------------------------------------------------
            Else
                Call SP_CRTxLogMallCardEvent("Keyboard", "KeyPress[12]", IIf(tVB_CNTrnDiscType <> "1850", "Discount", "Mall Card Register"), Trim(otbCusID.Text))
               If Not W_PRCxFindMallCardByID(tVB_DiscMallCardID) Then
'                   Call SP_MSGnShowing("tms_CN005068", nCS_Error)           '*Ao 2013/02/04
                   bW_FoundMallCard = False
                   tW_CardName = "" '*KT 52-11-26 PH2-ML-SF-170 ใน Ticket ไม่พิมพ์ชื่อพนักงานที่ใช้ส่วนลด Employee discount
                   otbCusID.Text = ""                      '*AO51-03-20    Comm.Sheet  SF-108
                   orbMonth.Text = ""
                   orbYear.Text = ""
                   olaName.Caption = ""
                   olaRate.Caption = ""
                   If (tVB_CNTrnDiscType = "1850") Then tVB_MallCardIDRef = "" '*Tao 60-04-27
                   Call SP_OTBxSelText(otbCusID)
                   Call SP_CRTxLogMallCardEvent("Keyboard", "KeyPress[13]", IIf(tVB_CNTrnDiscType <> "1850", "Discount", "Mall Card Register"), Trim(otbCusID.Text))
               Else
                    '*KT 52-09-30 PH2-ML-SF-149 หมายเลขบัตร Mall Card ไม่ตรงกัน ระหว่าง M-Price กับ รูด Mall Card (Mall8)
                    If tW_IsChkDigitMallCard = "1" Then
                        If Not SP_PRCbIsChkDigitMallCard(tVB_DiscMallCardID) Then
                            Call SP_MSGnShowing("tms_CN005216", nCS_Warning)
                            bW_FoundMallCard = False
                            tW_CardName = "" '*KT 52-11-26 PH2-ML-SF-170 ใน Ticket ไม่พิมพ์ชื่อพนักงานที่ใช้ส่วนลด Employee discount
                            otbCusID.Text = ""
                            orbMonth.Text = ""
                            orbYear.Text = ""
                            olaName.Caption = ""
                            olaRate.Caption = ""
                            If (tVB_CNTrnDiscType = "1850") Then tVB_MallCardIDRef = "" '*Tao 60-04-27
                            Call SP_OTBxSelText(otbCusID)
                            Call SP_CRTxLogMallCardEvent("Keyboard", "KeyPress[14]", IIf(tVB_CNTrnDiscType <> "1850", "Discount", "Mall Card Register"), Trim(otbCusID.Text))
                            Exit Sub
                        End If
                    End If
                    '*Tao 60-06-02 ใส่ Expire Date
                    If tW_CardExp <> "" Then
                        'orbMonth.Text = Left(tW_CardExp, 2)
                        'orbYear.Text = Mid(tW_CardExp, 3, 2)
                    End If
                    '========================
                    bW_FoundMallCard = True
                    'ทาง FSBT กำหนดว่ากรณีที่ Flag ใน Mall Card Maintanance นั้น Tik Check Expiredate ให้ตรวจสอบวันที่หมดอายุเฉพาะวันที่ที่ได้จากบัตรและตรวจสอบกับ System Date
                    'ส่วนกรณีที่เป็น N ไม่ต้องตรวจสอบวันที่ Expire ที่บัตร แต่ให้ไปตรวจสถานะใน MemberServer เท่านั้น ไม่สนใจวันที่หมดอายุ
                    'ยืนยันการแก้ไขจากพี่บัม วันที่ 30/06/2009
                    '*KT 52-03-16
                    Call SP_CRTxLogMallCardEvent("Keyboard", "KeyPress[15]", IIf(tVB_CNTrnDiscType <> "1850", "Discount", "Mall Card Register"), Trim(otbCusID.Text))
                    If ocnMember Is Nothing And bVB_ReConnectMem Then
                        Call SP_STAxMouse(Me, True)
                        tSplash = "Can not connect the database member, System is  reconnect database member again "
                        '*KT 52-05-26 PH 1.8 ML-SF-075 ในกรณีที่ POS ไม่สามารถติดต่อ Member Server ได้ ขอให้ POS แสดง Warning Message เพียงครั้งเดียว
                        'Call SP_CONbMemberServer(True, tSplash, False) '*Bump 56-10-11 V.4.0.79
                        '*Bump 56-10-11 V.4.0.79  ถ้า Online จึงให้ทำ Call SP_CONbMemberServer(True, tSplash, False)    CommSheet TK-ISS3000-368
                        If oDB.SQLbAvailableServer(oDB.OffServerName, oDB.OnComIP) Then Call SP_CONbMemberServer(True, tSplash, False)
                        Call SP_STAxMouse(Me, False)
                        bVB_ReConnectMem = False
                        Call SP_CRTxLogMallCardEvent("Keyboard", "KeyPress[16]", IIf(tVB_CNTrnDiscType <> "1850", "Discount", "Mall Card Register"), Trim(otbCusID.Text))
                    End If
                   '*KT 51-08-07 Communication Sheet ML-SF-436 Member Card Status.xls
                   '----------------------------------------------------------------------------
                   If Not ocnMember Is Nothing And oDB.SQLbAvailableServer(oDB.OffServerName, oDB.OnComIP) Then
On Error GoTo ErrHandle2
                            Call SP_STAxMouse(Me, True)
                            Call SP_CRTxLogMallCardEvent("Keyboard", "KeyPress[17]", IIf(tVB_CNTrnDiscType <> "1850", "Discount", "Mall Card Register"), Trim(otbCusID.Text))
'                           tSql = "SELECT Top 1  FTMcdCode,FTMcdSta,FDMcdExpDate,FCMcdTotalPoint,FTMcdName,FTMcdSurName "
                           tSql = "SELECT *"                                                                    '*Ao 2013/02/12    CommSheet TK-ISS3000-017 00.00.03 Replied by JESLIN_Aoo.xls
                           tSql = tSql & vbCrLf & "FROM TCNMMallCard"
                           tSql = tSql & vbCrLf & "WHERE FTMcdCode ='" & tCardNo & "'" & " ORDER BY FDMcdExpDate DESC" '*Bmp 61-02-15 Token '*Bump 55-12-26 เพิ่ม ORDER BY FDMcdExpDate DESC ในกรณีที่มีบัตรสมาชิกมากกว่า 1 ใบ ให้เอารายการที่วันหมดอายุมากที่สุดมาใช้งาน
                           'tSql = tSql & vbCrLf & "WHERE FTMcdCode ='" & Trim(otbCusID.Text) & "'" & " ORDER BY FDMcdExpDate DESC" '*Bump 55-12-26 เพิ่ม ORDER BY FDMcdExpDate DESC ในกรณีที่มีบัตรสมาชิกมากกว่า 1 ใบ ให้เอารายการที่วันหมดอายุมากที่สุดมาใช้งาน
                           nErr = SP_SQLvExecute(tSql, orsMem, ocnMember, , , , True)
                           Select Case nErr
                                Case -2147467259, 91, -2147217871, 3704 '*KT 54-04-07 ป้องกันโปรแกรมหลุด 3.0.56
                                    '*KT 52-05-26 PH 1.8 ML-SF-075 ในกรณีที่ POS ไม่สามารถติดต่อ Member Server ได้ ขอให้ POS แสดง Warning Message เพียงครั้งเดียว
                                    '*KT 52-05-26 ให้ทำการแจ้ง message เตือนเมื่อทำการ Connect Member Server
                                    'Call SP_MSGnShowing("tms_CN005134", tCS_CNMsgErr)
                                    Call SP_CRTxLogMallCardEvent("Keyboard", "KeyPress[18]", IIf(tVB_CNTrnDiscType <> "1850", "Discount", "Mall Card Register"), Trim(otbCusID.Text))
                            End Select
                           If Not (orsMem.BOF Or orsMem.EOF) Then
                                Call SP_CRTxLogMallCardEvent("Keyboard", "KeyPress[19]", IIf(tVB_CNTrnDiscType <> "1850", "Discount", "Mall Card Register"), Trim(otbCusID.Text))
                                ' If UCase(tW_ChkExp) = "N" Then
                                    Call SP_CRTxLogMallCardEvent("Keyboard", "KeyPress[20]", IIf(tVB_CNTrnDiscType <> "1850", "Discount", "Mall Card Register"), Trim(otbCusID.Text))
                                    tSta = SP_FEDtChkString(orsMem, "FTMcdSta")
                                    If UCase(tSta) = "A" Or UCase(tSta) = "1" Then
                                        bW_FoundMallCard = True
                                        '*KT 52-11-26 PH2-ML-SF-170 ใน Ticket ไม่พิมพ์ชื่อพนักงานที่ใช้ส่วนลด Employee discount
                                        tW_CardName = SP_FEDtChkString(orsMem, "FTMcdName") & " " & SP_FEDtChkString(orsMem, "FTMcdSurName")
                                        If Trim(tW_CardName) = "" Then tW_CardName = "New Member"
                                        '*Tao 60-06-02 ใส่ Expire Date
                                        If Not IsNull(orsMem.Fields("FDMcdExpDate")) Then
                                            If CDate(orsMem.Fields("FDMcdExpDate")) >= Date Then
                                                tW_CardExp = Format(CDate(orsMem.Fields("FDMcdExpDate")), "MMYY")
                                                'orbMonth.Text = Left(tW_CardExp, 2)
                                                'orbYear.Text = Mid(tW_CardExp, 3, 2)
                                            End If
                                        End If
                                        '========================
                                        Call SP_CRTxLogMallCardEvent("Keyboard", "KeyPress[21]", IIf(tVB_CNTrnDiscType <> "1850", "Discount", "Mall Card Register"), Trim(otbCusID.Text))
                                    Else
                                        bW_FoundMallCard = False
                                        Select Case UCase(tSta)
                                            Case "E", "2"
                                                Call SP_MSGnShowing("tms_CN005172", nCS_Information)
                                            Case "L", "3"
                                                Call SP_MSGnShowing("tms_CN005173", nCS_Information)
                                            Case "H", "4"
                                                Call SP_MSGnShowing("tms_CN005174", nCS_Information)
                                        End Select
                                        bW_FoundMallCard = False
                                        tW_CardName = "" '*KT 52-11-26 PH2-ML-SF-170 ใน Ticket ไม่พิมพ์ชื่อพนักงานที่ใช้ส่วนลด Employee discount
                                        otbCusID.Text = ""                      '*AO51-03-20    Comm.Sheet  SF-108
                                        orbMonth.Text = ""
                                        orbYear.Text = ""
                                        olaName.Caption = ""
                                        olaRate.Caption = ""
                                        If (tVB_CNTrnDiscType = "1850") Then tVB_MallCardIDRef = "" '*Tao 60-04-27
                                        Call SP_OTBxSelText(otbCusID)
                                        Call SP_CRTxLogMallCardEvent("Keyboard", "KeyPress[22]", IIf(tVB_CNTrnDiscType <> "1850", "Discount", "Mall Card Register"), Trim(otbCusID.Text))
                                   ' End If
                                End If
                           Else
                               bW_FoundMallCard = True
                               tW_CardName = "New Member" '*KT 52-11-26 PH2-ML-SF-170 ใน Ticket ไม่พิมพ์ชื่อพนักงานที่ใช้ส่วนลด Employee discount
                               Call SP_CRTxLogMallCardEvent("Keyboard", "KeyPress[23]", IIf(tVB_CNTrnDiscType <> "1850", "Discount", "Mall Card Register"), Trim(otbCusID.Text))
                           End If
                           Call SP_STAxMouse(Me, False)
                   Else
                        Set ocnMember = Nothing
                        bVB_ReConnectMem = False
                       bW_FoundMallCard = True
                       tW_CardName = "New Member" '*KT 52-11-26 PH2-ML-SF-170 ใน Ticket ไม่พิมพ์ชื่อพนักงานที่ใช้ส่วนลด Employee discount
                       Call SP_CRTxLogMallCardEvent("Keyboard", "KeyPress[24]", IIf(tVB_CNTrnDiscType <> "1850", "Discount", "Mall Card Register"), Trim(otbCusID.Text))
                   End If
                   '---------------------------------------------------------------------------
               End If
               Call SP_CRTxLogMallCardEvent("Keyboard", "KeyPress[25]", IIf(tVB_CNTrnDiscType <> "1850", "Discount", "Mall Card Register"), Trim(otbCusID.Text))
            End If
        End If
    End If
    Exit Sub
ErrHandle:
    Call SP_MSGxShowError(Err.Number, Err.Description)
    Call SP_CRTxLogMallCardEvent("Keyboard", "KeyPress[26]", IIf(tVB_CNTrnDiscType <> "1850", "Discount", "Mall Card Register"), Trim(otbCusID.Text))
    Call SP_STAxMouse(Me, False)
    Exit Sub
ErrHandle2:
    Select Case Err.Number
        Case -2147467259, 91, -2147217871, 3704 '*KT 54-04-07 ป้องกันโปรแกรมหลุด 3.0.56
            Set ocnMember = Nothing
            bVB_ReConnectMem = False
        Case Else
            Call SP_MSGxShowError(Err.Number, Err.Description)
    End Select
    bW_FoundMallCard = True
    Call SP_CRTxLogMallCardEvent("Keyboard", "KeyPress[27]", IIf(tVB_CNTrnDiscType <> "1850", "Discount", "Mall Card Register"), Trim(otbCusID.Text))
    Call SP_STAxMouse(Me, False)
End Sub

Private Sub otbCusID_GotFocus()
    otbCusID.SelStart = 0
    otbCusID.SelLength = Len(otbCusID.Text)
    If bVB_Discount91 Then    '*Eaw 56-01-05
        Call SP_OBJxSetFocus(ocmOk)
    End If
End Sub

Private Sub otbCusID_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDelete         'delete
            otbCusID.Text = ""
            orbMonth.Text = ""
            orbYear.Text = ""
            olaName.Caption = ""
            olaRate.Caption = ""
            If (tVB_CNTrnDiscType = "1850") Then tVB_MallCardIDRef = ""  '*Tao 60-04-20
        Case 13                 '*Ao 53-01-14
            If otbCusID.Text <> "" Then
                SendKeys "{TAB}"
            End If
        Case 48, 49, 50, 51, 52, 53, 54, 55, 56, 57, 96, 97, 98, 99, 100, 101, 102, 103, 104, 105
            If (Shift = 0) And (tVB_CNTrnDiscType = "1850") Then tVB_MallCardIDRef = "*" '*Tao 60-04-20 Manual Bonus Point Award
    End Select
End Sub

Private Sub otbCusID_LostFocus()
    otbCusID.Text = SP_TXTtTrimRightLeft(otbCusID.Text)
    If Not bW_FoundMallCard Then
        Call otbCusID_KeyPress(13)
    End If
End Sub

Private Sub W_PRCxPointMallCardEx()

'-----------------------------------------------------------------------------
'Ret:
'Cmt:   pocess point หาค่า last point ที่ member server
'*KT 52-02-18
'กรณีการแสดง last point ใน receipt/ej
'1 Connect member server ได้
'2 mall card flag member point
'3 parameter flag ให้แสดง
'----------------------------------------------------------------------------
    Dim tSql$
    Dim cLastPoint As Double
    Dim tMemberName As String
    Dim orsTemp As ADODB.Recordset
    Dim orsTemp2 As ADODB.Recordset
    Dim orsTemp3 As ADODB.Recordset
    Dim cPoint As Double
    Dim cPointAmt As Double
    Dim bPass As Boolean
    Dim tSta As String
    Dim bProcess As Boolean
On Error GoTo ErrHadle
    bProcess = False
    'ทาง FSBT กำหนดว่ากรณีที่ Flag ใน Mall Card Maintanance นั้น Tik Check Expiredate ให้ตรวจสอบวันที่หมดอายุเฉพาะวันที่ที่ได้จากบัตรและตรวจสอบกับ System Date
    'ส่วนกรณีที่เป็น N ไม่ต้องตรวจสอบวันที่ Expire ที่บัตร แต่ให้ไปตรวจสถานะใน MemberServer เท่านั้น ไม่สนใจวันที่หมดอายุ
    'ซึ่งในส่วนนี้เป็นการเก็บหมายเลขสมาชิกไว้สำหรับไปหา point จึงนำเงื่อนเช่นเดียวกันกับ Mall Card มาใช้ด้วยกัน
    'ยังไม่ยืนยันการแก้ไขจากพี่บัม วันที่ 30/06/2009
    'If tVB_StaMemPointAlw = "Y" Or tStaPointSpecial = "Y" Then
    If tVB_StaMemPointAlw = "Y" Then  'PH2-ML-SF-097
        If Not ocnMember Is Nothing And oDB.SQLbAvailableServer(oDB.OffServerName, oDB.OnComIP) Then
            If orsMem.State = adStateOpen Then
                If Not (orsMem.BOF And orsMem.EOF) Then
                         If Format(SP_FEDdChkDate(orsMem, "FDMcdExpDate"), "YYYY/MM/DD") >= Format(dVB_CurrentDate, "YYYY/MM/DD") Then 'ทาง FSBT ต้องการให้ตรวจสอบเรื่องวันที่ Expire Date จากวันที่ System Date
                             tSta = SP_FEDtChkString(orsMem, "FTMcdSta")
                            If UCase(tSta) = "A" Or UCase(tSta) = "1" Then
'                                cLastPoint = SP_FEDvChkNumeric(orsMem, "FCMcdTotalPoint")
                                cLastPoint = SP_FEDvChkNumeric(orsMem, "FCBalance")           '*Ao 2013/02/12    CommSheet TK-ISS3000-017 00.00.03 Replied by JESLIN_Aoo.xls
                                tMemberName = SP_FEDtChkString(orsMem, "FTMcdName") & " " & SP_FEDtChkString(orsMem, "FTMcdSurName")
                                bProcess = True
                            Else
                                bW_FoundMallCard = False
                                Select Case UCase(tSta)
                                    Case "E", "2"
                                        cLastPoint = 0
                                        tMemberName = ""
                                    Case "L", "3"
                                        cLastPoint = 0
                                        tMemberName = ""
                                    Case "H", "4"
                                        cLastPoint = 0
                                        tMemberName = ""
                                End Select
                            End If
                        Else
                            cLastPoint = 0
                            tMemberName = ""
                        End If
                Else
                    cLastPoint = 0
                    tMemberName = "New member"
                    bProcess = True
                End If
            Else
                cLastPoint = 0
                tMemberName = ""
            End If
        Else
            cLastPoint = 0
            tMemberName = ""
        End If
    Else
        cLastPoint = 0
        tMemberName = ""
    End If
    
    
'    '*KT 52-02-24 member point process คำนวณ point
'    '---------------------------------------------------------------------------------------------

    cPoint = 0
    cPointAmt = 0
    
    Select Case nVB_DiscMallCount
        Case 1
            ReDim aVB_MallCard(0)
            aVB_MallCard(0).ID = IIf(tVB_StaMemPointAlw = "Y", tVB_DiscMallCardID, "")
            aVB_MallCard(0).Type = tVB_DiscMallCardKey
            aVB_MallCard(0).Point = cPoint
'            aVB_MallCard(0).LPoint = IIf(tVB_StaMemPointAlw = "Y", cLastPoint + cPoint, 0) '*KT 55-09-17 Total Bonus Points
            aVB_MallCard(0).LPoint = cLastPoint         '*Ao 2013/02/12    CommSheet TK-ISS3000-017 00.00.03 Replied by JESLIN_Aoo.xls
            aVB_MallCard(0).Amt = cPointAmt
            aVB_MallCard(0).IDLastPoint = IIf(tVB_StaMemPointAlw = "Y", tVB_DiscMallCardID, "")
            aVB_MallCard(0).CardID = tVB_DiscMallCardID
            aVB_MallCard(0).tEvnCode = tVB_MemEvnCode '*KT 54-11-18 V.4.0.5  DCR_5411003 PRJ09010TheMall-R4-SE-External Spec.Upload Bonus Points-00.04.00-FSBT.doc
            aVB_MallCard(0).tFTMcdCardType = tVB_DiscMallCardType '*Bmp 61-02-15
        Case 2, 3
            ReDim Preserve aVB_MallCard(UBound(aVB_MallCard) + 1)
            aVB_MallCard(nVB_DiscMallCount - 1).ID = IIf(tVB_StaMemPointAlw = "Y", tVB_DiscMallCardID, "")
            aVB_MallCard(nVB_DiscMallCount - 1).Type = tVB_DiscMallCardKey
            aVB_MallCard(nVB_DiscMallCount - 1).Point = cPoint
'            aVB_MallCard(nVB_DiscMallCount - 1).LPoint = IIf(tVB_StaMemPointAlw = "Y", cLastPoint + cPoint, 0) '*KT 55-09-17 Total Bonus Points
            aVB_MallCard(nVB_DiscMallCount - 1).LPoint = cLastPoint         '*Ao 2013/02/12    CommSheet TK-ISS3000-017 00.00.03 Replied by JESLIN_Aoo.xls
            aVB_MallCard(nVB_DiscMallCount - 1).Amt = cPointAmt
            aVB_MallCard(nVB_DiscMallCount - 1).IDLastPoint = IIf(tVB_StaMemPointAlw = "Y", tVB_DiscMallCardID, "")
            aVB_MallCard(nVB_DiscMallCount - 1).CardID = tVB_DiscMallCardID
            aVB_MallCard(nVB_DiscMallCount - 1).tEvnCode = tVB_MemEvnCode '*KT 54-11-18 V.4.0.5  DCR_5411003 PRJ09010TheMall-R4-SE-External Spec.Upload Bonus Points-00.04.00-FSBT.doc
            aVB_MallCard(nVB_DiscMallCount - 1).tFTMcdCardType = tVB_DiscMallCardType '*Bmp 61-02-15
        Case Else
            Exit Sub
    End Select
    
'*KT 52-12-29 แก้ไชเงื่อนไขการให้แต้มกับบัตรใบแรกที่มีคุณสมบัติให้แต้มใบแรกตาม DCR  PRJ09010TheMall-R3-DCR_12001(Member Point)-01.00.00.xls
    If aVB_MallCard(nVB_DiscMallCount - 1).ID <> "" And (tVB_StaMemPointAlw = "Y") And tVB_MemPointID = "" Then 'PH2-ML-SF-097
        tSql = "UPDATE " & tVB_CNTblHD
        tSql = tSql & " SET FTCstCardCode ='" & aVB_MallCard(nVB_DiscMallCount - 1).IDLastPoint & "'"
        tSql = tSql & " ,FTCstCardType = '" & aVB_MallCard(nVB_DiscMallCount - 1).Type & "'"
        tSql = tSql & " ,FNCstCardPoint = " & CLng(aVB_MallCard(nVB_DiscMallCount - 1).Point)
        tSql = tSql & " ,FNCstCardLPoint = " & CLng(aVB_MallCard(nVB_DiscMallCount - 1).LPoint)
        tSql = tSql & " ,FCCstCardAmt = " & aVB_MallCard(nVB_DiscMallCount - 1).Amt
        tSql = tSql & " ,FTCstPointName = '" & SP_STRtSiggleCode(tMemberName) & "'" '*KT 53-02-10 Communication Sheet PH3-ML-SF-079 Can't read Mall card which customer name or surname included character'.xls
        tSql = tSql & " ,FTEcdCardType = '" & aVB_MallCard(nVB_DiscMallCount - 1).tFTMcdCardType & "'" '*Bmp 61-02-15 PRJ10001TK-PH2-External Spec-Tokenization-00.02.06 to TK.docx
        tSql = tSql & "," & UT_tLastUpdSql
        tSql = tSql & " WHERE  FTTmnNum='" & tVB_CNTerminalNum & "' AND FTShdTransNo='" & tVB_Running & "'"
        tSql = tSql & " AND  FTShdTransType='" & tVB_Transaction & "' AND FDShdTransDate='" & tUT_SaleDate & "'"
        Call SP_SQLbExecute(tSql)
        tVB_MemPointID = aVB_MallCard(nVB_DiscMallCount - 1).IDLastPoint
        nVB_CurMemPointCount = nVB_DiscMallCount
        nVB_MemPointMnu = 1
        Call SP_CRTxLogMallCardEvent("MSR/Keyboard", "MemberPoint[1]", IIf(tVB_CNTrnDiscType <> "1850", "Discount", "Mall Card Register"), Trim(otbCusID.Text))
    End If
    Call SP_DATxRsNothing(orsTemp)
    Call SP_DATxRsNothing(orsTemp2)
    Call SP_DATxRsNothing(orsTemp3)
    Exit Sub
ErrHadle:
    Call SP_MSGxShowError(Err.Number, Err.Description)
    Call SP_CRTxLogMallCardEvent("MSR/Keyboard", "MemberPoint[2]", IIf(tVB_CNTrnDiscType <> "1850", "Discount", "Mall Card Register"), Trim(otbCusID.Text))
    Call SP_DATxRsNothing(orsTemp)
    Call SP_DATxRsNothing(orsTemp2)
    Call SP_DATxRsNothing(orsTemp3)
End Sub

Private Sub W_PRCxRegMallCard()
    'pocess point
    Dim tSql$
    Dim cLastPoint As Double
    Dim tMemberName As String
    Dim orsTemp As ADODB.Recordset
    Dim orsTemp2 As ADODB.Recordset
    Dim orsTemp3 As ADODB.Recordset
    Dim cPoint As Double
    Dim cPointAmt As Double
    Dim bPass As Boolean
    Dim tSta As String
    Dim bProcess As Boolean
On Error GoTo ErrHadle
    bProcess = False
    'ทาง FSBT กำหนดว่ากรณีที่ Flag ใน Mall Card Maintanance นั้น Tik Check Expiredate ให้ตรวจสอบวันที่หมดอายุเฉพาะวันที่ที่ได้จากบัตรและตรวจสอบกับ System Date
    'ส่วนกรณีที่เป็น N ไม่ต้องตรวจสอบวันที่ Expire ที่บัตร แต่ให้ไปตรวจสถานะใน MemberServer เท่านั้น ไม่สนใจวันที่หมดอายุ
    'ซึ่งในส่วนนี้เป็นการเก็บหมายเลขสมาชิกไว้สำหรับไปหา point  ในขั้นตอนการรูด Mall Card Register จึงนำเงื่อนเช่นเดียวกันกับ Mall Card มาใช้ด้วยกัน
    'ยังไม่ยืนยันการแก้ไขจากพี่บัม วันที่ 30/06/2009
    If tVB_StaMemPointAlw = "Y" Then  'PH2-ML-SF-097
        If Not ocnMember Is Nothing And oDB.SQLbAvailableServer(oDB.OffServerName, oDB.OnComIP) Then
            If orsMem.State = adStateOpen Then
                If Not (orsMem.BOF And orsMem.EOF) Then
                         If Format(SP_FEDdChkDate(orsMem, "FDMcdExpDate"), "YYYY/MM/DD") >= Format(dVB_CurrentDate, "YYYY/MM/DD") Then 'ทาง FSBT ต้องการให้ตรวจสอบเรื่องวันที่ Expire Date จากวันที่ System Date
                             tSta = SP_FEDtChkString(orsMem, "FTMcdSta")
                            If UCase(tSta) = "A" Or UCase(tSta) = "1" Then
'                                cLastPoint = SP_FEDvChkNumeric(orsMem, "FCMcdTotalPoint")
                                cLastPoint = SP_FEDvChkNumeric(orsMem, "FCBalance")           '*Ao 2013/02/12    CommSheet TK-ISS3000-017 00.00.03 Replied by JESLIN_Aoo.xls
                                tMemberName = SP_FEDtChkString(orsMem, "FTMcdName") & " " & SP_FEDtChkString(orsMem, "FTMcdSurName")
                                bProcess = True
                            Else
                                bW_FoundMallCard = False
                                Select Case UCase(tSta)
                                    Case "E", "2"
                                        cLastPoint = 0
                                        tMemberName = ""
                                    Case "L", "3"
                                        cLastPoint = 0
                                        tMemberName = ""
                                    Case "H", "4"
                                        cLastPoint = 0
                                        tMemberName = ""
                                End Select
                            End If
                        Else
                            cLastPoint = 0
                            tMemberName = ""
                        End If
                Else
                    cLastPoint = 0
                    tMemberName = "New member"
                    bProcess = True
                End If
            Else
                cLastPoint = 0
                tMemberName = ""
            End If
        Else
            cLastPoint = 0
            tMemberName = ""
        End If
    Else
        cLastPoint = 0
        tMemberName = ""
    End If
    
    '*KT 52-02-24 member point process คำนวณ point
    '---------------------------------------------------------------------------------------------
    cPoint = 0
    cPointAmt = 0
    '*KT 52-11-16 Communication Sheet PH3-ML-SF-010 If Mall Card set to Not Allowed Mall Card.xls
    Select Case nVB_RegMallCount
        Case 1
            ReDim aVB_MallCardReg(0)
            aVB_MallCardReg(0).ID = tVB_DiscMallCardID ' IIf(tVB_StaMemPointAlw = "Y", tVB_DiscMallCardID, "")
            aVB_MallCardReg(0).Type = tVB_DiscMallCardKey
            aVB_MallCardReg(0).Point = cPoint
'            aVB_MallCardReg(0).LPoint = IIf(tVB_StaMemPointAlw = "Y", cLastPoint, 0)
            aVB_MallCardReg(0).LPoint = cLastPoint           '*Ao 2013/02/12    CommSheet TK-ISS3000-017 00.00.03 Replied by JESLIN_Aoo.xls
            aVB_MallCardReg(0).Amt = cPointAmt
            aVB_MallCardReg(0).IDLastPoint = IIf(tVB_StaMemPointAlw = "Y", tVB_DiscMallCardID, "")
            aVB_MallCardReg(0).CardID = tVB_DiscMallCardID
            aVB_MallCardReg(0).tFTMcdCardType = tVB_DiscMallCardType '*Bmp 61-02-15
        Case 2, 3
            ReDim Preserve aVB_MallCardReg(UBound(aVB_MallCardReg) + 1)
            aVB_MallCardReg(nVB_RegMallCount - 1).ID = tVB_DiscMallCardID 'IIf(tVB_StaMemPointAlw = "Y", tVB_DiscMallCardID, "")
            aVB_MallCardReg(nVB_RegMallCount - 1).Type = tVB_DiscMallCardKey
            aVB_MallCardReg(nVB_RegMallCount - 1).Point = cPoint
'            aVB_MallCardReg(nVB_RegMallCount - 1).LPoint = IIf(tVB_StaMemPointAlw = "Y", cLastPoint, 0)
            aVB_MallCardReg(nVB_RegMallCount - 1).LPoint = cLastPoint     '*Ao 2013/02/12    CommSheet TK-ISS3000-017 00.00.03 Replied by JESLIN_Aoo.xls
            aVB_MallCardReg(nVB_RegMallCount - 1).Amt = cPointAmt
            aVB_MallCardReg(nVB_RegMallCount - 1).IDLastPoint = IIf(tVB_StaMemPointAlw = "Y", tVB_DiscMallCardID, "")
            aVB_MallCardReg(nVB_RegMallCount - 1).CardID = tVB_DiscMallCardID
            aVB_MallCardReg(nVB_RegMallCount - 1).tFTMcdCardType = tVB_DiscMallCardType '*Bmp 61-02-15
        Case Else
            Exit Sub
    End Select
    
    tSql = "UPDATE " & tVB_CNTblHD
    Select Case nVB_RegMallCount
        Case 1
            tSql = tSql & " SET FTCstRegCardCode ='" & aVB_MallCardReg(nVB_RegMallCount - 1).ID & "'"
        Case 2
            tSql = tSql & " SET FTCstRegCardCode2 ='" & aVB_MallCardReg(nVB_RegMallCount - 1).ID & "'"
        Case 3
            tSql = tSql & " SET FTCstRegCardCode3 ='" & aVB_MallCardReg(nVB_RegMallCount - 1).ID & "'"
    End Select
    
'*KT 52-12-29 แก้ไชเงื่อนไขการให้แต้มกับบัตรใบแรกที่มีคุณสมบัติให้แต้มใบแรกตาม DCR  PRJ09010TheMall-R3-DCR_12001(Member Point)-01.00.00.xls
    If tVB_StaMemPointAlw = "Y" And tVB_MemPointID = "" Then 'PH2-ML-SF-097
        tSql = tSql & " ,FTCstCardCode ='" & aVB_MallCardReg(nVB_RegMallCount - 1).IDLastPoint & "'"
        tSql = tSql & " ,FTCstCardType = '" & aVB_MallCardReg(nVB_RegMallCount - 1).Type & "'"
        tSql = tSql & " ,FNCstCardPoint = " & CLng(aVB_MallCardReg(nVB_RegMallCount - 1).Point)
        tSql = tSql & " ,FNCstCardLPoint = " & CLng(aVB_MallCardReg(nVB_RegMallCount - 1).LPoint)
        tSql = tSql & " ,FCCstCardAmt = " & aVB_MallCardReg(nVB_RegMallCount - 1).Amt
        tSql = tSql & " ,FTCstPointName = '" & SP_STRtSiggleCode(tMemberName) & "'" '*KT 53-02-10 Communication Sheet PH3-ML-SF-079 Can't read Mall card which customer name or surname included character'.xls
        tVB_MemPointID = aVB_MallCardReg(nVB_RegMallCount - 1).IDLastPoint
        nVB_CurMemPointCount = nVB_RegMallCount
        nVB_MemPointMnu = 2
        Call SP_CRTxLogMallCardEvent("MSR/Keyboard", "MemberPointRegister[1]", IIf(tVB_CNTrnDiscType <> "1850", "Discount", "Mall Card Register"), Trim(otbCusID.Text))
    End If
    tSql = tSql & "," & UT_tLastUpdSql
    tSql = tSql & " WHERE  FTTmnNum='" & tVB_CNTerminalNum & "' AND FTShdTransNo='" & tVB_Running & "'"
    tSql = tSql & " AND  FTShdTransType='" & tVB_Transaction & "' AND FDShdTransDate='" & tUT_SaleDate & "'"
    Call SP_SQLbExecute(tSql)
    Exit Sub
ErrHadle:
    Call SP_MSGxShowError(Err.Number, Err.Description)
    Call SP_CRTxLogMallCardEvent("MSR/Keyboard", "MemberPointRegister[2]", IIf(tVB_CNTrnDiscType <> "1850", "Discount", "Mall Card Register"), Trim(otbCusID.Text))
    Call SP_DATxRsNothing(orsTemp)
    Call SP_DATxRsNothing(orsTemp2)
    Call SP_DATxRsNothing(orsTemp3)
End Sub

Private Function W_CHKbMallCardGroup(ByVal ptMcdGrp As String) As Boolean
'----------------------------------------
'Ret : True : มีการใช้ Mall Card ที่เป็นกลุ่มเดียวกันในการลดราคาก่อนหน้านี้ Fale : ไม่มีการใช้ Mall Card ที่เป็นกลุ่มเดียวกันในการลดราคาก่อนหน้านี้
'Call: หาว่าใน CD มีการใช้ Mall Card ที่เป็นกลุ่มเดียวกันหรือไม่
'*KT 52-02-17 ปรับปรุงให้สามารถตรวจสอบเงื่อนไขการใช้ Mall Card 3 Group
''*KT 53-03-30 PH1.8-ML-SF-016 ไม่ว่ากรณีใดๆๆๆถ้าไม่มีการใช้งาน Group ที่ 1 ก่อนจะใช้เรียงลำดับต่อมาไม่ได้ไม่ว่า จะใช้รูด Mall Card ใบเดียว
'----------------------------------------
    Dim tSql$
    Dim orsTemp As ADODB.Recordset
    Dim nI As Long
    Dim bPass As Boolean
    Dim nMin As Long
    Dim nMax As Long
    Dim nTemp As Long
    Dim tCardType As String
On Error GoTo ErrHandle
    W_CHKbMallCardGroup = False
    
'    tSql = "SELECT " & tVB_CNTblCD & ".FTScdCardID,TCNMMallCardMtn.FTMcdGroup"
'    tSql = tSql & vbCrLf & "FROM TCNMMallCardMtn INNER JOIN"
'    tSql = tSql & vbCrLf & "" & tVB_CNTblCD & " ON TCNMMallCardMtn.FTMcdCardType = " & tVB_CNTblCD & ".FTScdCardType"
'    tSql = tSql & vbCrLf & "WHERE  (" & tVB_CNTblCD & ".FTTmnNum = '" & tVB_CNTerminalNum & "') AND (" & tVB_CNTblCD & ".FTShdTransNo = '" & tVB_Running & "') "
'    tSql = tSql & vbCrLf & "AND (" & tVB_CNTblCD & ".FDShdTransDate  ='" & tUT_SaleDate & "')"
'    tSql = tSql & vbCrLf & "GROUP  BY " & tVB_CNTblCD & ".FTScdCardID,TCNMMallCardMtn.FTMcdGroup"
'    tSql = tSql & vbCrLf & "ORDER BY TCNMMallCardMtn.FTMcdGroup ASC"
'    Call SP_SQLvExecute(tSql, orsTemp) 'read master
'    If Not (orsTemp.BOF Or orsTemp.EOF) Then
'         orsTemp.Filter = "FTMcdGroup = '" & ptMcdGrp & "'"
'         If (orsTemp.BOF And orsTemp.EOF) Then
'            orsTemp.Filter = adFilterNone
'            bPass = False
'            orsTemp.MoveFirst
'            nMin = CLng(Val(Format(SP_FEDtChkString(orsTemp, "FTMcdGroup"))))
'            orsTemp.MoveLast
'            nMax = CLng(Val(Format(SP_FEDtChkString(orsTemp, "FTMcdGroup"))))
'            nTemp = CLng(Val(Format(ptMcdGrp)))
'            If nTemp > nMin And nTemp > nMax Then
'                For nI = 1 To nTemp - 1
'                    orsTemp.Filter = "FTMcdGroup = '" & CStr(nI) & "'"
'                    If (orsTemp.BOF And orsTemp.EOF) Then
'                        bPass = True
'                        Exit For
'                    End If
'                    orsTemp.Filter = adFilterNone
'                Next nI
'            Else
'                bPass = True
'            End If
'            W_CHKbMallCardGroup = bPass
'        Else
'            W_CHKbMallCardGroup = True
'        End If
'    Else
'        If Val(Format(ptMcdGrp)) <> 1 Then
'            W_CHKbMallCardGroup = True
'        End If
'    End If
'    Call SP_DATxRsNothing(orsTemp)
'    Exit Function


    tSql = "SELECT " & tVB_CNTblCD & ".FTScdCardID," & tVB_CNTblCD & ".FTScdCardType"
    tSql = tSql & vbCrLf & "FROM " & tVB_CNTblCD
    tSql = tSql & vbCrLf & "WHERE  (" & tVB_CNTblCD & ".FTTmnNum = '" & tVB_CNTerminalNum & "') "
    tSql = tSql & vbCrLf & "AND (" & tVB_CNTblCD & ".FTShdTransNo = '" & tVB_Running & "') "
    tSql = tSql & vbCrLf & "AND (" & tVB_CNTblCD & ".FDShdTransDate  ='" & tUT_SaleDate & "')"
    tSql = tSql & vbCrLf & "AND (" & tVB_CNTblCD & ".FNDctNo  =" & nEN_Disc11MallCrd & ")"
    '*Eaw 56-07-03 CommSheet TK-ISS3000-263.xls
    Select Case tVB_CNSmnID
        Case "1809"
            tSql = tSql & vbCrLf & "AND (" & tVB_CNTblCD & ".FTScdDiscType  ='1' )"
        Case "1887"
            tSql = tSql & vbCrLf & "AND (" & tVB_CNTblCD & ".FTScdDiscType  = '2' )"
    End Select
    tSql = tSql & vbCrLf & "GROUP  BY " & tVB_CNTblCD & ".FTScdCardID," & tVB_CNTblCD & ".FTScdCardType"
    If SP_SQLvExecute(tSql, orsTemp) = 0 Then
        If Not (orsTemp.BOF Or orsTemp.EOF) Then
            tCardType = ""
            Do While Not orsTemp.EOF
                tCardType = tCardType & "'" & SP_FEDtChkString(orsTemp, "FTScdCardType") & "'"
                orsTemp.MoveNext
                If Not orsTemp.EOF Then tCardType = tCardType & ","
            Loop
            Call SP_DATxRsClose(orsTemp)
            If tCardType <> "" Then
                tSql = "SELECT TCNMMallCardMtn.FTMcdGroup"
                tSql = tSql & vbCrLf & "FROM TCNMMallCardMtn "
                tSql = tSql & vbCrLf & "WHERE TCNMMallCardMtn.FTMcdCardType IN (" & tCardType & ")"
                tSql = tSql & vbCrLf & "GROUP  BY TCNMMallCardMtn.FTMcdGroup"
                tSql = tSql & vbCrLf & "ORDER  BY TCNMMallCardMtn.FTMcdGroup ASC"
                If SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine, , , , , True, True) = 0 Then 'read master
                    If Not (orsTemp.BOF And orsTemp.EOF) Then
                          orsTemp.Filter = "FTMcdGroup = '" & ptMcdGrp & "'"
                         If (orsTemp.BOF And orsTemp.EOF) Then
                            orsTemp.Filter = adFilterNone
                            bPass = False
                            orsTemp.MoveFirst
                            nMin = CLng(Val(Format(SP_FEDtChkString(orsTemp, "FTMcdGroup"))))
                            orsTemp.MoveLast
                            nMax = CLng(Val(Format(SP_FEDtChkString(orsTemp, "FTMcdGroup"))))
                            nTemp = CLng(Val(Format(ptMcdGrp)))
                            If nTemp > nMin And nTemp > nMax Then
                                For nI = 1 To nTemp - 1
                                    orsTemp.Filter = "FTMcdGroup = '" & CStr(nI) & "'"
                                    If (orsTemp.BOF And orsTemp.EOF) Then
                                        bPass = True
                                        Exit For
                                    End If
                                    orsTemp.Filter = adFilterNone
                                Next nI
                            Else
                                bPass = True
                            End If
                            W_CHKbMallCardGroup = bPass
                        Else
                            W_CHKbMallCardGroup = True
                        End If
                    Else
                        If Val(Format(ptMcdGrp)) <> 1 Then
                               W_CHKbMallCardGroup = True
                        End If
                    End If
                Else
                    If Val(Format(ptMcdGrp)) <> 1 Then
                        W_CHKbMallCardGroup = True
                    End If
                End If
            Else
                If Val(Format(ptMcdGrp)) <> 1 Then
                    W_CHKbMallCardGroup = True
                End If
            End If
        Else
            If Val(Format(ptMcdGrp)) <> 1 Then
                W_CHKbMallCardGroup = True
            End If
        End If
    Else
        If Val(Format(ptMcdGrp)) <> 1 Then
            W_CHKbMallCardGroup = True
        End If
    End If
    Call SP_DATxRsNothing(orsTemp)
    Exit Function
ErrHandle:
    W_CHKbMallCardGroup = False
    Call SP_DATxRsNothing(orsTemp)
    Exit Function
End Function

'Private Sub Command1_Click()
'    Dim tTrack1Data$ 'Bump 56-04-02 TC Card
'    Dim tTrack2Data$ 'Bump 54-04-02 TC Card
'    Dim tCheckCardRang  As String   'Bump 54-04-02 TC Card
'    Dim tCardID  As String   'Bump 54-04-02 TC Card
'    Dim nPos As Long
'                    tTrack1Data = "s60000120345389706440548160000000000000040200000000000000000000000060" '.Track1Data '-->s6000012035334920094027419nnn n คือ ตัวเลขที่ตามหลังหมายเลขบัตร
'                    tTrack2Data = "4538970644054816=04021011387700000947" '.Track2Data   '-->5334920094027419=0402nnn
'
''                    Call SP_CRTxLogMallCardEvent("MSR/TCCard", "MSR", "Track1Data:", tTrack1Data & ":")
''                    Call SP_CRTxLogMallCardEvent("MSR/TCCard", "MSR", "Track2Data:", tTrack2Data & ":")
'                    If tTrack1Data <> "" And tTrack2Data <> "" Then
'                        nPos = InStrRev(tTrack2Data, "=")
'                        If nPos > 1 Then
'                            tCardID = Trim(Mid(tTrack2Data, 1, nPos - 1))  '-->"4538970644054816" คือ หมายเลขบนบัตร TC Card
'                            nPos = InStrRev(tTrack1Data, tCardID)
'                                If nPos > 1 Then
'                                    tCheckCardRang = Mid(tTrack1Data, (nPos - 10), 26)
'                                    If Len(tCheckCardRang) < 26 Then tCheckCardRang = tCardID
'                                Else
'                                    tCheckCardRang = tCardID
'                                End If
''                                Call SP_CRTxLogMallCardEvent("MSR/TCCard", "MSR", "CardID:", tCardID & ":")
''                                Call SP_CRTxLogMallCardEvent("MSR/TCCard", "MSR", "IDMark->Validaty Code:", tCheckCardRang)
'                        End If
'                    End If
'
''                    nPos = InStrRev(tTrack2Data, "=")
''                    If nPos > 1 Then
''                        tCardID = Trim(Mid(tTrack2Data, 1, nPos - 1))  'หา card ID จาก track2 5334920094027419=0402nnn n คือ ตัวเลขที่ตามหลัง =YYMM ของบัตร
''                    End If
''                    nPos = InStrRev(tTrack1Data, tCardID)
''                    If nPos > 1 Then
''                        tCheckCardRang = Mid(tTrack1Data, (nPos - 10), 26)
''                        If Len(tCheckCardRang) < 26 Then tCheckCardRang = tCardID
''                    Else
''                        tCheckCardRang = tCardID
''                    End If
'End Sub

'Private Sub Command1_Click()
'    Call OPOSMSR1_DataEvent(0)
'End Sub
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
    tSend = W_EDCtGetString2Send("BP")        'หายอดเงินที่จะส่งไป    ?? fix SALE จริงๆต้องแยกตาม type
    'tSend = "C2095706000001"
'    tSend = "C209" '*Bump 58-07-29 ส่งำปเพื่ออ่านหมายเลขบัตรจาก EDC
    If tSend = "" Then
        tMsg = "Sending string is empty"
        Call SP_EDCxINSERTLog(tMsg)
        '----------------------------------------------------------------------
        '*Bump 57-05-12 Phase2 เก็บไว้ตรวจสอบปัญหา
        Call SP_CRTxLogMallCardExtend(("Date:" & Date & " Time:" & Time$ & " C209:POSRequestError: " & tMsg))
        '----------------------------------------------------------------------
        Call SP_MSGnShowing("tms_CN005065" & ";" & tMsg, nCS_Warning)      '*Korn ?? เปลี่ยน msg
        
        bW_ClickEDC = False
        
        Unload Me
        Exit Sub
    End If
    
    nGet = oMagic3.SendCommand(ptComPort, tSend)
    tMsg = W_EDCtReturnCase(nGet)
    If UCase(tMsg) <> "PASS" Then
        'อาจจะได้ case error 1460 TimeOut
        Call SP_EDCxINSERTLog(tMsg)
        '----------------------------------------------------------------------
        '*Bump 57-05-12 Phase2 เก็บไว้ตรวจสอบปัญหา
        Call SP_CRTxLogMallCardExtend(("Date:" & Date & " Time:" & Time$ & " C209:POSRequestError: " & tMsg))
        '----------------------------------------------------------------------
        tMsg = tMsg & vbCrLf & "Do you want to try again?"
        If SP_MSGnShowing("tms_CN005065" & ";" & tMsg, nCS_OkCancel) = vbOK Then       'ถ้าไม่ผ่าน, ยืนยันทำซ้ำอีกครั้ง
            otmEDC.Enabled = True   'เปิดใหม่อีกครั้ง
            bW_ChkESC = True       '*Ao 2013/12/03 CommSheet TK-ISS3000-FAPL-161
            Exit Sub
        Else
            bW_ClickEDC = False
            Unload Me       'ปิดไปเลย
            Exit Sub
        End If
    End If
    'แสดงว่าส่ง echo ผ่านแล้ว, ส่งมูลค่าผ่านแล้ว, ให้ส่งรอ response
    tRet = ""   'reset ค่า
    nGet = oMagic3.RecvResponse(ptComPort, tRet, pnSecWait)
    tMsg = W_EDCtReturnCase(nGet)
    If UCase(tMsg) <> "PASS" Then
           'อาจจะได้ case error 1460 TimeOut
        Call SP_EDCxINSERTLog(tMsg)
        '---------------------------
         '*Bump 57-05-12 Phase2 เก็บไว้ตรวจสอบปัญหา
        Call SP_CRTxLogMallCardExtend(("Date:" & Date & " Time:" & Time$ & " C209:POSRequestError: " & tMsg))
        '---------------------------
        tMsg = tMsg & vbCrLf & "Do you want to try again?"
        If SP_MSGnShowing("tms_CN005065" & ";" & tMsg, nCS_OkCancel) = vbOK Then       'ถ้าไม่ผ่าน, ยืนยันทำซ้ำอีกครั้ง
            otmEDC.Enabled = True   'เปิดใหม่อีกครั้ง
            bW_ChkESC = True       '*Ao 2013/12/03 CommSheet TK-ISS3000-FAPL-161
            Exit Sub
        Else
            bW_ClickEDC = False
            Unload Me       'ปิดไปเลย
            Exit Sub
        End If
    Else        'ผ่าน
        Call SP_CRTxLogMallCardExtend(("Date:" & Date & " Time:" & Time$ & " C209:EDCResponse(PASS): " & tMsg))
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
                bW_ClickEDC = False
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
'            cVB_Pay = 0
'            tVB_CrdCardNo = ""
'            tVB_CrdApprCode = ""
             'ส่งค่าไป split
            tCut = Left(tRet, 4)
            'Call SP_CRTxLogMallCardEvent("C209", "EDCResponse", "ResponseValue=", tRet) 'Bump*56-04-26 เก็บ Log
            Call SP_CRTxLogMallCardExtend(("Date:" & Date & " Time:" & Time$ & " C209:EDCResponse: " & tMsg))
            'R209
            'R20902165124920000045038140420032201C3902009800
            If UCase(tCut) = UCase("R209") Then
                'Call W_EDCxSplitResponse(Trim(tRet)) '*Tao 62-12-10 ของเดิม
                Call W_EDCxSplitResponse(tRet) '*Tao 62-12-10 ไม่ Trim
                '-----------------------------------------------------------
                '*Bump 57-05-12 Phase2 เก็บไว้ตรวจสอบปัญหา
                '* PRJ10001TK-PH2-External Spec-Goods Return-00.02.06 FSBT.docx
                '* Page.8
                '* TK-ISS3000-463 TK-ISS3000-FAPL-182   EDC error message and response code
'                Call SP_EDCxSplitResponseExtend("C", Trim(tRet)) '*C = Credit card, "E" = Staff discount/Employee discount
                Call SP_CRTxLogMallCardExtend(("Date:" & Date & " Time:" & Time$ & " tVB_Transaction:" & tVB_Transaction))  '*Bump 57-11-04 TK_Issue 27/10/2017 รับชำระด้วย Credit ทำให้ Transaction Type เปลี่ยนจาก 03 เป็น 31
                '-----------------------------------------------------------
            Else
'                If UCase(tCut) = UCase("R922") Then
'                    Call SP_EDCxSplitResponseExtend("CE", Trim(tRet)) '*Bump 57-07-07 *C = Credit card, "E" = Staff discount/Employee discount, "CE" คือ เอา staff discount cad มารูดที่ credit
'                    Call SP_CRTxLogMallCardExtend(("Date:" & Date & " Time:" & Time$ & " tVB_Transaction:" & tVB_Transaction))  '*Bump 57-11-04 TK_Issue 27/10/2017 รับชำระด้วย Credit ทำให้ Transaction Type เปลี่ยนจาก 03 เป็น 31
'                End If
                Call SP_MSGnShowing("tms_CN005065" & ";" & "Invalid Card Type", nCS_Warning)
                bW_ClickEDC = False
                Unload Me       'ปิดไปเลย
            End If
'            Call W_EDCxSplitResponse(Trim(tRet))
        End If
    End If
    Exit Sub
ErrHandle:
'    cVB_Pay = 0
'    tVB_CrdCardNo = ""
'    tVB_CrdApprCode = ""
    Call SP_MSGxShowError(Err.Number, Err.Description)
    bW_ClickEDC = False
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
            otcForm1(0).Enabled = False     'lock text box
            otcForm1(2).Enabled = False     'lock ปุ่ม
            otmForm.Enabled = False  'stop otmForm
            otmEDC.Enabled = True     'start otmEDC
            
        ElseIf nW_Timer < nW_TimeESC Then
            olaForm(8).Caption = nW_TimeESC - nW_Timer
            olaForm(9).Caption = nW_TimeESC - nW_Timer
        End If
    End If

End Sub

Private Sub otmTime_Timer()
'    Dim nInt As Integer
'    nInt = 0
'    Do While nInt <> 1000
'        nInt = nInt + 1
'        If nInt = 2 Then
'            otmTime.Enabled = False
'        End If
'    Loop
End Sub


