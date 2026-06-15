VERSION 5.00
Object = "{D6810AA0-45D7-11CF-88F3-02608CA3BE43}#1.0#0"; "Msr.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.1#0"; "Codejock.SuiteCtrls.Unicode.v11.1.3.ocx"
Object = "{9D6F2CC0-07CE-11CF-A800-0020AF9EBB73}#1.0#0"; "Scanner.ocx"
Begin VB.Form wCNDiscEmp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Employee Card;klg_Title"
   ClientHeight    =   3825
   ClientLeft      =   2460
   ClientTop       =   4230
   ClientWidth     =   7110
   ControlBox      =   0   'False
   Icon            =   "wCNDiscEmp.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   7110
   Begin VB.Frame ofaSwipeCard 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   2535
      Left            =   3480
      TabIndex        =   18
      Top             =   5760
      Width           =   3495
      Begin VB.Label olaForm 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Swipe or Insert Card into EDC"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   11
         Left            =   480
         TabIndex        =   19
         Top             =   960
         Width           =   2535
      End
   End
   Begin VB.Frame ofa3Sec 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   2535
      Left            =   120
      TabIndex        =   15
      Top             =   5760
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
         Index           =   8
         Left            =   1320
         TabIndex        =   17
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label olaForm 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "  Insert Card at EDC    "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   10
         Left            =   480
         TabIndex        =   16
         Top             =   960
         Width           =   2655
      End
   End
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   3360
      Index           =   0
      Left            =   0
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   -360
      Width           =   7140
      _Version        =   720897
      _ExtentX        =   12594
      _ExtentY        =   5927
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
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
         Top             =   1200
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
         Left            =   2520
         MaxLength       =   2
         TabIndex        =   1
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox otbEmpCard 
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
         Top             =   600
         Width           =   4215
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
         TabIndex        =   14
         Tag             =   "2;"
         Top             =   1200
         Width           =   2775
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
         TabIndex        =   12
         Top             =   1200
         Width           =   210
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
         TabIndex        =   11
         Tag             =   "2;"
         Top             =   1200
         Width           =   2130
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
         TabIndex        =   10
         Top             =   1800
         Width           =   4215
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Employee No.:;klg_EmpID"
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
         TabIndex        =   9
         Tag             =   "2;"
         Top             =   600
         Width           =   2130
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
         Top             =   2400
         Width           =   4215
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Type:;klg_EmpType"
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
         TabIndex        =   7
         Tag             =   "2;"
         Top             =   1800
         Width           =   2145
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
         TabIndex        =   6
         Tag             =   "2;"
         Top             =   2400
         Width           =   2235
      End
   End
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   1110
      Index           =   1
      Left            =   0
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   3000
      Width           =   7140
      _Version        =   720897
      _ExtentX        =   12594
      _ExtentY        =   1958
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
      PaintManager.Position=   2
      Begin MSRLib.MSR OPOSMSR1 
         Left            =   120
         Top             =   120
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
      Begin VB.Timer otmEDC 
         Enabled         =   0   'False
         Interval        =   300
         Left            =   1680
         Top             =   120
      End
      Begin VB.Timer otmForm 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   1200
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
         Height          =   510
         Left            =   5280
         TabIndex        =   4
         Tag             =   "2;"
         Top             =   150
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
         Height          =   510
         Left            =   3600
         TabIndex        =   3
         Tag             =   "2;"
         Top             =   150
         Width           =   1455
      End
   End
End
Attribute VB_Name = "wCNDiscEmp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private bW_FoundEmpCard As Boolean
Private tW_ChkExp  As String
Private tW_ChkDpt  As String
Private tW_ChkLimit  As String
Private tW_ChkBay  As String
Private tW_EcdFrom As String
Private tW_EcdTo As String
Private cW_MaxDis As Double
Private orsMemCard As ADODB.Recordset
Private nW_CaseErr As Integer
Private bW_KeyEnter As Boolean '***KT 51-12-01 ป้องกันการเคาะ Enter เบิ้ลทำใหเเกิดส่วนลดเบิ้ล
Private tW_CardName As String   '*KT 52-11-26 PH2-ML-SF-170 ใน Ticket ไม่พิมพ์ชื่อพนักงานที่ใช้ส่วนลด Employee discount

Private bW_ManaulMode As Boolean
Private bW_HandleESC As Boolean
Private nW_TimeESC As Long
Private nW_Timer As Long
Private tW_Comport As String, tW_EDCRun As String, tW_PURCHASE As String, tW_EmpRemark As String
Private bW_ChkMEM  As Boolean
Private tW_ChkAlwPoint As String '*BG 56-08-01 CommSheet TK-ISS3000-196-ST 00 00 02.xls
Private bW_ChkESC As Boolean

Private Function W_PRCbDisCountEMP(ByVal ptRet As String, ByRef pcEMP As Double) As Boolean
'----------------------------------------------------------------------------
'Call:
'Cmt:   '*Ao 2013/10/29
'----------------------------------------------------------------------------
Dim tTemp As String, tName As String, tPURCHASE As String, tCut As String, tSql As String
Dim aTag() As String
Dim n As Integer, nLen As Integer
Dim bChk As Boolean
Dim orsTemp As ADODB.Recordset
Dim tTmpCard As String
Dim tTmpDate As String
Dim tExpired As String
 
On Error GoTo ErrorHandles
    
    bChk = False
    tTemp = ptRet
    
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
        tTemp = Right(tTemp, Len(tTemp) - 2)    'ตัด temp ออก
        aTag(0, n) = tCut
        'เก็บค่า value ลง aTag(1, n)
        tCut = Left(tTemp, 2)   'tag length 2 ตัว(fixed)
        tTemp = Right(tTemp, Len(tTemp) - 2)   'ตัด temp ออก
        nLen = Int(tCut)        'ความยาวของ value
        tCut = Left(tTemp, nLen)    'tag value nLen ตัว
        tTemp = Right(tTemp, Len(tTemp) - nLen)   'ตัด temp ออก
        aTag(1, n) = tCut
    Loop
    
     '*TU 2020-09-09 CR_Encrypted Creditcard
    '---------------------------------------------------------------------------------------------------------
    tTmpCard = ""
    If Left(SP_EDCxCheckTag(aTag, n, "02"), 1) = "#" Then
       tTmpCard = SP_EDCtDecrypt(Trim(SP_EDCxCheckTag(aTag, n, "02")))
        If tTmpCard <> "" Then
            otbEmpCard.Text = Right(tTmpCard, Len(tTmpCard) - 2)
        Else
            otbEmpCard.Text = ""
        End If
    Else
        otbEmpCard.Text = SP_EDCxCheckTag(aTag, n, "02")
    End If
    tTmpDate = ""
    If Left(SP_EDCxCheckTag(aTag, n, "14"), 1) = "#" Then
        tTmpDate = SP_EDCtDecrypt(Trim(SP_EDCxCheckTag(aTag, n, "14")))
        tExpired = Right(tTmpDate, Len(tTmpDate) - 2)
    Else
        tExpired = SP_EDCxCheckTag(aTag, n, "14")
    End If
'    otbEmpCard.text = SP_EDCxCheckTag(aTag, n, "02")
'    orbYear.text = Left(SP_EDCxCheckTag(aTag, n, "14"), 2)
'    orbMonth.text = Right(SP_EDCxCheckTag(aTag, n, "14"), 2)
    orbYear.Text = Left(tExpired, 2)
    orbMonth.Text = Right(tExpired, 2)
    '---------------------------------------------------------------------------------------------------------
    tVB_DiscEmpCardID = Trim(otbEmpCard.Text)
    tName = SP_EDCxCheckTag(aTag, n, "98")

    
    Select Case Val(Format(tVB_TransactionSub))
        Case tEN_TrnDepositCancel                                   '*Ao 2014/04/29    PRJ10001TK-PH2-External Spec-EDC Layaway Cancel-00.02.04 FSBT.docx
            If W_PRCbFindEmpCardByEDC(tVB_DiscEmpCardID, tName, False) Then
                tSql = "SELECT SUM(FCScdAmtB4) AS FCScdAmtB4 FROM " & tVB_CNTblCD
                tSql = tSql & vbCrLf & "WHERE FTTmnNum = '" & tVB_CNTerminalNum & "'"
                tSql = tSql & vbCrLf & "AND FTShdTransNo = '" & tVB_Running & "'"
                tSql = tSql & vbCrLf & "AND FDShdTransDate = '" & tUT_SaleDate & "'"
                tSql = tSql & vbCrLf & "AND FNDctNo = 12"
                Call SP_SQLbExecute(tSql, orsTemp)
                If Not (orsTemp.BOF And orsTemp.EOF) Then
                    pcEMP = SP_FEDvChkNumeric(orsTemp, "FCScdAmtB4")
                    bChk = True
                Else
                    pcEMP = 0
                    bChk = True
                End If
                bW_FoundEmpCard = True
            End If
        Case Else
            If W_PRCbFindEmpCardByEDC(tVB_DiscEmpCardID, tName, False) Then
                oDisc.tEmplimit = ""
                oDisc.tEmpRemark = tW_EmpRemark
                oDisc.bRecalByItem = False
                oDisc.DISxClearAll
                oDisc.dTranDate = tUT_SaleDate
                oDisc.tTerminalNo = tVB_CNTerminalNum
                oDisc.tTransNo = tVB_Running
                oDisc.tCardID = tVB_DiscEmpCardID
                oDisc.tCardType = tVB_DiscEmpCardType
                oDisc.tCardKey = tVB_DiscEmpCardKey
                oDisc.cDiscValue = SP_VALcText2Double(olaRate.Caption)
                dVB_DiscEmpCardExp = DateSerial(2000 + Val(Format(orbYear.Text)), Val(Format(orbMonth.Text)), 1)
                oDisc.dCardExpDate = dVB_DiscEmpCardExp
                oDisc.tStaChkExpire = tW_ChkExp
                oDisc.tStaChkDepart = tW_ChkDpt
                oDisc.cAlwDis = cVB_DiscEmpRanking
                oDisc.tRefCode = tW_ChkBay & "," & tW_ChkLimit & "," & tW_ChkAlwPoint
                oDisc.tCusID = tW_EcdFrom & "," & tW_EcdTo
                oDisc.tCardName = tName
                oDisc.bChkEMP = True               'ใช้ INSERT EMP              ไม่ INSERT ลง CD
                pcEMP = oDisc.DIScDiscount(nEN_Disc12EmpCrd, eEN_VerifyDiscType.nEN_Type2ByDcn)
                pcEMP = oDisc.cDisEMP
                bChk = True
            Else
                tVB_DiscEmpCardID = ""
                tW_PURCHASE = ""
                pcEMP = 0
                bChk = False
            End If
    End Select
    W_PRCbDisCountEMP = bChk
    Exit Function
ErrorHandles:
    tVB_DiscEmpCardID = ""
    tW_PURCHASE = ""
    pcEMP = 0
    W_PRCbDisCountEMP = False
End Function

Private Sub W_PRCxCase(ByVal ptCase As String)
'--------------------------------------
'Call:
'Cmt:
'--------------------------------------
    Select Case ptCase
        Case "2013_1200"
            ofa3Sec.Left = 2040     'center 120
            ofa3Sec.Top = 360                   '5760
            ofaSwipeCard.Left = 3480    'ซ่อน
            ofaSwipeCard.Top = 5760
            otcForm1(0).Enabled = False     'lock text box
            otcForm1(1).Enabled = False     'lock ปุ่ม
            otmForm.Enabled = True                  'Start   01 02 03
            bW_HandleESC = True
            bW_ManaulMode = False
        Case "2013_1203"
            otbEmpCard.Locked = False
            otbEmpCard.BackColor = nVB_CNClrBUnLock
            bW_HandleESC = False
            bW_ManaulMode = True
            otmEDC.Enabled = False
            otmForm.Enabled = False
        Case "2013_1205"
            bW_HandleESC = False
            otmForm.Enabled = False
            otmEDC.Enabled = False
            ofa3Sec.Left = 120      'ซ่อน
            ofa3Sec.Top = 5760
            ofaSwipeCard.Left = 3480        'ซ่อน
            ofaSwipeCard.Top = 5760
            otcForm1(0).Enabled = True     'unlock text box
            otcForm1(1).Enabled = True     'unlock ปุ่ม
            bW_ManaulMode = True
            ocmCancel.Cancel = True
'            otbEmpCard.Locked = False
'            otbEmpCard.BackColor = nVB_CNClrBUnLock
            If bVB_ReEntryMode Then      '*Eaw 56-06-06  CommSheet TK-ISS3000-225-ST 00 00 01.xls แยกเคส Reentry Mode กับ Normal Mode
                otbEmpCard.Locked = False
                otbEmpCard.BackColor = nVB_CNClrBUnLock
            Else
                otbEmpCard.Locked = True
                otbEmpCard.BackColor = nVB_CNClrBLock
            End If
        Case "2013_1251"
            bW_HandleESC = False
            ofa3Sec.Left = 120      'ซ่อน
            ofa3Sec.Top = 5760
            ofaSwipeCard.Left = 2040   'center
            ofaSwipeCard.Top = 360
            otcForm1(0).Enabled = False     'lock text box
            otcForm1(1).Enabled = False     'lock ปุ่ม
            otmForm.Enabled = False  'stop otmForm
            otmEDC.Enabled = True     'start otmEDC
    End Select
    
End Sub

Private Function W_PRCbFindEmpCardByEDC(ByVal ptCusID As String, ByVal ptCardName As String, Optional ByVal pbChk As Boolean = False) As Boolean
'--------------------------------------
'Call:
'Cmt: ค้นหา Format ของเลขที่บัตร Mall Card อยู่ในช่วงใด     '*Ao 2013/02/05
'--------------------------------------
Dim orsTemp As ADODB.Recordset
Dim tSql As String
On Error GoTo ErrorHandles

    tSql = "SELECT TOP 1 * FROM  TCNMEmpCardMtn"
    tSql = tSql & vbCrLf & "WHERE '" & ptCusID & "' BETWEEN FTEcdFrom AND FTEcdTo"
    tSql = tSql & vbCrLf & "AND  LEN(" & ptCusID & ") <= LEN(FTEcdTo)"
    tSql = tSql & vbCrLf & "AND FTEcdStaRang = 'Y'"
    tSql = tSql & vbCrLf & "ORDER BY LEN(FTEcdFrom)"
    Call SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine, , , , , True, True) 'read master
    If Not (orsTemp.BOF Or orsTemp.EOF) Then
        olaName.Caption = SP_FEDtChkString(orsTemp, "FTEcdDesc1")
        cVB_DiscEmpRate = SP_FEDvChkNumeric(orsTemp, "FCEcdRate")
        olaRate.Caption = cVB_DiscEmpRate
        tVB_DiscEmpCardType = SP_FEDtChkString(orsTemp, "FTEcdCardType")
        tVB_DiscEmpCardKey = SP_FEDtChkString(orsTemp, "FTEcdCode")
        If pbChk Then
            tW_ChkExp = SP_FEDtChkString(orsTemp, "FTEcdStaChkExp")
            tW_EmpRemark = ""
            tW_PURCHASE = ""
        Else
            tW_ChkExp = "N"
            tW_EmpRemark = "EDC"
        End If
        tW_ChkDpt = SP_FEDtChkString(orsTemp, "FTEcdStaDeptCode")
        tW_ChkLimit = SP_FEDtChkString(orsTemp, "FTEcdStaCreditLimit")
        tVB_DiscEmpChkLimit = tW_ChkLimit '*KT 51-09-19 ML-SF- 558 Print  Remain Discount for Employee Card
        tW_ChkBay = SP_FEDtChkString(orsTemp, "FTEcdStaChkAlwBuy")
        tW_EcdFrom = SP_FEDtChkString(orsTemp, "FTEcdFrom")
        tW_EcdTo = SP_FEDtChkString(orsTemp, "FTEcdTo")
        cVB_DiscEmpRanking = SP_FEDvChkNumeric(orsTemp, "FCEcdAmtRank")
        tW_ChkAlwPoint = SP_FEDtChkString(orsTemp, "FTEcdStaAllowPoint")    '*BG 56-08-01 CommSheet TK-ISS3000-196-ST 00 00 02.xls
        tW_CardName = ptCardName
        W_PRCbFindEmpCardByEDC = True
        bW_FoundEmpCard = True
    Else
        otbEmpCard.Text = ""
        orbMonth.Text = ""
        orbYear.Text = ""
        olaName.Caption = ""
        olaRate.Caption = ""
        W_PRCbFindEmpCardByEDC = False
        bW_FoundEmpCard = False
    End If
    bW_ChkMEM = False
    Call SP_DATxRsNothing(orsTemp)
    Exit Function
ErrorHandles:
    Call SP_MSGxShowError(Err.Number, Err.Description)
    W_PRCbFindEmpCardByEDC = False
    bW_FoundEmpCard = False
End Function

Private Function W_EDCtGetString2Send(ByVal ptType As String, ByVal pcEMP As Double) As String
'-----------------------------------------------------------
'   Cmt : get message to send
'   Create by :
'-----------------------------------------------------------
Dim tSend$, tAmt$
Dim aPayAmt() As String
Dim tDec As String

    tSend = ""
'    tAmt = "000000000001"
    '------------------------------------------------------------------------------------------------------------------
    tAmt = Trim(Str(pcEMP))   'เปลี่ยนยอดเงินให้เป็น string                 '*Ao 2013/10/29
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
    '------------------------------------------------------------------------------------------------------------------
    tW_EDCRun = SP_EDCtRunning("", True)
    'แยกตามประเภทการส่ง
    'C200 = Sale, C201=PreAuth, C203=Refund, C300=Void
    Select Case UCase(ptType)
        Case "EMP"
            Select Case Val(Format(tVB_TransactionSub))
                Case tEN_TrnReturn, tEN_TrnReturnTaxInc, tEN_TrnReturnManualTax, tEN_TrnDepositCancel   '*Ao 2014/04/29    PRJ10001TK-PH2-External Spec-EDC Layaway Cancel-00.02.04 FSBT.docx
                    tSend = "C203"
                Case Else
                    tSend = "C200"
            End Select
'            tSend = "C200"                      'command Sale Employee Card
            tSend = tSend & "0412"       'tag 04 ยอดเงิน จำนวน 12 หลัก ,    fix
            tSend = tSend & tAmt            'ยอดเงิน รวมทศนิยม
            'tSend = tSend & "5601Y"    'tag 56 ตาราง tag 1.6.7 LIST OF HOST TYPES จำนวน 01 หลัก, fix Y=Taka Host '*Tao 62-09-25 Employee Card ให้เพิ่ม Tag26 ป้องกันเอาบัตรเครดิตมาเสียบ (จะตัดเงินจากบัตร)
            tSend = tSend & "5706"        'tag 57 ตาราง tag 1.6.4 จำนวน 06 หลัก, fix
            tSend = tSend & tW_EDCRun  'sale fix 000001    ??ไม่แน่ใจว่าจะต้องเป็น running หรือไม่
        Case "EMP2"
            tSend = "C922"                      'command Refund Employee Card
            tSend = tSend & "0412"       'tag 04 ยอดเงิน จำนวน 12 หลัก ,    fix
            tSend = tSend & tAmt            'ยอดเงิน รวมทศนิยม
            'tSend = tSend & "5601Y"    'tag 56 ตาราง tag 1.6.7 LIST OF HOST TYPES จำนวน 01 หลัก, fix Y=Taka Host '*Tao 62-09-25 Employee Card ให้เพิ่ม Tag26 ป้องกันเอาบัตรเครดิตมาเสียบ (จะตัดเงินจากบัตร)
            tSend = tSend & "5706"        'tag 57 ตาราง tag 1.6.4 จำนวน 06 หลัก, fix
            tSend = tSend & tW_EDCRun  'sale fix 000001    ??ไม่แน่ใจว่าจะต้องเป็น running หรือไม่
    End Select
    
    Call SP_CRTxLogMallCardEvent("EmployeeDiscount", "EDCSendString", "SendStringValue=", tSend)
    
    W_EDCtGetString2Send = tSend
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
    Dim tTag39  As String
    Dim tTag55  As String   '*Bump 56-12-04 V.4.0.109
    Dim tSql As String
    Dim tTmpCard As String
    Dim tTmpDate As String
    Dim tExpired As String
    On Error GoTo ErrorHandles
    
    tTemp = ptRet
    
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
        tTemp = Right(tTemp, Len(tTemp) - 2)    'ตัด temp ออก
        aTag(0, n) = tCut
        
        'เก็บค่า value ลง aTag(1, n)
        tCut = Left(tTemp, 2)   'tag length 2 ตัว(fixed)
        tTemp = Right(tTemp, Len(tTemp) - 2)   'ตัด temp ออก
        nLen = Int(tCut)        'ความยาวของ value
        tCut = Left(tTemp, nLen)    'tag value nLen ตัว
        tTemp = Right(tTemp, Len(tTemp) - nLen)   'ตัด temp ออก
        aTag(1, n) = tCut
        
'        Debug.Print "aTag[0/1," & n & "]=>" & aTag(0, n) & "-" & aTag(1, n)
    Loop
    
    'เช็คว่ามีค่ารหัสบัตรหรือไม่ (02:มีรหัสบัตร)
'    If aTag(0, 1) <> "02" Then      'ถ้าไม่มีรหัสบัตร แสดงว่าอาจมีปัญหา
'        If aTag(0, 2) = "39" Then    'error แบบ track ได้   ส่วนใหญ่รหัส 39
'            Select Case aTag(1, 2)
'                Case "XX": tMsg = "Terminal not properly setup."
'                Case "IC": tMsg = "Invalid Card"
'                Case "TA": tMsg = "Transaction aborted by user."
'                Case Else
'                    tMsg = "Invalid Card"
'            End Select
'        Else
'            'all error, แบบ track ไม่ได้
'            tMsg = "Invalid Account Number !!"
'        End If
'        Call SP_EDCxINSERTLog(tMsg)
'        Call SP_MSGnShowing("tms_CN005065" & ";" & tMsg, nCS_Error)
'        Unload Me
'        Exit Sub
'    End If
    '**************************************************************
    '*Bump 56-12-04 เพิ่มการตรวจสอบ Tag55
    'Tag55 = "F" คือ Staff Siscount     ถ้าเป็นค่าอื่น เช่น V, M, A, J, D, C, E, P, R เป็น Credit card
'    tTag55 = SP_EDCxCheckTag(aTag, n, "55")
'    If UCase(tTag55) <> UCase("F") Then
'        Call SP_MSGnShowing("tms_CN005253", nCS_Warning) '*tms_CN005253 = Invalid Card.
'        Unload Me
'        Exit Sub
'    End If
    '**************************************************************
    tTag39 = SP_EDCxCheckTag(aTag, n, "39")
    If tTag39 = "" Then
        tMsg = SP_GETtMsgErrEDC(tTag39)
'        MsgBox tMsg, vbOKOnly, "EDC ERROR"
'        Unload Me
'        Exit Sub
        tMsg = tMsg & vbCrLf & "Do you want to try again?"
        If SP_MSGnShowing("tms_CN005065" & ";" & tMsg, nCS_OkCancel) = vbOK Then      '*Ao 2014/08/07 CommSheet TK-ISS3000-543-PH2  V 4.2.140
            otmEDC.Enabled = True   'เปิดใหม่อีกครั้ง
            Exit Sub
        Else
            Unload Me
            Exit Sub
        End If
    Else
        If tTag39 <> "00" Then
            tMsg = SP_GETtMsgErrEDC(tTag39)
'            MsgBox tMsg, vbOKOnly, "EDC ERROR"
'            Unload Me
'            Exit Sub
            tMsg = tMsg & vbCrLf & "Do you want to try again?"
            If SP_MSGnShowing("tms_CN005065" & ";" & tMsg, nCS_OkCancel) = vbOK Then       '*Ao 2014/08/07 CommSheet TK-ISS3000-543-PH2  V 4.2.140
                otmEDC.Enabled = True   'เปิดใหม่อีกครั้ง
                Exit Sub
            Else
                Unload Me
                Exit Sub
            End If
        End If
    End If
    'แสดงว่ารูดบัตร ผ่าน
    'ซ่อน Swipe card แต่ไม่ปลด lock ปุ่ม ok, cancel
    ofaSwipeCard.Left = 3480        'ซ่อน
    ofaSwipeCard.Top = 5760
    
    'นำค่าลงในหน้าจอ
         '*TU 2020-09-09 CR_Encrypted Creditcard
    '---------------------------------------------------------------------------------------------------------
    tTmpCard = ""
    If Left(SP_EDCxCheckTag(aTag, n, "02"), 1) = "#" Then
       tTmpCard = SP_EDCtDecrypt(Trim(SP_EDCxCheckTag(aTag, n, "02")))
        If tTmpCard <> "" Then
            otbEmpCard.Text = Right(tTmpCard, Len(tTmpCard) - 2)
        Else
            otbEmpCard.Text = ""
        End If
    Else
        otbEmpCard.Text = SP_EDCxCheckTag(aTag, n, "02")
    End If
    tTmpDate = ""
    If Left(SP_EDCxCheckTag(aTag, n, "14"), 1) = "#" Then
        tTmpDate = SP_EDCtDecrypt(Trim(SP_EDCxCheckTag(aTag, n, "14")))
        tExpired = Right(tTmpDate, Len(tTmpDate) - 2)
    Else
        tExpired = SP_EDCxCheckTag(aTag, n, "14")
    End If
'    otbEmpCard.text = SP_EDCxCheckTag(aTag, n, "02")
'    orbYear.text = Left(SP_EDCxCheckTag(aTag, n, "14"), 2)
'    orbMonth.text = Right(SP_EDCxCheckTag(aTag, n, "14"), 2)
    orbYear.Text = Left(tExpired, 2)
    orbMonth.Text = Right(tExpired, 2)
    '---------------------------------------------------------------------------------------------------------
    
    tVB_DiscEmpCardID = Trim(otbEmpCard.Text)
    tW_PURCHASE = SP_EDCxCheckTag(aTag, n, "63")
'    tName = SP_EDCxCheckTag(aTag, n, "98")
'    Data type(2 bytes) + PURCHASE LIMIT BEFORE DEDUCTION(12 bytes) + PURCHASE LIMIT AFTER DEDUCTION(12 bytes) +
'    tW_PURCHASE = SP_EDCxCheckTag(aTag, n, "63")
'    If W_PRCbFindEmpCardByEDC(tVB_DiscEmpCardID, tName, False) Then
'        tW_EDCRun = SP_EDCtRunning(tW_EDCRun, False)
'        Call ocmOK_Click            'สมบูรณ์แล้ว จะ unload โดยอัติโนมัติ
'    Else
'        Call W_PRCxCase("2013_1205")
'        Call SP_MSGnShowing("tms_CN005069", nCS_Warning)
'        Unload Me  '*Eaw 56-06-25 CommSheet TK-ISS3000-225-ST 00 00 05 FSBT NG.xlsx
'    End If
    '-------------------------------------------------------------------------------------------
    tW_EDCRun = SP_EDCtRunning(tW_EDCRun, False)
'    Call ocmOK_Click            'สมบูรณ์แล้ว จะ unload โดยอัติโนมัติ
    bVB_ChkEDC = True                '*Ao 2014/07/04
    
    Select Case Val(Format(tVB_TransactionSub))         '*Ao 2014/04/29    PRJ10001TK-PH2-External Spec-EDC Layaway Cancel-00.02.04 FSBT.docx
        Case tEN_TrnDepositCancel
            tSql = "UPDATE " & tVB_CNTblCD
            tSql = tSql & vbCrLf & "SET FTRemark = '" & tW_PURCHASE & "' "                      '*Ao 2014/08/14 CommSheet TK-ISS3000-545-PH2   V_4.2.142
            tSql = tSql & vbCrLf & "WHERE FTTmnNum = '" & tVB_CNTerminalNum & "' "
            tSql = tSql & vbCrLf & "AND FTShdTransNo = '" & tVB_Running & "' "
            tSql = tSql & vbCrLf & "AND FDShdTransDate = '" & tUT_SaleDate & "' "
            tSql = tSql & vbCrLf & "AND FNDctNo = 12"
            Call SP_SQLvExecute(tSql)
            SleepEx 2000, 0     '1000 = 1 วินาที
            Unload Me
        Case Else
            Call ocmOK_Click            'สมบูรณ์แล้ว จะ unload โดยอัติโนมัติ
    End Select

    Exit Sub
ErrorHandles:
      Call SP_MSGxShowError(Err.Number, Err.Description)
End Sub

Private Sub W_EDCxMagic3CreditCard(ByVal ptComPort$, ByVal pnSecWait As Long, ByVal ptType$)
'-----------------------------------------------------------
'   Cmt : EDC for Magic3
'              ** Bump ยืนยันว่าไม่มีการแบ่งจ่าย ถ้าเป็น EDC ให้รับยอดเต็มทั้งหมด    '*Ao แบ่งจ่ายได้
'   Create by : *Korn 55-10-08 Edit เพิ่ม procedure ใหม่
'-----------------------------------------------------------
Dim oMagic3 As New MagIC3Connector
Dim nGet As Long, tMsg$, tSend$, tRet$
Dim tCut  As String
Dim cEMP  As Double

On Error GoTo ErrHandle
    
    pnSecWait = 80000
    If pnSecWait = 0 Then pnSecWait = 100000     'ถ้าค่าเป็น  0 ให้ตั้งค่ารอที่ 40 วินาที
    
    '*ได้ทดสอบแล้วว่า ถ้าใช้ Dll จากของเขายังไงก็รอ ไม่ว่าจะเป็น demo(.Net ที่gเขาให้มา) ก็ตาม
    'ดังนั้น ก็ไม่จำเป็นจะต้อง loop หรือให้ Timer ทำงาน จึงต้องปิด otmEDC(T2) ก่อนเลยเพื่อให้มีการเรียกฟังก์ชั่นนี้แค่รอบเดียว, แต่ถ้าจะทำซ้ำค่อยเปิด enable ใหม่
    otmForm.Enabled = False  'เผื่อไว้ ปกติก็ปิดอยู่แล้ว T1
    otmEDC.Enabled = False  'ปิด T2 ก่อน
    
    'แสดงว่าส่ง echo ผ่านแล้ว, ให้ส่งมูลค่าไปต่อ
    Select Case Val(Format(tVB_TransactionSub))
        Case tEN_TrnDepositCancel                                   '*Ao 2014/04/29    PRJ10001TK-PH2-External Spec-EDC Layaway Cancel-00.02.04 FSBT.docx
            cEMP = wSub.olaTotal.Caption
        Case Else
            cEMP = wDisc.olaBal.Caption
    End Select
    
    tSend = W_EDCtGetString2Send(ptType, cEMP)       'หายอดเงินที่จะส่งไป    ?? fix SALE จริงๆต้องแยกตาม type
    If tSend = "" Then
        tMsg = "Sending string is empty"
        Call SP_EDCxINSERTLog(tMsg)
        '---------------------------
        '*Bump 57-05-12 Phase2 เก็บไว้ตรวจสอบปัญหา
        Call SP_CRTxLogMallCardExtend(("Date:" & Date & " Time:" & Time$ & " Error: " & tMsg))
        '---------------------------
        Call SP_MSGnShowing("tms_CN005065" & ";" & tMsg, nCS_Error)      '*Korn ?? เปลี่ยน msg
        Unload Me
        Exit Sub
    End If
    cEMP = 0
'*Eaw 56-07-09 CommSheet TK-ISS3000-268-TR.xls
    'check send Echo
'    nGet = oMagic3.EchoTest(ptComPort, tRet)
'    tMsg = SP_EDCtReturnCase(nGet)
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
    nGet = oMagic3.SendCommand(ptComPort, tSend)
    tMsg = SP_EDCtReturnCase(nGet)
    If UCase(tMsg) <> "PASS" Then
        'อาจจะได้ case error 1460 TimeOut
        Call SP_EDCxINSERTLog(tMsg)
        '---------------------------
        '*Bump 57-05-12 Phase2 เก็บไว้ตรวจสอบปัญหา
        Call SP_CRTxLogMallCardExtend(("Date:" & Date & " Time:" & Time$ & " Error: " & tMsg))
        '---------------------------
        tMsg = tMsg & vbCrLf & "Do you want to try again?"
        If SP_MSGnShowing("tms_CN005065" & ";" & tMsg, nCS_OkCancel) = vbOK Then       'ถ้าไม่ผ่าน, ยืนยันทำซ้ำอีกครั้ง
            otmEDC.Enabled = True   'เปิดใหม่อีกครั้ง
            Exit Sub
        Else
            Unload Me       'ปิดไปเลย
            Exit Sub
        End If
    End If

    'แสดงว่าส่ง echo ผ่านแล้ว, ส่งมูลค่าผ่านแล้ว, ให้ส่งรอ response
    tRet = ""   'reset ค่า
    nGet = oMagic3.RecvResponse(ptComPort, tRet, pnSecWait)
    tMsg = SP_EDCtReturnCase(nGet)
    If UCase(tMsg) <> "PASS" Then
        Call SP_EDCxINSERTLog(tMsg)
        '---------------------------
        '*Bump 57-05-12 Phase2 เก็บไว้ตรวจสอบปัญหา
        Call SP_CRTxLogMallCardExtend(("Date:" & Date & " Time:" & Time$ & " Error: " & tMsg))
        '---------------------------
        tMsg = tMsg & vbCrLf & "Do you want to try again?"
        If SP_MSGnShowing("tms_CN005065" & ";" & tMsg, nCS_OkCancel) = vbOK Then       'ถ้าไม่ผ่าน, ยืนยันทำซ้ำอีกครั้ง
            otmEDC.Enabled = True   'เปิดใหม่อีกครั้ง
            Exit Sub
        Else
            Unload Me
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
                Exit Sub
            Else
                Unload Me
                Exit Sub
            End If
        Else
            'R92202168888709999999904140415123902005415TAKA-STAFF-CARD5501F5601Y5706000001632240000000000000000000009819STAFF DISCOUNT CARD
            'tRet = "R92202168888700798005903140415123902005415TAKA-STAFF-CARD5501F5601Y5706000001632240000000000000000000009819STAFF DISCOUNT CARD"
            tCut = Left(tRet, 4)
            Call SP_CRTxLogMallCardEvent("EmployeeDiscount", "EDCResponse", "ResponseValue=", tRet)
            If UCase(tCut) = UCase("R922") Then
                '--------------------------------------------------------------------------------------------
                If W_PRCbDisCountEMP(tRet, cEMP) Then
                    tSend = W_EDCtGetString2Send("EMP2", cEMP)
                    nGet = oMagic3.SendCommand(ptComPort, tSend)
                    tMsg = SP_EDCtReturnCase(nGet)
                    If UCase(tMsg) <> "PASS" Then
                        '--------------------------------------------------------------------------------------------
                        '*Bump 57-05-12 Phase2 เก็บไว้ตรวจสอบปัญหา
                        Call SP_CRTxLogMallCardExtend(("Date:" & Date & " Time:" & Time$ & " Error: " & tMsg))
                        '--------------------------------------------------------------------------------------------
                        tMsg = tMsg & vbCrLf & "Do you want to try again?"
                        If SP_MSGnShowing("tms_CN005065" & ";" & tMsg, nCS_OkCancel) = vbOK Then
                            otmEDC.Enabled = True
                            Exit Sub
                        Else
                            Unload Me
                            Exit Sub
                        End If
                    End If
                    tRet = ""
                    nGet = oMagic3.RecvResponse(ptComPort, tRet, pnSecWait)
                    tMsg = SP_EDCtReturnCase(nGet)
                    If UCase(tMsg) <> "PASS" Then
                        '--------------------------------------------------------------------------------------------
                        '*Bump 57-05-12 Phase2 เก็บไว้ตรวจสอบปัญหา
                        Call SP_CRTxLogMallCardExtend(("Date:" & Date & " Time:" & Time$ & " Error: " & tMsg))
                        '--------------------------------------------------------------------------------------------
                        tMsg = tMsg & vbCrLf & "Do you want to try again?"
                        If SP_MSGnShowing("tms_CN005065" & ";" & tMsg, nCS_OkCancel) = vbOK Then
                            otmEDC.Enabled = True
                            Exit Sub
                        Else
                            Unload Me
                            Exit Sub
                        End If
                    End If
                    'R20002168888709999999904041200000000100007101029161405140415122201M3712DEMODEMODEMO3806DEMO51390200410877777777420410005204TAKA5415TAKA-STAFF-CARD5501F5601Y5706000001620600049963403031303030303030303030313030303030303030640600000196100000000004
                    'R20002168888700798005903041200000000010007101030141008140426122201M3712000000AA00013806AA0001390200410877777777420410005204TAKA5415TAKA-STAFF-CARD5501F5601Y5706000218620600051263402800000001309500000001309400000000000000640600000196100000000004
                    Call SP_CRTxLogMallCardEvent("EmployeeDiscount", "EDCResponse", "ResponseValue=", tRet)
                    '--------------------------------------------------------------------------------------------
                    '*Bump 57-05-12 Phase2 เก็บไว้ตรวจสอบปัญหา
                    '* PRJ10001TK-PH2-External Spec-Goods Return-00.02.06 FSBT.docx
                    '* Page.8
                    '* TK-ISS3000-463 TK-ISS3000-FAPL-182   EDC error message and response code
                    'Call SP_EDCxSplitResponseExtend("E", Trim(tRet)) '*C = Credit card, "E" = Staff discount/Employee discount '*Tao 62-12-10 ของเดิม
                    Call SP_EDCxSplitResponseExtend("E", tRet) '*C = Credit card, "E" = Staff discount/Employee discount '*Tao 62-12-10 ไม่ Trim
                    '--------------------------------------------------------------------------------------------
                    If Trim(tRet) = "" Then
                        tMsg = tMsg & vbCrLf & "Do you want to try again?"
                        If SP_MSGnShowing("tms_CN005065" & ";" & tMsg, nCS_OkCancel) = vbOK Then
                            otmEDC.Enabled = True
                            Exit Sub
                        Else
                            Unload Me
                            Exit Sub
                        End If
                    End If
                    tCut = Left(tRet, 4)
                    If UCase(tCut) = UCase("R200") Or UCase(tCut) = UCase("R203") Then
                        'Call W_EDCxSplitResponse(Trim(tRet)) '*Tao 62-12-10 ของเดิม
                        Call W_EDCxSplitResponse(tRet) '*Tao 62-12-10 ไม่ Trim
                    Else
'                        MsgBox "Invalid Response code", vbCritical + vbOKOnly, "EDC ERROR"
                        tMsg = "Invalid Response code"
                        tMsg = tMsg & vbCrLf & "Do you want to try again?"
                        If SP_MSGnShowing("tms_CN005065" & ";" & tMsg, nCS_OkCancel) = vbOK Then      '*Ao 2014/08/07 CommSheet TK-ISS3000-543-PH2  V 4.2.140
                            otmEDC.Enabled = True   'เปิดใหม่อีกครั้ง
                            Exit Sub
                        Else
                            Unload Me       'ปิดไปเลย
                            Exit Sub
                        End If
                     End If
                    '--------------------------------------------------------------------------------------------
'                    Call W_EDCxSplitResponse(Trim(tRet))
                Else
'                    Call W_PRCxCase("2013_1205")
'                    Call SP_MSGnShowing("tms_CN005069", nCS_Warning)
'                    Unload Me
                    tMsg = "Invalid Employee Card"                                      '*Ao 2014/0814   CommSheet TK-ISS3000-543-PH2   V 4.2.142
                    tMsg = tMsg & vbCrLf & "Do you want to try again?"
                    If SP_MSGnShowing("tms_CN005065" & ";" & tMsg, nCS_OkCancel) = vbOK Then
                        SleepEx 3000, 0                         '*Ao 2014/08/14   รอให้ EDC เริ่มต้นใหม่
                        otmEDC.Enabled = True   'เปิดใหม่อีกครั้ง
                        Exit Sub
                    Else
                        Unload Me
                        Exit Sub
                    End If
                End If
            Else
'                Call W_PRCxCase("2013_1205")                   '*Ao 2014/07/03
                '*Bump 56-12-04 CommSheet TK-ISS3000-441-FAPL-165.xls   V.109   561204
                '* อาจจะส่งมาเป็น แบบนี้  กรณีที่ กด CANCEL ที่ EDC หรือ ไม่รูดบัตรที่ EDC แต่ปล่อยให้หมดเวลา(TimeOut)
                '*R20004120000000200003902TA5706000113
                'tCut = Left(tRet, 4)
                '*Bump 57-07-07 กรณีเอาบัตรเครดิตมารูดบัตรที่ Staff Discount
                'Call SP_EDCxSplitResponseExtend("EE", Trim(tRet)) '*C = Credit card, "EE" = Invalid Staff discount/Employee discount '*Tao 62-12-10 ของเดิม
                Call SP_EDCxSplitResponseExtend("EE", tRet) '*C = Credit card, "EE" = Invalid Staff discount/Employee discount '*Tao 62-12-10 ไม่ Trim
                'Call SP_CRTxLogMallCardExtend(("Date:" & Date & " Time:" & Time$ & " Informs: This card is Invalid card for Staff Discount card."))
                '------------------------------
'                If UCase(tCut) = UCase("R200") Or UCase(tCut) = UCase("R203") Then
''                    Call W_EDCxSplitResponse(Trim(tRet))
'                    tMsg = SP_GETtMsgErrEDC("TA")
'                    tMsg = tMsg & vbCrLf & "Do you want to try again?"
'                    If SP_MSGnShowing("tms_CN005065" & ";" & tMsg, nCS_OkCancel) = vbOK Then       'ถ้าไม่ผ่าน, ยืนยันทำซ้ำอีกครั้ง
'                        otmEDC.Enabled = True   'เปิดใหม่อีกครั้ง
'                        Exit Sub
'                    Else
'                        Unload Me
'                        Exit Sub
'                    End If
'                Else
'                    MsgBox "Invalid Response code", vbCritical + vbOKOnly, "EDC ERROR"
'                    Unload Me
'                    Exit Sub
'                End If
                tMsg = SP_GETtMsgErrEDC("TA")                                       '*Ao 2014/08/07 CommSheet TK-ISS3000-543-PH2  V 4.2.140
                tMsg = tMsg & vbCrLf & "Do you want to try again?"
                If SP_MSGnShowing("tms_CN005065" & ";" & tMsg, nCS_OkCancel) = vbOK Then
                    SleepEx 3000, 0             '*Ao 2014/08/14   รอให้ EDC เริ่มต้นใหม่
                    otmEDC.Enabled = True   'เปิดใหม่อีกครั้ง
                    Exit Sub
                Else
                    Unload Me
                    Exit Sub
                End If
                    
            End If
        End If
    End If
    Exit Sub
ErrHandle:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Sub

Private Sub W_EDCxEmpCardEDC()
'________________________________
'   Call:
'   Cmt:    '*Ao 2013/01/28
'________________________________
Dim tSql As String, tEdcPort As String
Dim orsTemp As ADODB.Recordset
Dim bChk As Boolean

On Error GoTo ErrHandle
    
    nW_Timer = 0
    nW_TimeESC = 3
    bChk = SP_EDCbCheckEDC
    tW_Comport = tVB_EDCComport
    If Not bVB_CNEdcPos Then
        Call W_PRCxCase("2013_1203")
        Call SP_OBJxSetFocus(otbEmpCard)
        tW_EmpRemark = ""
    Else
        olaForm(8).Caption = nW_TimeESC
        Call W_PRCxCase("2013_1200")
    End If
    Exit Sub
ErrHandle:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Sub

'Private Function W_CHKbEmpCardMember(ByVal ptEmpCard As String) As Boolean
'Dim tSql As String, tTime As String, tTimeST  As String, tTimeED As String
'Dim orsTemp As ADODB.Recordset
'
'On Error GoTo ErrHandle
'
'    W_CHKbEmpCardMember = False
'
'    tTime = Format(Time$, "hh:mm")
'    tTime = Replace(tTime, ":", "")
'    tTimeST = SP_FEDtChkString(orsMemCard, "FTEstSTime1")
'    tTimeED = SP_FEDtChkString(orsMemCard, "FTEstETime1")
'    If tTime >= tTimeST And tTime <= tTimeED Then Exit Function
'    tTimeST = SP_FEDtChkString(orsMemCard, "FTEstSTime2")
'    tTimeED = SP_FEDtChkString(orsMemCard, "FTEstETime2")
'    If tTime >= tTimeST And tTime <= tTimeED Then Exit Function
'    tTimeST = SP_FEDtChkString(orsMemCard, "FTEstSTime3")
'    tTimeED = SP_FEDtChkString(orsMemCard, "FTEstETime3")
'    If tTime >= tTimeST And tTime <= tTimeED Then Exit Function
'    cW_MaxDis = SP_FEDvChkNumeric(orsMemCard, "FCEcdMaxDis")
'    W_CHKbEmpCardMember = True
'    Exit Function
'ErrHandle:
'    W_CHKbEmpCardMember = False
'    Call SP_MSGShowError
'End Function

Private Function W_PRCxFindEmpCardByID(ByVal ptCusID$, Optional ByVal pbChk As Boolean = True) As Boolean
'--------------------------------------
'Call:
'Cmt: ค้นหา Format ของเลขที่บัตร Mall Card อยู่ในช่วงใด
'--------------------------------------
    Dim tSql$
    'Dim nChkDigit As Long
    Dim orsTemp As ADODB.Recordset
On Error GoTo ErrHandle
     Call SP_STAxMouse(Me, True)
    W_PRCxFindEmpCardByID = False
    nW_CaseErr = 3
    
    tW_CardName = ""    '*Neung 55-02-10 Urgent Check for Employee Discount
    
     '*KT 52-05-26 PH 1.8 ML-SF-079 ระบบไม่ทำการตรวจสอบ Rang ของ Employee Card
    tSql = "SELECT  FTEcdCode,FTEcdCardType,FTEcdDesc1,FTEcdDesc2,FCEcdAmtRank"
    tSql = tSql & vbCrLf & ",FCEcdRate,FTEcdStaRang,FTEcdFrom,FTEcdTo,FTEcdStaChkExp,FTEcdStaDeptCode"
    tSql = tSql & vbCrLf & ",FTEcdStaCreditLimit,FTEcdStaChkAlwBuy"
    tSql = tSql & vbCrLf & " FROM   TCNMEmpCardMtn"
    tSql = tSql & vbCrLf & " WHERE '" & ptCusID & "' BETWEEN FTEcdFrom AND FTEcdTo"
    tSql = tSql & vbCrLf & "  AND  LEN(" & ptCusID & ") <= LEN(FTEcdTo)"
    tSql = tSql & vbCrLf & " AND FTEcdStaRang = 'Y'"
    tSql = tSql & vbCrLf & " ORDER BY LEN(FTEcdFrom)"
    Call SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine, , , , , True, True) 'read master
    If Not (orsTemp.BOF Or orsTemp.EOF) Then
        If ocnMember Is Nothing Then
                nW_CaseErr = 4
                olaName.Caption = SP_FEDtChkString(orsTemp, "FTEcdDesc1")
                cVB_DiscEmpRate = SP_FEDvChkNumeric(orsTemp, "FCEcdRate")
                olaRate.Caption = cVB_DiscEmpRate
                tVB_DiscEmpCardType = SP_FEDtChkString(orsTemp, "FTEcdCardType")
                tVB_DiscEmpCardKey = SP_FEDtChkString(orsTemp, "FTEcdCode")
                tW_ChkExp = SP_FEDtChkString(orsTemp, "FTEcdStaChkExp")
                tW_ChkDpt = SP_FEDtChkString(orsTemp, "FTEcdStaDeptCode")
                tW_ChkLimit = SP_FEDtChkString(orsTemp, "FTEcdStaCreditLimit")
                tVB_DiscEmpChkLimit = tW_ChkLimit '*KT 51-09-19 ML-SF- 558 Print  Remain Discount for Employee Card
                tW_ChkBay = SP_FEDtChkString(orsTemp, "FTEcdStaChkAlwBuy")
                tW_EcdFrom = SP_FEDtChkString(orsTemp, "FTEcdFrom")
                tW_EcdTo = SP_FEDtChkString(orsTemp, "FTEcdTo")
                cVB_DiscEmpRanking = SP_FEDvChkNumeric(orsTemp, "FCEcdAmtRank")
                tW_CardName = "" '*KT 52-11-26 PH2-ML-SF-170 ใน Ticket ไม่พิมพ์ชื่อพนักงานที่ใช้ส่วนลด Employee discount
                W_PRCxFindEmpCardByID = True
                Call SP_DATxRsNothing(orsTemp)
                Call SP_STAxMouse(Me, False)
                Exit Function
        End If
        
        If UCase(SP_FEDtChkString(orsTemp, "FTEcdStaCreditLimit")) = "N" And UCase(SP_FEDtChkString(orsTemp, "FTEcdStaChkAlwBuy")) = "N" Then
                nW_CaseErr = 5
                olaName.Caption = SP_FEDtChkString(orsTemp, "FTEcdDesc1")
                cVB_DiscEmpRate = SP_FEDvChkNumeric(orsTemp, "FCEcdRate")
                olaRate.Caption = cVB_DiscEmpRate
                tVB_DiscEmpCardType = SP_FEDtChkString(orsTemp, "FTEcdCardType")
                tVB_DiscEmpCardKey = SP_FEDtChkString(orsTemp, "FTEcdCode")
                tW_ChkExp = SP_FEDtChkString(orsTemp, "FTEcdStaChkExp")
                tW_ChkDpt = SP_FEDtChkString(orsTemp, "FTEcdStaDeptCode")
                tW_ChkLimit = SP_FEDtChkString(orsTemp, "FTEcdStaCreditLimit")
                tVB_DiscEmpChkLimit = tW_ChkLimit '*KT 51-09-19 ML-SF- 558 Print  Remain Discount for Employee Card
                tW_ChkBay = SP_FEDtChkString(orsTemp, "FTEcdStaChkAlwBuy")
                tW_EcdFrom = SP_FEDtChkString(orsTemp, "FTEcdFrom")
                tW_EcdTo = SP_FEDtChkString(orsTemp, "FTEcdTo")
                cVB_DiscEmpRanking = SP_FEDvChkNumeric(orsTemp, "FCEcdAmtRank")
               
                Call SP_DATxRsClose(orsMemCard)
                tSql = "SELECT  Top 1  TCNMEmpCard.FTEcdCode,TCNMEmpCard.FCEcdMaxDis,TCNMEmpCard.FTEcdName"
                tSql = tSql & vbCrLf & "FROM TCNMEmpCard "
                tSql = tSql & vbCrLf & " WHERE TCNMEmpCard.FTEcdCode = '" & ptCusID & "'"
                If SP_SQLvExecute(tSql, orsMemCard, ocnMember, , , , True) = 0 Then
                    If Not (orsMemCard.BOF Or orsMemCard.EOF) Then
                         tW_CardName = SP_FEDvChkNumeric(orsMemCard, "FTEcdName") '*KT 52-11-26 PH2-ML-SF-170 ใน Ticket ไม่พิมพ์ชื่อพนักงานที่ใช้ส่วนลด Employee discount
                    Else
                         tW_CardName = "" '*KT 52-11-26 PH2-ML-SF-170 ใน Ticket ไม่พิมพ์ชื่อพนักงานที่ใช้ส่วนลด Employee discount
                    End If
                Else
                    tW_CardName = "" '*KT 52-11-26 PH2-ML-SF-170 ใน Ticket ไม่พิมพ์ชื่อพนักงานที่ใช้ส่วนลด Employee discount
                End If
               
                W_PRCxFindEmpCardByID = True
                Call SP_DATxRsNothing(orsTemp)
                Call SP_STAxMouse(Me, False)
                Exit Function
        End If
On Error GoTo ErrHandle2
        tSql = "SELECT  Top 1  TCNMEmpCard.FTEcdCode,TCNMEmpCard.FCEcdMaxDis"
        tSql = tSql & vbCrLf & ",TCNMEmpShopping.FTEcdPosition,TCNMEmpShopping.FTEstAlwFlag,TCNMEmpShopping.FDEstDate"
        tSql = tSql & vbCrLf & ",TCNMEmpShopping.FTEstSTime1,TCNMEmpShopping.FTEstETime1,TCNMEmpShopping.FTEstSTime2"
        tSql = tSql & vbCrLf & ",TCNMEmpShopping.FTEstETime2 ,TCNMEmpShopping.FTEstSTime3,TCNMEmpShopping.FTEstETime3"
        tSql = tSql & vbCrLf & ",TCNMEmpCard.FTEcdName"
        tSql = tSql & vbCrLf & "FROM TCNMEmpCard INNER JOIN"
        tSql = tSql & vbCrLf & "TCNMEmpShopping  ON TCNMEmpCard.FTEcdPosition = TCNMEmpShopping.FTEcdPosition"
        tSql = tSql & vbCrLf & " WHERE TCNMEmpCard.FTEcdCode = '" & ptCusID & "'"
        tSql = tSql & vbCrLf & "AND FDEstDate = '" & Format(SP_GETdChkDateDisc, "yyyy/MM/dd") & "'"
        Call SP_SQLvExecute(tSql, orsMemCard, ocnMember, , , , True)
         If Not (orsMemCard.BOF Or orsMemCard.EOF) Then
            If UCase(SP_FEDtChkString(orsMemCard, "FTEstAlwFlag")) = "Y" Then '*KT 52-06-04 ตรวจสอบวันที่หมดอายุตามที่กำหนดใน parameter
                nW_CaseErr = 0
                olaName.Caption = SP_FEDtChkString(orsTemp, "FTEcdDesc1")
                cVB_DiscEmpRate = SP_FEDvChkNumeric(orsTemp, "FCEcdRate")
                olaRate.Caption = cVB_DiscEmpRate
                tVB_DiscEmpCardType = SP_FEDtChkString(orsTemp, "FTEcdCardType")
                tVB_DiscEmpCardKey = SP_FEDtChkString(orsTemp, "FTEcdCode")
                tW_ChkExp = SP_FEDtChkString(orsTemp, "FTEcdStaChkExp")
                tW_ChkDpt = SP_FEDtChkString(orsTemp, "FTEcdStaDeptCode")
                tW_ChkLimit = SP_FEDtChkString(orsTemp, "FTEcdStaCreditLimit")
                tVB_DiscEmpChkLimit = tW_ChkLimit '*KT 51-09-19 ML-SF- 558 Print  Remain Discount for Employee Card
                tW_ChkBay = SP_FEDtChkString(orsTemp, "FTEcdStaChkAlwBuy")
                tW_EcdFrom = SP_FEDtChkString(orsTemp, "FTEcdFrom")
                tW_EcdTo = SP_FEDtChkString(orsTemp, "FTEcdTo")
                tW_CardName = SP_FEDvChkNumeric(orsMemCard, "FTEcdName") '*KT 52-11-26 PH2-ML-SF-170 ใน Ticket ไม่พิมพ์ชื่อพนักงานที่ใช้ส่วนลด Employee discount
                cVB_DiscEmpRanking = SP_FEDvChkNumeric(orsTemp, "FCEcdAmtRank")
                cW_MaxDis = SP_FEDvChkNumeric(orsMemCard, "FCEcdMaxDis")    '*Ao 09-12-18 PH3-ML-SF-050
                W_PRCxFindEmpCardByID = True
            Else
                nW_CaseErr = 2 'ไม่อนุญาตให้มีการด้วย Emp Card  นี้ แต่มีการเก็บการรูด Employee Card
                olaName.Caption = SP_FEDtChkString(orsTemp, "FTEcdDesc1")
                cVB_DiscEmpRate = SP_FEDvChkNumeric(orsTemp, "FCEcdRate")
                olaRate.Caption = cVB_DiscEmpRate
                tVB_DiscEmpCardType = SP_FEDtChkString(orsTemp, "FTEcdCardType")
                tVB_DiscEmpCardKey = SP_FEDtChkString(orsTemp, "FTEcdCode")
                tW_ChkExp = SP_FEDtChkString(orsTemp, "FTEcdStaChkExp")
                tW_ChkDpt = SP_FEDtChkString(orsTemp, "FTEcdStaDeptCode")
                tW_ChkLimit = SP_FEDtChkString(orsTemp, "FTEcdStaCreditLimit")
                tVB_DiscEmpChkLimit = tW_ChkLimit '*KT 51-09-19 ML-SF- 558 Print  Remain Discount for Employee Card
                tW_ChkBay = SP_FEDtChkString(orsTemp, "FTEcdStaChkAlwBuy")
                tW_EcdFrom = SP_FEDtChkString(orsTemp, "FTEcdFrom")
                tW_EcdTo = SP_FEDtChkString(orsTemp, "FTEcdTo")
                tW_CardName = SP_FEDvChkNumeric(orsMemCard, "FTEcdName") '*KT 52-11-26 PH2-ML-SF-170 ใน Ticket ไม่พิมพ์ชื่อพนักงานที่ใช้ส่วนลด Employee discount
                cVB_DiscEmpRanking = SP_FEDvChkNumeric(orsTemp, "FCEcdAmtRank")
                cW_MaxDis = SP_FEDvChkNumeric(orsMemCard, "FCEcdMaxDis")    '*Ao 09-12-18 PH3-ML-SF-050
                W_PRCxFindEmpCardByID = True
            End If
        Else
            nW_CaseErr = 1 'ไม่พบหมายเลข Emp Card ใน Member
            olaName.Caption = SP_FEDtChkString(orsTemp, "FTEcdDesc1")
            cVB_DiscEmpRate = SP_FEDvChkNumeric(orsTemp, "FCEcdRate")
            olaRate.Caption = cVB_DiscEmpRate
            tVB_DiscEmpCardType = SP_FEDtChkString(orsTemp, "FTEcdCardType")
            tVB_DiscEmpCardKey = SP_FEDtChkString(orsTemp, "FTEcdCode")
            tW_ChkExp = SP_FEDtChkString(orsTemp, "FTEcdStaChkExp")
            tW_ChkDpt = SP_FEDtChkString(orsTemp, "FTEcdStaDeptCode")
            tW_ChkLimit = SP_FEDtChkString(orsTemp, "FTEcdStaCreditLimit")
            tVB_DiscEmpChkLimit = tW_ChkLimit '*KT 51-09-19 ML-SF- 558 Print  Remain Discount for Employee Card
            tW_ChkBay = SP_FEDtChkString(orsTemp, "FTEcdStaChkAlwBuy")
            tW_EcdFrom = SP_FEDtChkString(orsTemp, "FTEcdFrom")
            tW_EcdTo = SP_FEDtChkString(orsTemp, "FTEcdTo")
            cVB_DiscEmpRanking = SP_FEDvChkNumeric(orsTemp, "FCEcdAmtRank")
            
            Call SP_DATxRsClose(orsMemCard)
            tSql = "SELECT  Top 1  TCNMEmpCard.FTEcdCode,TCNMEmpCard.FCEcdMaxDis,TCNMEmpCard.FTEcdName"
            tSql = tSql & vbCrLf & "FROM TCNMEmpCard "
            tSql = tSql & vbCrLf & " WHERE TCNMEmpCard.FTEcdCode = '" & ptCusID & "'"
            If SP_SQLvExecute(tSql, orsMemCard, ocnMember, , , , True) = 0 Then
                If Not (orsMemCard.BOF Or orsMemCard.EOF) Then
                     tW_CardName = SP_FEDvChkNumeric(orsMemCard, "FTEcdName") '*KT 52-11-26 PH2-ML-SF-170 ใน Ticket ไม่พิมพ์ชื่อพนักงานที่ใช้ส่วนลด Employee discount
                Else
                     tW_CardName = "" '*KT 52-11-26 PH2-ML-SF-170 ใน Ticket ไม่พิมพ์ชื่อพนักงานที่ใช้ส่วนลด Employee discount
                End If
            Else
                tW_CardName = "" '*KT 52-11-26 PH2-ML-SF-170 ใน Ticket ไม่พิมพ์ชื่อพนักงานที่ใช้ส่วนลด Employee discount
            End If
            
            W_PRCxFindEmpCardByID = True
        End If
    Else
         olaName.Caption = ""
         olaRate.Caption = ""
    End If
    Call SP_DATxRsNothing(orsTemp)
    Call SP_STAxMouse(Me, False)
    Exit Function
ErrHandle:
    nW_CaseErr = 4
    olaName.Caption = SP_FEDtChkString(orsTemp, "FTEcdDesc1")
    cVB_DiscEmpRate = SP_FEDvChkNumeric(orsTemp, "FCEcdRate")
    olaRate.Caption = cVB_DiscEmpRate
    tVB_DiscEmpCardType = SP_FEDtChkString(orsTemp, "FTEcdCardType")
    tVB_DiscEmpCardKey = SP_FEDtChkString(orsTemp, "FTEcdCode")
    tW_ChkExp = SP_FEDtChkString(orsTemp, "FTEcdStaChkExp")
    tW_ChkDpt = SP_FEDtChkString(orsTemp, "FTEcdStaDeptCode")
    tW_ChkLimit = SP_FEDtChkString(orsTemp, "FTEcdStaCreditLimit")
    tVB_DiscEmpChkLimit = tW_ChkLimit '*KT 51-09-19 ML-SF- 558 Print  Remain Discount for Employee Card
    tW_ChkBay = SP_FEDtChkString(orsTemp, "FTEcdStaChkAlwBuy")
    tW_EcdFrom = SP_FEDtChkString(orsTemp, "FTEcdFrom")
    tW_EcdTo = SP_FEDtChkString(orsTemp, "FTEcdTo")
    tW_CardName = "" '*KT 52-11-26 PH2-ML-SF-170 ใน Ticket ไม่พิมพ์ชื่อพนักงานที่ใช้ส่วนลด Employee discount
    cVB_DiscEmpRanking = SP_FEDvChkNumeric(orsTemp, "FCEcdAmtRank")
    W_PRCxFindEmpCardByID = True
    Call SP_DATxRsNothing(orsTemp)
    Call SP_STAxMouse(Me, False)
    Exit Function
ErrHandle2:
    Select Case Err.Number
        Case -2147467259, 91, -2147217871, 3704
                Call SP_STAxMouse(Me, True)
                Call SP_CONbMemberServer
                Call SP_STAxMouse(Me, False)
        Case Else
                Call SP_MSGxShowError(Err.Number, Err.Description)
    End Select
    nW_CaseErr = 4
    olaName.Caption = SP_FEDtChkString(orsTemp, "FTEcdDesc1")
    cVB_DiscEmpRate = SP_FEDvChkNumeric(orsTemp, "FCEcdRate")
    olaRate.Caption = cVB_DiscEmpRate
    tVB_DiscEmpCardType = SP_FEDtChkString(orsTemp, "FTEcdCardType")
    tVB_DiscEmpCardKey = SP_FEDtChkString(orsTemp, "FTEcdCode")
    tW_ChkExp = SP_FEDtChkString(orsTemp, "FTEcdStaChkExp")
    tW_ChkDpt = SP_FEDtChkString(orsTemp, "FTEcdStaDeptCode")
    tW_ChkLimit = SP_FEDtChkString(orsTemp, "FTEcdStaCreditLimit")
    tVB_DiscEmpChkLimit = tW_ChkLimit '*KT 51-09-19 ML-SF- 558 Print  Remain Discount for Employee Card
    tW_ChkBay = SP_FEDtChkString(orsTemp, "FTEcdStaChkAlwBuy")
    tW_EcdFrom = SP_FEDtChkString(orsTemp, "FTEcdFrom")
    tW_EcdTo = SP_FEDtChkString(orsTemp, "FTEcdTo")
    tW_CardName = "" '*KT 52-11-26 PH2-ML-SF-170 ใน Ticket ไม่พิมพ์ชื่อพนักงานที่ใช้ส่วนลด Employee discount
    cVB_DiscEmpRanking = SP_FEDvChkNumeric(orsTemp, "FCEcdAmtRank")
    W_PRCxFindEmpCardByID = True
    Call SP_DATxRsNothing(orsTemp)
    Call SP_STAxMouse(Me, False)
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)         '*Ao 2014/08/07 CommSheet TK-ISS3000-544-PH2  V 4.2.140
     'Debug.Print "Form_KeyDown"
    If bVB_CNSplash Then KeyCode = 0 '*KT 51-09-17 ป้องการ Key กรณีที่ Splash ขึ้น
'     If KeyCode = 13 Then SendKeys "{Tab}" '*KT 53-01-19 ไม่ใช้งานแล้ว ทำให้เกิดการ setfocus มายัง TabIndex 1

    Select Case KeyCode
        Case vbKeyEscape
            If bW_HandleESC Then
'                Call W_PRCxCase("2013_1205")    '*Eaw 56-06-24 กด Esc เพื่อใช้ manual Mode ไม่ได้ CommSheet TK-ISS3000-225-ST 00 00 04 FSBT NG_eaw.xlsx
'                Call SP_OBJxSetFocus(otbEmpCard)   '*Eaw 56-06-06  CommSheet TK-ISS3000-225-ST 00 00 01.xls แยกเคส Reentry Mode กับ Normal Mode
                If bVB_ReEntryMode Then Call SP_OBJxSetFocus(otbEmpCard)
            Else
'                Call ocmCancel_Click            '*Ao 2014/08/07 CommSheet TK-ISS3000-544-PH2  V 4.2.140
            End If
        End Select
End Sub

Private Sub Form_Load()
On Error GoTo ErrHandle
    bVB_DiscAccept = False
    bW_FoundEmpCard = False
    bW_KeyEnter = False '***KT 51-12-01 V 1.0.1
    bW_ChkESC = False                                                   '*Ao 2014/08/07 CommSheet TK-ISS3000-544-PH2  V 4.2.140
    bVB_ChkEDC = False
    Call SP_FRMxChildCenter(Me)
    Call SP_FRMxGetCaptionFrmDB2Tag(Me)
    Call SP_FRMxSetCapForm(Me)
    Call SP_OBJxSetFocus(otbEmpCard)
'    otbCV.BackColor = nVB_CNClrBLock
    If bVB_ReEntryMode Then                                     '*Eaw 56-06-06  CommSheet TK-ISS3000-225-ST 00 00 01.xls แยกเคส Reentry Mode กับ Normal Mode
        Call SP_OBJxSetFocus(otbEmpCard)
    Else
        otbEmpCard.Locked = True                                    '*Ao 2014/05/30
        otbEmpCard.BackColor = nVB_CNClrBLock
        orbMonth.Locked = True
        orbMonth.BackColor = nVB_CNClrBLock
        orbYear.Locked = True
        orbYear.BackColor = nVB_CNClrBLock
    End If
    If bVB_Res640X480 Then
        Me.Top = 2100
    End If
    If bVB_ReEntryMode Then               '*Eaw 56-06-06  CommSheet TK-ISS3000-225-ST 00 00 01.xls เพิ่ม If ครอบไว้หากเป็น Reentry Mode ให้ใช้ MSR ส่วน Normal Mode ไม่ให้ใช้ MSR
        If bVB_OPOSMSREnable Then
            With OPOSMSR1
                'Open the device
                'Use a Logical Device Name which has been set on the SetupPOS.
                .Open tVB_CashMsr
                
                'Get the exclusive control right for the opened device.
                'Then the device is disable from other application.
                
                '(Notice:When using an old CO, use the Claim.)
    '            .ClaimDevice 1000 'Team3000
                .Claim 1000 'TeamPos2000
        
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
'    Call W_EDCxEmpCardEDC
    If Not bVB_ReEntryMode Then  '*Eaw 56-06-27 ถ้าเป็น Re-Entry ไม่ใช้ EDC
        Call W_EDCxEmpCardEDC
    End If
    bW_ChkMEM = False
    Exit Sub
ErrHandle:
    bW_ChkMEM = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    If bVB_ReEntryMode Then     '*Eaw 56-06-06  CommSheet TK-ISS3000-225-ST 00 00 01.xls เพิ่ม If ครอบไว้หากเป็น Reentry Mode ให้ใช้ MSR ส่วน Normal Mode ไม่ให้ใช้ MSR
        With OPOSMSR1
            If bVB_OPOSMSREnable Then
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
    End If
    With oScanner
        If bVB_OPOSScannerEnable Then
            If .DeviceEnabled Then
                .DeviceEnabled = False
                .Release
                .Close
            End If
        End If
    End With
'    Set wCNDiscEmp = Nothing
    Call SP_OBJxSetNothing(Me)
End Sub

Private Sub ocmCancel_Click()
'    If bVB_OPOSMSREnable Then OPOSMSR1.DataEventEnabled = False         '*Ao 10-02-09
    If bW_ChkESC Then               '*Ao 2014/08/07 CommSheet TK-ISS3000-544-PH2  V 4.2.140
        Exit Sub
    End If
    If bVB_ReEntryMode Then  '*Eaw 56-06-06  CommSheet TK-ISS3000-225-ST 00 00 01.xls เพิ่ม If ครอบไว้หากเป็น Reentry Mode ให้ใช้ MSR ส่วน Normal Mode ไม่ให้ใช้ MSR
        If bVB_OPOSMSREnable Then OPOSMSR1.DataEventEnabled = False
    End If
    If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = False
    bVB_DiscAccept = False
    Unload Me
End Sub

Private Sub ocmOK_Click()
On Error GoTo ErrHandle
    If bVB_OPOSMSREnable Then OPOSMSR1.DataEventEnabled = False         '*Ao 10-02-09
    If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = False
    
'    If Not bW_FoundEmpCard Then
'        bW_KeyEnter = False '***KT 51-12-01 V 1.0.1
'        If bVB_OPOSMSREnable Then OPOSMSR1.DataEventEnabled = True          '*Ao 10-02-09
'        If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True
'        Exit Sub
'    End If
    If UCase(tW_ChkExp) = "Y" Then 'Comsheet ML-SF-171
        If orbMonth.Text = "" Or orbYear.Text = "" Then
            Call SP_MSGnShowing("tms_CN005142", nCS_Error)
            If orbMonth.Text = "" Then
                Call SP_OBJxSetFocus(orbMonth)
            Else
                Call SP_OBJxSetFocus(orbYear)
            End If
            bW_KeyEnter = False '***KT 51-12-01 V 1.0.1
            If bVB_OPOSMSREnable Then OPOSMSR1.DataEventEnabled = True          '*Ao 10-02-09
            If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True
            Exit Sub
        End If
        If Format(orbMonth.Text, "00") = "00" Then         '*Ao 52-02-05
            Call SP_MSGnShowing("tms_CN005142", nCS_Error)
            orbMonth.Text = ""
            orbYear.Text = ""
            Call SP_OBJxSetFocus(orbMonth)
            bW_KeyEnter = False '***KT 51-12-01 V 1.0.1
            If bVB_OPOSMSREnable Then OPOSMSR1.DataEventEnabled = True          '*Ao 10-02-09
            If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True
            Exit Sub
        End If
        '____________________________________________________
        dVB_DiscEmpCardExp = DateSerial(2000 + Val(Format(orbYear.Text)), Val(Format(orbMonth.Text)) + 1, 1)
         If Format(dVB_DiscEmpCardExp, "YYYY/MM/DD") <= Format(dVB_CurrentDate, "YYYY/MM/DD") Then 'ทาง FSBT ต้องการให้ตรวจสอบวันที่ Expire Date ตาม System Date
            Call SP_MSGnShowing("tms_CN005142", nCS_Error)
            orbMonth.Text = ""        '*AO51-03-20      Comm.Sheet SF-108
            orbYear.Text = ""
            Call SP_OBJxSetFocus(orbMonth)
            bW_KeyEnter = False '***KT 51-12-01 V 1.0.1
            If bVB_OPOSMSREnable Then OPOSMSR1.DataEventEnabled = True          '*Ao 10-02-09
            If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True
            Exit Sub
        End If
    Else
        If orbYear.Text = "" Or orbMonth.Text = "" Then
            Call SP_MSGnShowing("tms_CN005142", nCS_Error)
            If orbMonth.Text = "" Then
                Call SP_OBJxSetFocus(orbMonth) 'orbMonth.SetFocus
            Else
                Call SP_OBJxSetFocus(orbYear) 'orbYear.SetFocus
            End If
            bW_KeyEnter = False '***KT 51-12-01 V 1.0.1
            If bVB_OPOSMSREnable Then OPOSMSR1.DataEventEnabled = True          '*Ao 10-02-09
            If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True
            Exit Sub
        End If
    End If
    
    tVB_DiscEmpCardID = Trim(otbEmpCard.Text)
    '____________________________________
    oDisc.bReConnectMember = False                          'ไม่ต้อง เช็ค Member
    oDisc.tEmplimit = tW_PURCHASE
    oDisc.tEmpRemark = tW_EmpRemark                 'EDC
    oDisc.bRecalByItem = False
    oDisc.DISxClearAll
    oDisc.dTranDate = tUT_SaleDate
    oDisc.tTerminalNo = tVB_CNTerminalNum
    oDisc.tTransNo = tVB_Running
    oDisc.tCardID = tVB_DiscEmpCardID
    oDisc.tCardType = tVB_DiscEmpCardType
    oDisc.tCardKey = tVB_DiscEmpCardKey
    oDisc.cDiscValue = SP_VALcText2Double(olaRate.Caption)
    dVB_DiscEmpCardExp = DateSerial(2000 + Val(Format(orbYear.Text)), Val(Format(orbMonth.Text)), 1)
    oDisc.dCardExpDate = dVB_DiscEmpCardExp
    oDisc.tStaChkExpire = tW_ChkExp
    oDisc.tStaChkDepart = tW_ChkDpt
    oDisc.cAlwDis = cVB_DiscEmpRanking
    oDisc.tRefCode = tW_ChkBay & "," & tW_ChkLimit & "," & tW_ChkAlwPoint '*BG 56-08-01 CommSheet TK-ISS3000-196-ST 00 00 02.xls
    oDisc.tCusID = tW_EcdFrom & "," & tW_EcdTo
    oDisc.tCardName = tW_CardName '*KT 52-11-26 PH2-ML-SF-170 ใน Ticket ไม่พิมพ์ชื่อพนักงานที่ใช้ส่วนลด Employee discount
    oDisc.bChkEMP = False               'ใช้ INSERT EMP    INSERT ลง CD
    Call oDisc.DIScDiscount(nEN_Disc12EmpCrd, eEN_VerifyDiscType.nEN_Type2ByDcn)
    bVB_DiscAccept = True
    If tW_CardName <> "" Then
        ofaSwipeCard.Left = 3480        'ซ่อน
        ofaSwipeCard.Top = 5760
        DoEvents
        SleepEx 4000, 0     '1000 = 1 วินาที                แก้ 9000 --> 4000
    End If
    If tW_ChkAlwPoint = "N" Then '*Eaw 56-08-21
        bVB_ChkEmpPoint = False
    Else
        bVB_ChkEmpPoint = True
    End If
    bVB_ChkEDC = True                '*Ao 2014/07/04
    Unload Me
    Exit Sub
ErrHandle:
    bW_KeyEnter = False
    Call SP_MSGxShowError(Err.Number, Err.Description)
    If bVB_OPOSMSREnable Then OPOSMSR1.DataEventEnabled = True          '*Ao 10-02-09
    If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True
End Sub

Private Sub OPOSMSR1_DataEvent(ByVal Status As Long)
    Dim tExpDate$
    Dim tTranc2$
On Error GoTo ErrHandle '*KT 52-02-11  PH2-ML-SF-092 โปรแกรมหลุดมาที่ Desktop ตอนพิมพ์ receipt , ชำระด้วย Credit
    With oScanner                                                           '*Ao 10-02-09
     If bVB_OPOSScannerEnable Then
            'Enable the event.
            .DataEventEnabled = False
     End If
     End With
    With OPOSMSR1
        'otbEmpCard.Text = .AccountNumber     '(Account No.)
        'tExpDate = .ExpirationDate    '(Valid date)
        tTranc2 = .Track2Data
        .ClearInput
        .DataEventEnabled = True
         If tTranc2 <> "" Then
             If Len(Trim(Mid(tTranc2, 1, InStr(1, tTranc2, "=") - 1))) <= 20 Then   '*KT 51-06-18 Communication Sheet ML-SF-349
                otbEmpCard.Text = Trim(Mid(tTranc2, 1, InStr(1, tTranc2, "=") - 1))
                tExpDate = Trim(Mid(tTranc2, InStr(1, tTranc2, "=") + 1, Len(tTranc2) - InStr(1, tTranc2, "=")))
                orbYear.Text = Mid(tExpDate, 1, 2)
                orbMonth.Text = Mid(tExpDate, 3, 2)
            Else
                Call SP_MSGnShowing("tms_CN005069", nCS_Error)
                bW_FoundEmpCard = False
                otbEmpCard.Text = ""
                orbMonth.Text = ""
                orbYear.Text = ""
                olaName.Caption = ""
                olaRate.Caption = ""
                Call SP_OBJxSetFocus(otbEmpCard)
                Call SP_OTBxSelText(otbEmpCard)
                If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True          '*Ao 10-02-09
                Unload Me '*Eaw 56-06-25 CommSheet TK-ISS3000-225-ST 00 00 05 FSBT NG.xlsx
                Exit Sub
            End If
        Else
            Call SP_MSGnShowing("tms_CN005069", nCS_Error)
            bW_FoundEmpCard = False
            otbEmpCard.Text = ""
            orbMonth.Text = ""
            orbYear.Text = ""
            olaName.Caption = ""
            olaRate.Caption = ""
            Call SP_OBJxSetFocus(otbEmpCard)
            Call SP_OTBxSelText(otbEmpCard)
            If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True          '*Ao 10-02-09
            Unload Me '*Eaw 56-06-25 CommSheet TK-ISS3000-225-ST 00 00 05 FSBT NG.xlsx
            Exit Sub
        End If
        
'        tVB_DiscEmpCardID = Mid(Trim(otbEmpCard.Text), 9, 6)               '*Ao 2013/01/29
        tVB_DiscEmpCardID = Trim(otbEmpCard.Text)
         '*KT 51-07-14 Comsheet ML-SF- 356
        If Not W_PRCxFindEmpCardByID(tVB_DiscEmpCardID) Then
            Select Case nW_CaseErr
                Case 0
                    'none
                Case 1
                    'Call SP_MSGnShowing("tms_CN005186", nCS_Warning)
                Case 2
                    'Call SP_MSGnShowing("tms_CN005160", nCS_Warning)
                Case 3
                    Call SP_MSGnShowing("tms_CN005069", nCS_Warning)
                Case 4
                    'Call SP_MSGnShowing("tms_CN005069", nCS_Warning)
            End Select
            bW_FoundEmpCard = False
            otbEmpCard.Text = ""                      '*AO51-03-20    Comm.Sheet  SF-108
            orbMonth.Text = ""
            orbYear.Text = ""
            olaName.Caption = ""
            olaRate.Caption = ""
            Call SP_OBJxSetFocus(otbEmpCard)
            Call SP_OTBxSelText(otbEmpCard)
            If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True          '*Ao 10-02-09
            Exit Sub
        Else
            bW_FoundEmpCard = True
        End If

        If UCase(tW_ChkExp) = "Y" Then
            dVB_DiscEmpCardExp = DateSerial(2000 + Val(Format(orbYear.Text)), Val(Format(orbMonth.Text)) + 1, 1)
            If Format(dVB_DiscEmpCardExp, "YYYY/MM/DD") <= Format(dVB_CurrentDate, "YYYY/MM/DD") Then 'ทาง FSBT ต้องการให้ตรวจสอบวันที่ Expire Date ตาม System Date
                Call SP_MSGnShowing("tms_CN005079", nCS_Warning)
                bW_FoundEmpCard = False
                otbEmpCard.Text = ""
                orbMonth.Text = ""
                orbYear.Text = ""
                olaName.Caption = ""
                olaRate.Caption = ""
                Call SP_OBJxSetFocus(otbEmpCard)
                Call SP_OTBxSelText(otbEmpCard)
                If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True          '*Ao 10-02-09
                Exit Sub
            End If
        Else
            If orbYear.Text = "" Or orbMonth.Text = "" Then
                Call SP_OBJxSetFocus(orbMonth)
                If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True          '*Ao 10-02-09
                Exit Sub
            End If
        End If
    End With
    Call otbEmpCard_KeyPress(13)
    'ocmOk.SetFocus
    Call SP_OBJxSetFocus(ocmOk)
    If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True          '*Ao 10-02-09
    Exit Sub
ErrHandle:
    bW_FoundEmpCard = False
    Call SP_MSGxShowError(Err.Number, Err.Description)
    Call SP_OBJxSetFocus(otbEmpCard)
    If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True          '*Ao 10-02-09
End Sub

Private Sub orbMonth_GotFocus()
'    orbMonth.SelStart = 0
'    orbMonth.SelLength = Len(orbMonth.Text)
    Call SP_OTBxSelText(orbMonth)
End Sub

Private Sub orbMonth_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDelete         'delete
            orbMonth.Text = ""
        Case 13
            If Trim(orbMonth.Text) <> "" Then
                SendKeys "{Tab}"
            End If
    End Select
End Sub

Private Sub orbMonth_KeyPress(KeyAscii As Integer)
    Call UT_ChkKeyAscii(KeyAscii)
    Select Case KeyAscii
        Case 8, 6           'backspace, delete
            'do nothing
        Case 13
            'SendKeys "{TAB}"
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
'    orbYear.Text = ""
'    orbYear.SelStart = 0
'    orbYear.SelLength = Len(Trim(orbYear.Text))
    Call SP_OTBxSelText(orbYear)
End Sub

Private Sub orbYear_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDelete         'delete
            orbYear.Text = ""
        Case 13
            If Trim(orbYear.Text) <> "" Then
                SendKeys "{Tab}"
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
            '------------------------------------------------------------------
            If bW_KeyEnter Then
                KeyAscii = 0
            Else
                bW_KeyEnter = True
                Call ocmOK_Click
            End If
            '-------------------------------------------------------------------
        Case Else
            If Not IsNumeric(Chr(KeyAscii)) Then
                KeyAscii = 0
            End If
    End Select
End Sub

Private Sub oScanner_DataEvent(ByVal Status As Long)
On Error GoTo ErrHandle
    With OPOSMSR1                                   '*Ao 10-02-09
        If bVB_OPOSMSREnable Then
            'Enable the event.
            .DataEventEnabled = False
        End If
    End With
    With oScanner
        'otbEmpCard.Text = .ScanData '*Bump 56-02-28 Comment บรรทัดนี้
        '*Bump 56-02-28 เพิ่ม code 8 บรรทัด
        If .DeviceName = tSNR_TOUCH_SCANNER And Not .DecodeData Then
            .DecodeData = True
        End If
        If .DeviceName = tSNR_TOUCH_SCANNER Then
            otbEmpCard.Text = SP_SplitBarcodeData(.ScanDataType, .ScanData, .ScanDataLabel) '*ต้องย้ายจาก mCNSTP.SP_SplitBarcodeData  ไปที่ mCNSP.SP_SplitBarcodeData
        Else
            otbEmpCard.Text = .ScanData
        End If
'        DoEvents            '*KT 53-01-14
        If otbEmpCard.Text <> "" Then
            '*KT 52-05-06 PH 1.8 ML-SF-056 TP 2000 ไม่สามารถใช้ Scanner  Scan สินค้าได้ โดยเมื่อ Scan สินค้าแล้วมีรหัสสินค้าปรากฏขึ้นมาถูกต้องแต่ จะขึ้นข้อความไม่พบสินค้านั้น กรณีที่ OPOS บางตัวมีการอ่านค่า Enter ต่อท้าย ต้องตัดออก
            '*KT 52-05-15 PH 1.8 ML-SF-061 Price Check ไม่สามารถใช้ Scanner ได้ จะขึ้นข้อความไม่พบรหัสสินค้า
            If Asc(Right(otbEmpCard.Text, 1)) = 13 Then
                otbEmpCard.Text = Mid(otbEmpCard.Text, 1, Len(otbEmpCard.Text) - 1)
            End If
            Call otbEmpCard_KeyPress(13)
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
    If bVB_OPOSMSREnable Then OPOSMSR1.DataEventEnabled = True      '*Ao 10-02-09
    Exit Sub
ErrHandle:
    MsgBox "Scaner error" & vbCrLf & "Runtime : " & Err.Number & vbCrLf & Err.Description
    If bVB_OPOSMSREnable Then OPOSMSR1.DataEventEnabled = True      '*Ao 10-02-09
End Sub

'Private Sub otbEmpCard_Change()            '*Ao 2014/04/29
'    bW_FoundEmpCard = False
'End Sub

Private Sub otbEmpCard_GotFocus()
    otbEmpCard.SelStart = 0
    otbEmpCard.SelLength = Len(otbEmpCard.Text)
End Sub

Private Sub otbEmpCard_KeyDown(KeyCode As Integer, Shift As Integer)
    'Debug.Print "otbEmpCard_KeyDown"
    Select Case KeyCode
        Case vbKeyDelete         'delete
            otbEmpCard.Text = ""
            orbMonth.Text = ""
            orbYear.Text = ""
            olaName.Caption = ""
            olaRate.Caption = ""
            bW_FoundEmpCard = False
        Case vbKeyBack
            orbMonth.Text = ""
            orbYear.Text = ""
            olaName.Caption = ""
            olaRate.Caption = ""
            bW_FoundEmpCard = False
        Case 13
            If Trim(otbEmpCard.Text) <> "" Then
                SendKeys "{Tab}"
            End If
    End Select
End Sub

Private Sub otbEmpCard_KeyPress(KeyAscii As Integer)
'---------------------------------------------------------------------------------------
'   Call:
'   Cmt:
'---------------------------------------------------------------------------------------
On Error GoTo ErrHandle
    Call UT_ChkKeyAscii(KeyAscii)
    Call SP_ChkKeyAsciiNumber(KeyAscii)
    If Trim(otbEmpCard.Text) = "" Then Exit Sub
    If KeyAscii = 13 Then
        tVB_DiscEmpCardID = Trim(otbEmpCard.Text)
        If W_PRCbFindEmpCardByEDC(tVB_DiscEmpCardID, "", True) Then
'            Call SP_OBJxSetFocus(orbMonth)
            If bW_ChkMEM Then
                
            Else
                'ไม่เช็ค MEMBER
            End If
        Else
            '*Bump 56-02-22  อ้างอิง TK-ISS3000-058 Employee card expried and invalid number not show message
            Call SP_MSGnShowing("tms_CN005069", nCS_Error)
            bW_FoundEmpCard = False
            otbEmpCard.Text = ""
            orbMonth.Text = ""
            orbYear.Text = ""
            olaName.Caption = ""
            olaRate.Caption = ""
            Call SP_OBJxSetFocus(otbEmpCard)
        End If
    End If
ErrHandle:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Sub

'Private Sub otbEmpCard_KeyPress(KeyAscii As Integer)
'   Dim orsMem As ADODB.Recordset
'   Dim tSta As String
'   Dim tSql As String
'   Dim tSplash As String
'   Dim nErr As Long
'On Error GoTo ErrHandle
'    Call UT_ChkKeyAscii(KeyAscii)
'    Call SP_ChkKeyAsciiNumber(KeyAscii)
'    If otbEmpCard.Text <> "" Then
'        If KeyAscii = 13 Then
''            tVB_DiscEmpCardID = Mid(Trim(otbEmpCard.Text), 9, 6)            '*Ao 2013/01/29
'            tVB_DiscEmpCardID = Trim(otbEmpCard.Text)
'            If tVB_DiscEmpCardID = "" Then  '*Eaw 55-12-19 ถ้าตัดหรัสพนักงานได้เป็นค่าว่าจะแสดงข้อควมแจ้ง เพื่อให้ใส่ค่าใหม่
'                Call SP_MSGnShowing("tms_CN005069", nCS_Warning)
'                Call SP_OTBxSelText(otbEmpCard)
'                Exit Sub
'            Else
'                tSql = "SELECT Top 1  FTEcdCode,FTEcdSta "
'                tSql = tSql & vbCrLf & "FROM TCNMEmpCard"
'                'tSql = tSql & vbCrLf & "WHERE FTEcdCode ='" & Trim(otbEmpCard.Text) & "'"'*Bump 55-12-19 Comment บรรทัดนี้
'                tSql = tSql & vbCrLf & "WHERE FTEcdCode ='" & tVB_DiscEmpCardID & "'" '*Bump 55-12-19 ต้องเอาหมายเลขบัตรพนักงานที่ได้จาก digits ที่ 9 ถึง digits ที่ 14
'                nErr = SP_SQLvExecute(tSql, orsMem, ocnMember, , , , True)
'                If (orsMem.BOF And orsMem.EOF) Then
'                    Call SP_MSGnShowing("tms_CN005069", nCS_Warning)
'                    Call SP_OTBxSelText(otbEmpCard)
'                    Exit Sub
'                End If
'            End If
'            If bW_FoundEmpCard Then
'On Error GoTo ErrHandle2
'                Select Case nW_CaseErr
'                        Case 0, 1, 2
'                            '*KT 52-03-16
'                            If ocnMember Is Nothing And bVB_ReConnectEmp Then
'                                    Call SP_STAxMouse(Me, True)
'                                    tSplash = "Can not connect the database member, System is  reconnect database member again."  ' tms_CN005208
'                                    '*KT 52-05-26 PH 1.8 ML-SF-075 ในกรณีที่ POS ไม่สามารถติดต่อ Member Server ได้ ขอให้ POS แสดง Warning Message เพียงครั้งเดียว
'                                    Call SP_CONbMemberServer(True, tSplash, False)
'                                    bVB_ReConnectMem = False
'                                    Call SP_STAxMouse(Me, False)
'                            End If
'                            If Not ocnMember Is Nothing And oDB.SQLbAvailableServer(oDB.OffServerName, oDB.OnComIP) Then
'            '               If Not ocnMember Is Nothing And ocnAll.ConnectionString <> oDB.ocnOffLine.ConnectionString Then
'                                tSql = "SELECT Top 1  FTEcdCode,FTEcdSta "
'                                tSql = tSql & vbCrLf & "FROM TCNMEmpCard"
'                                'tSql = tSql & vbCrLf & "WHERE FTEcdCode ='" & Trim(otbEmpCard.Text) & "'"'*Bump 55-12-19 Comment บรรทัดนี้
'                                tSql = tSql & vbCrLf & "WHERE FTEcdCode ='" & tVB_DiscEmpCardID & "'" '*Bump 55-12-19 ต้องเอาหมายเลขบัตรพนักงานที่ได้จาก digits ที่ 9 ถึง digits ที่ 14
'                                nErr = SP_SQLvExecute(tSql, orsMem, ocnMember, , , , True)
'                                Select Case nErr
'                                    Case -2147467259, 91, -2147217871, 3704 '*KT 54-04-07 ป้องกันโปรแกรมหลุด 3.0.56
'                                        '*KT 52-05-26 PH 1.8 ML-SF-075 ในกรณีที่ POS ไม่สามารถติดต่อ Member Server ได้ ขอให้ POS แสดง Warning Message เพียงครั้งเดียว
'                                         '*KT 52-05-26 ให้ทำการแจ้ง message เตือนเมื่อทำการ Connect Member Server
'                                        'Call SP_MSGnShowing("tms_CN005134", tCS_CNMsgErr)
'                                End Select
'                                If Not (orsMem.BOF Or orsMem.EOF) Then
'                                    tSta = SP_FEDtChkString(orsMem, "FTEcdSta")
'                                    If UCase(tSta) = "1" Or UCase(tSta) = "" Or UCase(tSta) = "A" Then
'                                        bW_FoundEmpCard = True
'                                    Else
'                                        bW_FoundEmpCard = False
'                                        Select Case UCase(tSta)
'                                            Case "P", "2"
'                                                Call SP_MSGnShowing("tms_CN005177", nCS_Information) 'ระงับ(เกินวงเงิน)
'                                            Case "R", "3"
'                                                Call SP_MSGnShowing("tms_CN005176", nCS_Information) 'ลาออก
'                                        End Select
'                                        bW_FoundEmpCard = False
'                                        otbEmpCard.Text = ""                      '*AO51-03-20    Comm.Sheet  SF-108
'                                        orbMonth.Text = ""
'                                        orbYear.Text = ""
'                                        olaName.Caption = ""
'                                        olaRate.Caption = ""
'                                        Call SP_OBJxSetFocus(otbEmpCard)
'                                        Call SP_OTBxSelText(otbEmpCard)
'                                    End If
'                                Else
''                                    MsgBox "Employee Card is not found in the database Member Server."
'                                    Call SP_MSGnShowing("tms_CN005241", nCS_Information)         '*Eaw 55-12-19 เปลี่ยนใช้ msg จาก TSyslang
'                                    bW_FoundEmpCard = False
'                                    otbEmpCard.Text = ""                      '*AO51-03-20    Comm.Sheet  SF-108
'                                    orbMonth.Text = ""
'                                    orbYear.Text = ""
'                                    olaName.Caption = ""
'                                    olaRate.Caption = ""
'                                    Call SP_OBJxSetFocus(otbEmpCard)
'                                    Call SP_OTBxSelText(otbEmpCard)
'                                End If
'                            Else
'                                Set ocnMember = Nothing
'                                bVB_ReConnectEmp = False
'                                bW_FoundEmpCard = True
'                            End If
'                            '----------------------------------------------------------------------------
'                End Select
'            Else
'                 '*KT 51-07-14 Comsheet ML-SF- 356
'                If Not W_PRCxFindEmpCardByID(tVB_DiscEmpCardID) Then
'                    Select Case nW_CaseErr
'                        Case 0
'                            'none
'                        Case 1
'                            'Call SP_MSGnShowing("tms_CN005186", nCS_Warning)
'                        Case 2
'                            'Call SP_MSGnShowing("tms_CN005160", nCS_Warning)
'                        Case 3
'                            Call SP_MSGnShowing("tms_CN005069", nCS_Warning)
'                        Case 4
'                           ' Call SP_MSGnShowing("tms_CN005069", nCS_Warning)
'                        Case 5
'                            'Limit =N ,Allow Bay = N
'                    End Select
'                    bW_FoundEmpCard = False
'                    otbEmpCard.Text = ""                      '*AO51-03-20    Comm.Sheet  SF-108
'                    orbMonth.Text = ""
'                    orbYear.Text = ""
'                    olaName.Caption = ""
'                    olaRate.Caption = ""
'                    Call SP_OBJxSetFocus(otbEmpCard)
'                    Call SP_OTBxSelText(otbEmpCard)
'                Else
'                   bW_FoundEmpCard = True
'                    '*KT 51-08-07 Communication Sheet ML-SF-436 Member Card Status.xls
'                    '----------------------------------------------------------------------------
'                    Select Case nW_CaseErr
'                        Case 0, 1, 2
'On Error GoTo ErrHandle2
'                            '*KT 52-03-16
'                            If ocnMember Is Nothing And bVB_ReConnectEmp Then
'                                    MsgBox "Reconnect Member Server"
'                                    Call SP_STAxMouse(Me, True)
'                                    tSplash = "Can not connect the database member, System is  reconnect database member again." ' tms_CN005208
'                                    '*KT 52-05-26 PH 1.8 ML-SF-075 ในกรณีที่ POS ไม่สามารถติดต่อ Member Server ได้ ขอให้ POS แสดง Warning Message เพียงครั้งเดียว
'                                    Call SP_CONbMemberServer(True, tSplash, False)
'                                    Call SP_STAxMouse(Me, False)
'                                    bVB_ReConnectMem = False
'                            End If
'                             If Not ocnMember Is Nothing And oDB.SQLbAvailableServer(oDB.OffServerName, oDB.OnComIP) Then
''                            If Not ocnMember Is Nothing And ocnAll.ConnectionString <> oDB.ocnOffLine.ConnectionString Then
'                                tSql = "SELECT Top 1  FTEcdCode,FTEcdSta "
'                                tSql = tSql & vbCrLf & "FROM TCNMEmpCard"
'                                tSql = tSql & vbCrLf & "WHERE FTEcdCode ='" & tVB_DiscEmpCardID & "'" '*Bump 55-12-19 ต้องเอาหมายเลขบัตรพนักงานที่ได้จาก digits ที่ 9 ถึง digits ที่ 14
''                                tSql = tSql & vbCrLf & "WHERE FTEcdCode ='" & Trim(otbEmpCard.Text) & "'"
'                                nErr = SP_SQLvExecute(tSql, orsMem, ocnMember, , , , True)
'                                Select Case nErr
'                                    Case -2147467259, 91, -2147217871, 3704 '*KT 54-04-07 ป้องกันโปรแกรมหลุด 3.0.56
'                                        '*KT 52-05-26 PH 1.8 ML-SF-075 ในกรณีที่ POS ไม่สามารถติดต่อ Member Server ได้ ขอให้ POS แสดง Warning Message เพียงครั้งเดียว
'                                        '*KT 52-05-26 ให้ทำการแจ้ง message เตือนเมื่อทำการ Connect Member Server
'                                        'Call SP_MSGnShowing("tms_CN005134", tCS_CNMsgErr)
'                                End Select
'                                If Not (orsMem.BOF Or orsMem.EOF) Then
'                                    tSta = SP_FEDtChkString(orsMem, "FTEcdSta")
'                                    If UCase(tSta) = "1" Or UCase(tSta) = "" Or UCase(tSta) = "A" Then
'                                        bW_FoundEmpCard = True
'                                    Else
'                                        bW_FoundEmpCard = False
'                                        Select Case UCase(tSta)
'                                            Case "P", "2"
'                                                Call SP_MSGnShowing("tms_CN005177", nCS_Information) 'ระงับ(เกินวงเงิน)
'                                            Case "R", "3"
'                                                Call SP_MSGnShowing("tms_CN005176", nCS_Information) 'ลาออก
'                                        End Select
'                                        bW_FoundEmpCard = False
'                                        otbEmpCard.Text = ""                      '*AO51-03-20    Comm.Sheet  SF-108
'                                        orbMonth.Text = ""
'                                        orbYear.Text = ""
'                                        olaName.Caption = ""
'                                        olaRate.Caption = ""
'                                        Call SP_OBJxSetFocus(otbEmpCard)
'                                        Call SP_OTBxSelText(otbEmpCard)
'                                    End If
'                                Else
'                                    MsgBox "Employee Card is not found in the database Member Server."
'                                    bW_FoundEmpCard = False
'                                    otbEmpCard.Text = ""                      '*AO51-03-20    Comm.Sheet  SF-108
'                                    orbMonth.Text = ""
'                                    orbYear.Text = ""
'                                    olaName.Caption = ""
'                                    olaRate.Caption = ""
'                                    Call SP_OBJxSetFocus(otbEmpCard)
'                                    Call SP_OTBxSelText(otbEmpCard)
'                                End If
'                            Else
'                                Set ocnMember = Nothing
'                                bVB_ReConnectEmp = False
'                                bW_FoundEmpCard = True
'                            End If
'                            '----------------------------------------------------------------------------
'                    End Select
'                End If
'            End If
'        End If
'    End If
'    Call SP_STAxMouse(Me, False)
'    Exit Sub
'ErrHandle:
'    Call SP_STAxMouse(Me, False)
'    Call SP_DATxRsNothing(orsMem)
'    Call SP_MSGxShowError(Err.Number, Err.Description)
'    bW_FoundEmpCard = False
'    Exit Sub
'ErrHandle2:
'    Select Case Err.Number
'        Case -2147467259, 91, -2147217871, 3704 '*KT 54-04-07 ป้องกันโปรแกรมหลุด 3.0.56
'            Set ocnMember = Nothing
'            bVB_ReConnectEmp = False
'        Case Else
'            Call SP_MSGxShowError(Err.Number, Err.Description)
'    End Select
'    Call SP_STAxMouse(Me, False)
'    bW_FoundEmpCard = True
'    Call SP_DATxRsNothing(orsMem)
'End Sub

Private Sub otbEmpCard_LostFocus()
    otbEmpCard.Text = SP_TXTtTrimRightLeft(otbEmpCard.Text)
    If Not bW_FoundEmpCard Then
        If otbEmpCard.Text <> "" Then       '*Ao 2013/01/29
            Call otbEmpCard_KeyPress(13)
        End If
    End If
End Sub
'_______________________________________________________
'*Ao 2013/01/28
Private Sub otmForm_Timer()
    nW_Timer = nW_Timer + 1
    bW_ChkESC = False                                       '*Ao 2014/08/07 CommSheet TK-ISS3000-544-PH2  V 4.2.140
    
    If bW_HandleESC Then
        If nW_Timer = nW_TimeESC Then
            Call W_PRCxCase("2013_1251")
        ElseIf nW_Timer < nW_TimeESC Then
            olaForm(8).Caption = nW_TimeESC - nW_Timer
        End If
    End If
End Sub

Private Sub otmEDC_Timer()
    If bW_ManaulMode = False Then
        If tW_Comport = "" Then
            otmEDC.Enabled = False
            bW_HandleESC = True
            Call Form_KeyDown(vbKeyEscape, 0)
        Else
            bW_ChkESC = True                                                                                            '*Ao 2014/08/07 CommSheet TK-ISS3000-544-PH2  V 4.2.140
            Call W_EDCxMagic3CreditCard(tW_Comport, 0, "EMP")
        End If
    End If
End Sub
'_______________________________________________________




