VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.1#0"; "Codejock.SuiteCtrls.v11.1.3.ocx"
Begin VB.Form wCNPayCreditAmt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Amount;klg_Title"
   ClientHeight    =   3210
   ClientLeft      =   7575
   ClientTop       =   8700
   ClientWidth     =   6525
   ControlBox      =   0   'False
   Icon            =   "wCNPayCreditAmt.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   6525
   Tag             =   "2;"
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   2535
      Index           =   2
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   -240
      Width           =   6525
      _Version        =   720897
      _ExtentX        =   11509
      _ExtentY        =   4471
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
      Begin VB.TextBox otbMaxCal 
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
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   480
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
         TabIndex        =   0
         Text            =   "0"
         Top             =   1680
         Width           =   3975
      End
      Begin VB.TextBox otbOnTopAmt 
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
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   1080
         Width           =   3975
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pay;klg_Pay"
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
         TabIndex        =   9
         Tag             =   "2;"
         Top             =   480
         Width           =   1485
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
         TabIndex        =   7
         Tag             =   "2;"
         Top             =   1680
         Width           =   2160
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "On Top;klg_OnTop"
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
         TabIndex        =   3
         Tag             =   "2;"
         Top             =   1080
         Width           =   2385
      End
   End
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   1200
      Index           =   1
      Left            =   0
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2010
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
         Left            =   3360
         TabIndex        =   6
         TabStop         =   0   'False
         Tag             =   "2;"
         Top             =   480
         Width           =   1455
      End
   End
End
Attribute VB_Name = "wCNPayCreditAmt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cW_MaxCal As Double
Dim cW_MaxDisc As Double
Dim cW_MaxPay As Double
Dim cW_BalPay As Double
Dim bW_NotAlw As Boolean
Dim oDisc As New cPSDisc
Private orsTier As ADODB.Recordset

Private Sub Form_Load()
    Call SP_STAxMouse(Me, True)
    Call SP_FRMxChildCenter(Me)
    Call SP_FRMxGetCaptionFrmDB2Tag(Me)
    Call SP_FRMxSetCapForm(Me)
    Call W_FRMxSetInitial
    Call W_GETbRateOntopByTier(tVB_CrdKey, tVB_CrdType)
    bVB_IsCrdCard = False
    cVB_CrdAmt = 0 '*Tao 58-12-14 เคลียร์ค่าการชำระเงิน
    Call SP_STAxMouse(Me, False)
End Sub

Private Sub W_FRMxSetInitial()
'-----------------------------------------------------------
'Call:
'Cmt: *KT 51-08-20 การการแสดงการเริ่มต้นที่ Load Form
'*KT 51-11-03
'-----------------------------------------------------------
'   NEW
'-----------------------------------------------------------
    bW_NotAlw = True
    cW_MaxCal = cVB_CNOntopMaxAlw
    cW_MaxDisc = cVB_CNOntopMaxDisc
    cW_MaxPay = cVB_CNOntopMaxCal   'ยอดจ่าย
    otbOnTopAmt.Text = SP_FMTcCurPoint(cW_MaxDisc, False, True)
    otbMaxCal.Text = SP_FMTcCurPoint(cW_MaxCal, False, True)
    
    otbNum.Text = SP_FMTcCurPoint((cW_MaxPay + ((cVB_Balance - cVB_RoundRcv + oDisc.cAmtAllMall) - cW_MaxCal)), True, True)  'SP_FMTcCurPoint((cW_MaxPay + ((cVB_Balance - cVB_RoundRcv + oDisc.cAmtAllMall) - cW_MaxCal)), False, True)
    cW_BalPay = CDbl(Val(Format(otbNum.Text)))
    'Lock contorl
    otbOnTopAmt.Locked = True
    otbOnTopAmt.BackColor = nVB_CNClrBLock
    otbMaxCal.Locked = True
    otbMaxCal.BackColor = nVB_CNClrBLock
    bW_NotAlw = False
    'Sel Text
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SP_OBJxSetNothing(Me)
End Sub

Private Sub ocmCancel_Click()
    bVB_IsCrdCard = False
    cVB_CrdAmt = 0
    Unload Me
End Sub

Private Sub ocmOK_Click()
    Dim cPay As Double
    Dim cOntop As Double
    Dim cx  As Double
'New
'-------------------------------------------------------------
    cPay = SP_VALcText2Double(otbNum.Text)
    cOntop = SP_VALcText2Double(otbOnTopAmt.Text)
    cx = SP_VALcText2Double(otbMaxCal.Text)

    If CDbl(SP_FMTcCurPoint((cPay + cOntop), True, False)) > CDbl(SP_FMTcCurPoint(Val(Format(wSub.olaBalance.Caption)) + oDisc.cAmtAllMall, True, False)) Then   'check amount > balance  '*KT 52-02-23 640 or 800
         Call SP_MSGnShowing("tms_CN005066", nCS_Error)
         Call SP_OBJxSetFocus(otbNum)
         Exit Sub
     End If

    cVB_CrdAmt = cPay
    cVB_CNOntopMaxDisc = cOntop
    bVB_IsCrdCard = True
    Call SP_STAxMouse(Me, False)
    Unload Me
End Sub

Private Sub otbNum_Change() '*AO
    Dim cPay As Double
    Dim cInput  As Double
    Dim cDisc As Double
    
    'NEW
    '--------------------------------------------------------------------------------------
    
    otbNum.Text = SP_DEFtChangeNum(Trim(otbNum.Text))
    If bW_NotAlw Then Exit Sub
    otbNum.SelStart = Len(otbNum.Text)
    cInput = SP_VALcText2Double(otbNum.Text)
    If cInput > (cW_MaxPay) Then
        cInput = (cW_MaxPay)
    End If
    '***KT 51-11-26 V 1.0.1
    '------------------------------------------------------------------------------------------
    '***KT 51-11-22
   ' If cInput <= cW_MaxPay Then

   If cInput <= cW_MaxPay Then
        'ถ้า = ให้เท่ากับ cW_MaxDisc
         If cInput = cW_MaxPay Then
            otbOnTopAmt.Text = SP_FMTcCurPoint(cW_MaxDisc, False, True)
            otbMaxCal.Text = SP_FMTcCurPoint(cW_MaxCal, False, True)
        Else
            ''*KT 54-06-01 V.4.0.0 Req.19 OnTop Specify POSNo. *หา rate % ส่วนลดตาม Tier ใหม่
            If tVB_CrdStaTierOntopType = "1" Then '0 or ว่าง: Tier Amt 1: Tier %
                cVB_CrdRate = W_GETnRateOntopByTier(cInput)
                cPay = (cInput * 100) / (100 - cVB_CrdRate)
                otbMaxCal.Text = SP_FMTcCurPoint(cPay, False, True)
                otbOnTopAmt.Text = SP_FMTcCurPoint(cPay * cVB_CrdRate / 100, False, True)
            Else
                cDisc = W_GETcRateOntopByTierAmt(cInput)
                cPay = cInput + cDisc
                otbMaxCal.Text = SP_FMTcCurPoint(cPay, False, True)
                otbOnTopAmt.Text = SP_FMTcCurPoint(cDisc, False, True)
            End If
        End If
    Else
        otbOnTopAmt.Text = SP_FMTcCurPoint(cW_MaxDisc, False, True)
        otbNum.Text = SP_FMTcCurPoint((cInput + ((cVB_Balance - cVB_RoundRcv + oDisc.cAmtAllMall) - cW_MaxCal)), False, True)
        otbMaxCal.Text = SP_FMTcCurPoint(cW_MaxCal, False, True)
    End If
    '----------------------------------------------------------------------------------------------
End Sub

Private Sub otbNum_GotFocus()
    Call SP_OTBxSelText(otbNum)
End Sub

Private Sub otbNum_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDelete         'delete
            otbNum.Text = ""
    End Select
End Sub

Private Sub otbNum_KeyPress(KeyAscii As Integer)
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


Private Function W_GETbRateOntopByTier(ByVal ptTdmCode As String, ByVal ptCardType As String) As Boolean
'CREATE BY            # Kitti                              LAST MODIFIED BY           # Kitti
'CREATE DATE      #  54-06-03                            LAST MODIFIED DATE     # 54-06-03
'DESCRIPTION       #
'ARGUMENT LIST #
'RETURN VALUE  #
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim tSql As String
    
On Error GoTo ErrHandle
    W_GETbRateOntopByTier = False
    tSql = "SELECT * "
    tSql = tSql & vbCrLf & "FROM TCNMCreditCardTierOnTop"
    tSql = tSql & vbCrLf & "WHERE FTTdmCode='" & ptTdmCode & "'"
    tSql = tSql & vbCrLf & "AND FTCdcCreditType='" & ptCardType & "'"
    tSql = tSql & vbCrLf & "ORDER BY FTCtoSeq DESC"
    Call SP_SQLvExecute(tSql, orsTier, oDB.ocnOnLine, , , , , True, True)
    W_GETbRateOntopByTier = True
    Exit Function
ErrHandle:
    W_GETbRateOntopByTier = False
    Call SP_DATxRsClose(orsTier)
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Function


Private Function W_GETnRateOntopByTier(ByVal pcTotalAmt As Double) As Integer
'CREATE BY            # Kitti                              LAST MODIFIED BY           # Kitti
'CREATE DATE      #  54-06-03                            LAST MODIFIED DATE     # 54-06-03
'DESCRIPTION       #
'ARGUMENT LIST #
'RETURN VALUE  #
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim tSql As String
    
On Error GoTo ErrHandle
    W_GETnRateOntopByTier = 0
    If Not orsTier Is Nothing Then
        If orsTier.State = adStateOpen Then
            orsTier.Filter = "FTCtoType = '1'" 'Tier %
            orsTier.Sort = "FTCtoSeq DESC"
            If Not (orsTier.BOF And orsTier.EOF) Then
                Do While Not orsTier.EOF
                    If pcTotalAmt >= SP_FEDvChkNumeric(orsTier, "FCCtoAmt") Then
                        'W_GETnRateOntopByTier = SP_FEDvChkNumeric(orsTier, "FNCtoPerOff")
                        W_GETnRateOntopByTier = SP_FEDvChkNumeric(orsTier, "FCCtoPerOff") '*Bump 57-08-28 เปลี่ยน ชื่อฟิลด์จาก  FNCtoPerOff เป็น FCCtoPerOff  ให้ตรงตาม DB
                        Exit Do
                    End If
                    orsTier.MoveNext
                Loop
            End If
        End If
    End If
    orsTier.Filter = adFilterNone
    Exit Function
ErrHandle:
    W_GETnRateOntopByTier = 0
    If Not orsTier Is Nothing Then
        If orsTier.State = adStateOpen Then
              orsTier.Filter = adFilterNone
        End If
    End If
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Function


Private Function W_GETcRateOntopByTierAmt(ByVal pcTotalAmt As Double) As Double
'CREATE BY            # Kitti                              LAST MODIFIED BY           # Kitti
'CREATE DATE      #  54-06-03                            LAST MODIFIED DATE     # 54-06-03
'DESCRIPTION       #
'ARGUMENT LIST #
'RETURN VALUE  #
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim tSql As String
    Dim cCtoAmt As Double
On Error GoTo ErrHandle

    W_GETcRateOntopByTierAmt = 0
    If Not orsTier Is Nothing Then
        If orsTier.State = adStateOpen Then
            orsTier.Filter = "FTCtoType = '0'" 'Tier %
            orsTier.Sort = "FTCtoSeq DESC"
            If Not (orsTier.BOF And orsTier.EOF) Then
                Do While Not orsTier.EOF
                    cCtoAmt = SP_FEDvChkNumeric(orsTier, "FCCtoAmt")
                    If pcTotalAmt >= cCtoAmt Then
                        W_GETcRateOntopByTierAmt = W_GETcRateOntopByTierAmt + SP_FEDvChkNumeric(orsTier, "FCCtoAmtOff")
                        pcTotalAmt = pcTotalAmt - cCtoAmt
                        If tVB_CrdStaAlwTierAmtLast = "Y" Then 'Y : Last N : Loop
                            Exit Do
                        End If
                    Else
                        orsTier.MoveNext
                    End If
                Loop
            End If
        End If
    End If
    orsTier.Filter = adFilterNone
    Exit Function
ErrHandle:
    W_GETcRateOntopByTierAmt = 0
    If Not orsTier Is Nothing Then
        If orsTier.State = adStateOpen Then
              orsTier.Filter = adFilterNone
        End If
    End If
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Function

