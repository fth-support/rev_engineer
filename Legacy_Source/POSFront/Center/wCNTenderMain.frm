VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.1#0"; "Codejock.SuiteCtrls.Unicode.v11.1.3.ocx"
Begin VB.Form wTenderMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tender Amount;klg_Title"
   ClientHeight    =   5475
   ClientLeft      =   7305
   ClientTop       =   3075
   ClientWidth     =   7800
   ControlBox      =   0   'False
   Icon            =   "wCNTenderMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   7800
   Tag             =   "2;"
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   1320
      Index           =   2
      Left            =   0
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   4440
      Width           =   7815
      _Version        =   720897
      _ExtentX        =   13785
      _ExtentY        =   2328
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
      PaintManager.Position=   2
      Begin VB.CommandButton ocmOK 
         Caption         =   "&OK;klg_Ok"
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
         Left            =   4320
         TabIndex        =   3
         TabStop         =   0   'False
         Tag             =   "2;"
         Top             =   240
         Width           =   1455
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
         Left            =   5880
         TabIndex        =   5
         TabStop         =   0   'False
         Tag             =   "2;"
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton ocmDetail 
         Caption         =   "DENO;klg_Deno"
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
         Left            =   2760
         TabIndex        =   4
         TabStop         =   0   'False
         Tag             =   "2;"
         Top             =   240
         Width           =   1455
      End
   End
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   4800
      Index           =   0
      Left            =   0
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   -360
      Width           =   7815
      _Version        =   720897
      _ExtentX        =   13785
      _ExtentY        =   8467
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
      ItemCount       =   2
      SelectedItem    =   1
      Item(0).Caption =   "Item"
      Item(0).ControlCount=   7
      Item(0).Control(0)=   "otbCoin"
      Item(0).Control(1)=   "otbCurrency"
      Item(0).Control(2)=   "otbSumTender"
      Item(0).Control(3)=   "olaTotal"
      Item(0).Control(4)=   "Label1"
      Item(0).Control(5)=   "Label2"
      Item(0).Control(6)=   "Label3"
      Item(1).Caption =   "Item"
      Item(1).ControlCount=   13
      Item(1).Control(0)=   "Label4"
      Item(1).Control(1)=   "Label5"
      Item(1).Control(2)=   "Label6"
      Item(1).Control(3)=   "otbCoin2"
      Item(1).Control(4)=   "otbCurrency2"
      Item(1).Control(5)=   "otbForeign2"
      Item(1).Control(6)=   "otbSumTender2"
      Item(1).Control(7)=   "olaForeign"
      Item(1).Control(8)=   "olaGV"
      Item(1).Control(9)=   "otbGV2"
      Item(1).Control(10)=   "otbVVch2"
      Item(1).Control(11)=   "olaVVch"
      Item(1).Control(12)=   "olaSumTender"
      Begin VB.TextBox otbVVch2 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   2880
         MaxLength       =   11
         TabIndex        =   9
         Text            =   "0.00"
         Top             =   3120
         Visible         =   0   'False
         Width           =   4455
      End
      Begin VB.TextBox otbSumTender2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   675
         Left            =   2880
         Locked          =   -1  'True
         MaxLength       =   11
         TabIndex        =   11
         TabStop         =   0   'False
         Text            =   "0.00"
         Top             =   3840
         Width           =   4455
      End
      Begin VB.TextBox otbForeign2 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   2880
         MaxLength       =   11
         TabIndex        =   10
         Text            =   "0.00"
         Top             =   3120
         Width           =   4455
      End
      Begin VB.TextBox otbGV2 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   2880
         MaxLength       =   11
         TabIndex        =   8
         Text            =   "0.00"
         Top             =   2400
         Width           =   4455
      End
      Begin VB.TextBox otbCurrency2 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   2880
         MaxLength       =   11
         TabIndex        =   7
         Text            =   "0.00"
         Top             =   1680
         Width           =   4455
      End
      Begin VB.TextBox otbCoin2 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   2880
         MaxLength       =   11
         TabIndex        =   6
         Text            =   "0.00"
         Top             =   960
         Width           =   4455
      End
      Begin VB.TextBox otbCoin 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   -67000
         MaxLength       =   11
         TabIndex        =   0
         Text            =   "0.00"
         Top             =   840
         Visible         =   0   'False
         Width           =   4455
      End
      Begin VB.TextBox otbCurrency 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   -67000
         MaxLength       =   11
         TabIndex        =   1
         Text            =   "0.00"
         Top             =   1800
         Visible         =   0   'False
         Width           =   4455
      End
      Begin VB.TextBox otbSumTender 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   -67000
         Locked          =   -1  'True
         MaxLength       =   11
         TabIndex        =   2
         TabStop         =   0   'False
         Text            =   "0.00"
         Top             =   2760
         Visible         =   0   'False
         Width           =   4455
      End
      Begin VB.Label olaVVch 
         BackStyle       =   0  'Transparent
         Caption         =   "Vendor Voucher;klg_VVch"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   24
         Tag             =   "2;"
         Top             =   3240
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Label olaForeign 
         BackStyle       =   0  'Transparent
         Caption         =   "For;klg_Foreign"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   23
         Tag             =   "2;"
         Top             =   3240
         Width           =   2655
      End
      Begin VB.Label olaGV 
         BackStyle       =   0  'Transparent
         Caption         =   "Gift Voucher;klg_GV"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   22
         Tag             =   "2;"
         Top             =   2520
         Width           =   2655
      End
      Begin VB.Label olaSumTender 
         BackStyle       =   0  'Transparent
         Caption         =   "Total;klg_Tol"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   21
         Tag             =   "2;"
         Top             =   3960
         Width           =   2775
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Currency;klg_Cur"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   20
         Tag             =   "2;"
         Top             =   1800
         Width           =   2655
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Coin;klg_Coin"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   19
         Tag             =   "2;"
         Top             =   1080
         Width           =   2655
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Amount;klg_Amount"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4320
         TabIndex        =   18
         Tag             =   "2;"
         Top             =   360
         Width           =   3015
      End
      Begin VB.Label olaTotal 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Amount;klg_Amount"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -65560
         TabIndex        =   17
         Tag             =   "2;"
         Top             =   360
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Coin;klg_Coin"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -69160
         TabIndex        =   16
         Tag             =   "2;"
         Top             =   960
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Currency;klg_Cur"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -69160
         TabIndex        =   15
         Tag             =   "2;"
         Top             =   1920
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Total;klg_Tol"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -69160
         TabIndex        =   14
         Tag             =   "2;"
         Top             =   2880
         Visible         =   0   'False
         Width           =   2175
      End
   End
End
Attribute VB_Name = "wTenderMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const tW_StdHD = "Shd"
Const tW_StdDT = "Sdt"
'Dim bW_FirstLoad As Boolean

Private nW_DiscType As Integer
Private tW_KeyPress As String

Private Type TypeCurrency
    FTSkuCode As String
    FCSdtStkFac As Double
    FCSdtRegPrice As Double
End Type

Private Sub W_UPDxVirtualPosRC()
'---------------------------------------------------------------
'   Call:
'   Cmt: '*Eaw 56-10-01 เก็บค่า VirtualPos ให้ตาราง SaleRC ใน Re-Entry Mode
'---------------------------------------------------------------
    Dim tSql As String
    Dim tTerminal As String
    Dim tMsg As String
 
On Error GoTo HandleError
    Select Case SP_VALnText2Double(tVB_TransactionSub)
        Case tEN_TrnOverShortReportTaxInc, tEN_TrnOverShortReportManualTax
            tTerminal = ""
            If bVB_ReEntryMode Then
                tTerminal = tVB_CNVirTerminalNum
            End If
            tSql = "UPDATE " & tVB_CNTblRC
            tSql = tSql & vbCrLf & "SET FTRemark = '" & tTerminal & "'"
            tSql = tSql & " WHERE  FTTmnNum='" & tVB_CNTerminalNum & "' AND FTShdTransNo='" & tVB_Running & "'"
            tSql = tSql & " AND  FTShdTransType='" & tVB_Transaction & "' AND FDShdTransDate='" & tUT_SaleDate & "'"
            'Call SP_SQLvExecute(tSql) '*Tao 59-01-14 ของเดิม
            
            '*Tao 59-01-14 ถ้ามีข้อผิดพลาด ให้แจ้ง error และเก็บ log
            If SP_SQLvExecute(tSql) <> 0 Then
                tMsg = cUT.UT_STRtToken(tMS_CN502, ";", CInt(nVB_CNCutLang))
                Call SP_MSGnShowing(tMsg, nCS_Warning, True)
                Call SP_CRTxWriteLogTerminated("Form:" & Me.Name & " Events:W_UPDxVirtualPosR " & "(" & tSql & ")")
            End If
            '========================================
            
    End Select
HandleError:
    '*Tao 59-01-14  ถ้ามีข้อผิดพลาด ให้แจ้ง error และเก็บ log
    Call SP_MSGxShowError(Err.Number, Err.Description)
    Call SP_CRTxLogPrg("ErrorNumber:" & Err.Number & " ErrorDesc:" & Err.Description & " Form:" & Me.Name & " Events:W_UPDxVirtualPosRC")
    Call SP_CRTxWriteLogTerminated("ErrorNumber:" & Err.Number & " ErrorDesc:" & Err.Description & "Form:" & Me.Name & " Events:W_UPDxVirtualPosR")
End Sub

Private Sub W_UPDxVirtualPosDT()
'---------------------------------------------------------------
'   Call:
'   Cmt: '*Eaw 56-10-01 เก็บค่า VirtualPos ให้ตาราง SaleDT ใน Re-Entry Mode
'---------------------------------------------------------------
    Dim tSql As String
    Dim tTerminal As String
On Error GoTo HandleError

    Select Case SP_VALnText2Double(tVB_TransactionSub)
        Case tEN_TrnOverShortReportTaxInc, tEN_TrnOverShortReportManualTax  'tEN_TrnOverShortReport  Normal Mode ไม่เก็บ
            tTerminal = ""
            If bVB_ReEntryMode Then
                tTerminal = tVB_CNVirTerminalNum
            Else
                tTerminal = tVB_CNTerminalNum
            End If
            tSql = "UPDATE " & tVB_CNTblDT
            tSql = tSql & vbCrLf & "SET FTWahCode = '" & tTerminal & "'"
            tSql = tSql & " WHERE  FTTmnNum='" & tVB_CNTerminalNum & "' AND FTShdTransNo='" & tVB_Running & "'"
            tSql = tSql & " AND  FTShdTransType='" & tVB_Transaction & "' AND FDShdTransDate='" & tUT_SaleDate & "'"
            Call SP_SQLvExecute(tSql)
    End Select
HandleError:

End Sub

Private Function W_CHKbOverShortReport(ByVal pcSumEnter As Double) As Boolean
'---------------------------------------------------------------
'   Call:
'   Cmt: '*Ao 2013/07/31
'---------------------------------------------------------------
Dim tSql As String, tTenderType As String, tTmp As String, tTransWhere As String
Dim orsTemp As ADODB.Recordset
Dim cTenderALL As Double, cTender As Double, cChg As Double, cGVch As Double, cCurrency As Double, cACTUAL As Double, cTmp As Double, cVVch As Double '*Tao 62-09-05 เพิ่ม cVVch
Dim bChk As Boolean
Dim cCashTransfer As Double, CashTransfer As Double, cCashTrans As Double, cGVchTrans As Double, cCurrencyTrans As Double, cVVchTrans As Double '*Tao 62-09-05 เพิ่ม cVVchTrans
Dim cCHK  As Double, cCurrencyTmp As Double, cCurrencyTransTmp As Double
Dim aCurrency() As TypeCurrency
Dim nI  As Integer, nLoop As Integer
Dim tSKUCode As String

On Error GoTo HandleError
    
    W_CHKbOverShortReport = False
    
    tSql = wFunctionKB.W_GETtCloseTenderSQLEx(tVB_CNTblHD, tVB_CNTblRC, "FTShdTransType", True)
    Call SP_SQLvExecute(tSql, orsTemp)
    
    cTender = 0
    If (orsTemp.BOF And orsTemp.EOF) Then
        cTender = 0
    Else
        orsTemp.Filter = "FTTdmCode = 'T001'"
        If Not (orsTemp.BOF And orsTemp.EOF) Then
            orsTemp.MoveFirst
            Do While Not orsTemp.EOF
                cTender = cTender + CDbl(SP_FEDvChkNumeric(orsTemp, "FCSrcNet"))
                orsTemp.MoveNext
            Loop
        End If
    End If
    
    orsTemp.Filter = adFilterNone
    
    cChg = 0
    If (orsTemp.BOF And orsTemp.EOF) Then
        cChg = 0
    Else
        orsTemp.MoveFirst
        Do While Not (orsTemp.EOF)
            tTenderType = SP_FEDtChkString(orsTemp, "FTTdmCode")
            Select Case UCase(Trim(tTenderType))
                Case "T001", "T004", "T005", "T006", "T007", "T008", "T009", "T010", "T011", "T012", "T014"
                    cChg = cChg + CDbl(SP_FEDvChkNumeric(orsTemp, "FCSrcChg"))
            End Select
            orsTemp.MoveNext
        Loop
    End If
    cTenderALL = cTender - cChg
    
    orsTemp.Filter = adFilterNone
    
    cGVch = 0
    If (orsTemp.BOF And orsTemp.EOF) Then
        cGVch = 0
    Else
        orsTemp.Filter = "FTTdmCode = 'T005'"
        If Not (orsTemp.BOF And orsTemp.EOF) Then
            orsTemp.MoveFirst
            Do While Not orsTemp.EOF
                If SP_FEDtChkString(orsTemp, "FTTdmCode") = "T018" And (SP_FEDtChkString(orsTemp, "FTShdTransType") = "04" Or SP_FEDtChkString(orsTemp, "FTShdTransType") = "39" Or SP_FEDtChkString(orsTemp, "FTShdTransType") = "44" Or SP_FEDtChkString(orsTemp, "FTShdTransType") = "14") Then
                    cGVch = cGVch + 0
                Else
                    cGVch = cGVch + CDbl(SP_FEDvChkNumeric(orsTemp, "FCSrcNet"))
                End If
                orsTemp.MoveNext
            Loop
        End If
    End If
    
    orsTemp.Filter = adFilterNone
    
    '*Tao 62-09-05 เพิ่ม cVVch
    cVVch = 0
    If (orsTemp.BOF And orsTemp.EOF) Then
        cVVch = 0
    Else
        orsTemp.Filter = "FTTdmCode = 'T006'"
        If Not (orsTemp.BOF And orsTemp.EOF) Then
            orsTemp.MoveFirst
            Do While Not orsTemp.EOF
                cVVch = cVVch + CDbl(SP_FEDvChkNumeric(orsTemp, "FCSrcNet"))
                orsTemp.MoveNext
            Loop
        End If
    End If

    orsTemp.Filter = adFilterNone
    
    Call SP_DATxRsNothing(orsTemp)
    
    tSql = wFunctionKB.W_GETtCloseTenderCurSQL(tVB_CNTblHD, tVB_CNTblRC, "FTShdTransType", True)
    Call SP_SQLvExecute(tSql, orsTemp)
    
    ReDim aCurrency(0)        '*Ao 2013/10/11
    cCurrency = 0
    If (orsTemp.BOF And orsTemp.EOF) Then
        cCurrency = 0
    Else
        orsTemp.MoveFirst
        Do While Not orsTemp.EOF
'            cCurrency = cCurrency + CDbl(SP_FEDvChkNumeric(orsTemp, "FCSrcNet"))        '*Ao 2013/10/09   การปัดเศษทำให้มีเศษเหลือ
'            cCurrency = cCurrency + SP_FEDvChkNumeric(orsTemp, "FCSrcRteAmt")
'            cCurrencyTmp = SP_FMTcCurPoint(SP_FEDvChkNumeric(orsTemp, "FCSrcRteFac") * SP_FEDvChkNumeric(orsTemp, "FCSrcRteAmt"), True, False)
'            cCurrency = cCurrency + cCurrencyTmp
            '--------------------------------------------------------------------------------------'*Ao 2013/10/11
            ReDim Preserve aCurrency(UBound(aCurrency) + 1)
            aCurrency(UBound(aCurrency)).FTSkuCode = Trim(SP_FEDtChkString(orsTemp, "FTRteCode"))
            aCurrency(UBound(aCurrency)).FCSdtStkFac = SP_FEDvChkNumeric(orsTemp, "FCSrcRteFac")
            aCurrency(UBound(aCurrency)).FCSdtRegPrice = SP_FEDvChkNumeric(orsTemp, "FCSrcRteAmt")
            '--------------------------------------------------------------------------------------
            orsTemp.MoveNext
        Loop
    End If
    '----------------------------------------------------------------------------
    '*Ao 2013/09/23     CommSheet TK-ISS3000-305-UAT_00.00.03 FSBT NG
    
    tTransWhere = tVB_CNTblHD & ".FDShdTransDate = '" & tUT_SaleDate & "' AND " & tVB_CNTblHD & ".FTTmnNum = '" & tVB_CNTerminalNum & "'"
    tTransWhere = tTransWhere & vbCrLf & "AND " & tVB_CNTblHD & ".FTShdTransType = '12'"
    tTransWhere = tTransWhere & vbCrLf & "GROUP BY FTSkuCode,FTSkuAbbName,FTSkuAbbNameSnd,FTSdtStaPdt,FCSdtStkFac"
    
    tSql = "SELECT FTSkuCode,FTSkuAbbName,FTSkuAbbNameSnd,FCSdtStkFac,SUM(FCSdtSaleAmt) AS FCSum,SUM(FCSdtRegPrice) AS FCSum2,FTSdtStaPdt"
    tSql = tSql & vbCrLf & "FROM " & tVB_CNTblHD & " INNER JOIN"
    tSql = tSql & vbCrLf & tVB_CNTblDT & " ON " & tVB_CNTblHD & ".FTTmnNum = " & tVB_CNTblDT & ".FTTmnNum AND " & tVB_CNTblHD & ".FTShdTransNo = " & tVB_CNTblDT & ".FTShdTransNo AND"
    tSql = tSql & vbCrLf & tVB_CNTblHD & ".FTShdTransType = " & tVB_CNTblDT & ".FTShdTransType And " & tVB_CNTblHD & ".FDShdTransDate = " & tVB_CNTblDT & ".FDShdTransDate"
    tSql = tSql & vbCrLf & "WHERE " & tTransWhere
    Call SP_SQLvExecute(tSql, orsTemp)
    
    cCashTransfer = 0
    cCashTrans = 0
    cGVchTrans = 0
    cCurrencyTrans = 0
    cVVchTrans = 0 '*Tao 62-09-05 เพิ่ม cVVchTrans
    
    If Not (orsTemp.BOF And orsTemp.EOF) Then
        orsTemp.Filter = "FTSdtStaPdt = '0' OR FTSdtStaPdt = '1'"
        If Not (orsTemp.BOF And orsTemp.EOF) Then
            orsTemp.MoveFirst
            Do While Not orsTemp.EOF
                cCashTrans = cCashTrans + CDbl(SP_FEDvChkNumeric(orsTemp, "FCSum"))
                orsTemp.MoveNext
            Loop
        Else
            cCashTrans = 0
        End If
        
        orsTemp.Filter = adFilterNone
        orsTemp.Filter = "FTSdtStaPdt = '2'"
        If Not (orsTemp.BOF And orsTemp.EOF) Then
            orsTemp.MoveFirst
            Do While Not orsTemp.EOF
                cGVchTrans = cGVchTrans + CDbl(SP_FEDvChkNumeric(orsTemp, "FCSum"))
                orsTemp.MoveNext
            Loop
        Else
            cGVchTrans = 0
        End If
        
        '*Tao 62-09-05 เพิ่ม cVVchTrans
        orsTemp.Filter = adFilterNone
        orsTemp.Filter = "FTSdtStaPdt = '6'"
        If Not (orsTemp.BOF And orsTemp.EOF) Then
            orsTemp.MoveFirst
            Do While Not orsTemp.EOF
                cVVchTrans = cVVchTrans + CDbl(SP_FEDvChkNumeric(orsTemp, "FCSum"))
                orsTemp.MoveNext
            Loop
        Else
            cVVchTrans = 0
        End If
        
        orsTemp.Filter = adFilterNone
        orsTemp.Filter = "FTSdtStaPdt = '3'"
        If Not (orsTemp.BOF And orsTemp.EOF) Then
            orsTemp.MoveFirst
            Do While Not orsTemp.EOF
'                cCurrencyTrans = cCurrencyTrans + CDbl(SP_FEDvChkNumeric(orsTemp, "FCSum"))      '*Ao 2013/10/09   การปัดเศษทำให้มีเศษเหลือ
'                cCurrencyTransTmp = SP_FMTcCurPoint(SP_FEDvChkNumeric(orsTemp, "FCSdtStkFac") * SP_FEDvChkNumeric(orsTemp, "FCSum2"), True, False)
'                cCurrencyTrans = cCurrencyTrans + cCurrencyTransTmp
'                orsTemp.MoveNext
                '----------------------------------------------------------------------------------------------'*Ao 2013/10/11
                tSKUCode = Trim(SP_FEDtChkString(orsTemp, "FTSkuCode"))
                nI = -1
                For nLoop = 0 To UBound(aCurrency)
                    If tSKUCode = aCurrency(nLoop).FTSkuCode Then
                        nI = nLoop
                        Exit For
                    Else
                        nI = -1
                    End If
                Next nLoop
                If nI = -1 Then
                    cCurrencyTransTmp = SP_FMTcCurPoint(SP_FEDvChkNumeric(orsTemp, "FCSdtStkFac") * SP_FEDvChkNumeric(orsTemp, "FCSum2"), True, False)
                    cCurrencyTrans = cCurrencyTrans + cCurrencyTransTmp
                Else
                    cCurrencyTmp = SP_FMTcCurPoint((aCurrency(nLoop).FCSdtRegPrice - SP_FEDvChkNumeric(orsTemp, "FCSum2")) * SP_FEDvChkNumeric(orsTemp, "FCSdtStkFac"), True, False)
                    cCurrency = cCurrency + cCurrencyTmp
                    aCurrency(nLoop).FCSdtRegPrice = 0
                    aCurrency(nLoop).FCSdtStkFac = 0
                End If
                orsTemp.MoveNext
                '----------------------------------------------------------------------------------------------
            Loop
        Else
            cCurrencyTrans = 0
        End If
'        '----------------------------------------------------------------------------------------------'*Ao 2013/10/11
'        For nLoop = 0 To UBound(aCurrency)
'            cCurrencyTmp = SP_FMTcCurPoint(aCurrency(nLoop).FCSdtRegPrice * aCurrency(nLoop).FCSdtStkFac, True, False)
'            cCurrency = cCurrency + cCurrencyTmp
'        Next nLoop
'        '----------------------------------------------------------------------------------------------
    End If
    '----------------------------------------------------------------------------------------------'*Ao 2013/10/11
    For nLoop = 0 To UBound(aCurrency)
        cCurrencyTmp = SP_FMTcCurPoint(aCurrency(nLoop).FCSdtRegPrice * aCurrency(nLoop).FCSdtStkFac, True, False)
        cCurrency = cCurrency + cCurrencyTmp
    Next nLoop
    '----------------------------------------------------------------------------------------------
        
    cCashTransfer = cCashTrans + cGVchTrans + cCurrencyTrans + cVVchTrans '*Tao 62-09-05 เพิ่ม cVVchTrans
    '----------------------------------------------------------------------------
    cACTUAL = cTenderALL + cGVch + cCurrency + cVVch '*Tao 62-09-05 เพิ่ม cVVch
    
    cCHK = cACTUAL - cCashTransfer
    
    If SP_FMTcCurPoint(pcSumEnter, True, False) = SP_FMTcCurPoint(cCHK, True, False) Then                 '*Ao 2013/09/23  If pcSumEnter = cACTUAL Then
        bChk = False
    Else
        If pcSumEnter > cCHK Then                           'If pcSumEnter > cACTUAL Then
            tTmp = "Amount Input is Excess."
        Else
            tTmp = "Amount Input is Short."
        End If
'        cTmp = SP_FMTcCurPoint(CDbl(pcSumEnter) - CDbl(cACTUAL), False, True)
        cTmp = SP_FMTcCurPoint(CDbl(pcSumEnter) - CDbl(cCHK), False, True)
        wOver.tW_Msg = tTmp
        wOver.cW_Amount = cTmp
        wOver.Show vbModal
        If bVB_OverShort Then
            bChk = False
        Else
            bChk = True
        End If
    End If
    W_CHKbOverShortReport = bChk
    
    Call SP_DATxRsNothing(orsTemp)
Exit Function

HandleError:
    Call SP_DATxRsNothing(orsTemp)
    W_CHKbOverShortReport = False
End Function

Private Sub Form_Load()
On Error GoTo ErrHandle
    Call SP_FRMxChildCenter(Me)
    Call SP_FRMxGetCaptionFrmDB2Tag(Me)
    Call SP_FRMxSetCapForm(Me)
    Call W_FRMxSetInitial
    If Not oOpos Is Nothing Then
        If bVB_OPOSCashDrawerEnable Then Call oOpos.KickDrawer
    Else
         Call cUT_Print.UT_KickDrawer(Me)
    End If
    Exit Sub
ErrHandle:
    Call SP_CRTxLogPrg("ErrorNumber:" & Err.Number & " ErrorDesc:" & Err.Description & " Form:" & Me.Name & " Control:Form" & " Events:Form_Load")
End Sub

Private Sub W_FRMxSetInitial()
    Select Case SP_VALnText2Double(tVB_TransactionSub)
        Case tEN_TrnStartBank
            otcForm1(0).Item(0).Selected = True
            otbSumTender.Text = "0.00"
            ocmCancel.Enabled = False
        Case tEN_TrnDepositMoney, tEN_TrnPickUp
            otcForm1(0).Item(1).Selected = True
            otbSumTender2.Text = "0.00"
            ocmCancel.Enabled = True
        Case tEN_TrnOverShortReport, tEN_TrnOverShortReportTaxInc, tEN_TrnOverShortReportManualTax
            ocmCancel.Enabled = True
            '*Tao 62-09-11 Show Vendor Voucher (VVch)
            olaVVch.Visible = True
            otbVVch2.Visible = True
            otbVVch2.TabStop = True
            olaForeign.Top = 3960
            otbForeign2.Top = 3840
            olaSumTender.Top = 4680
            otbSumTender2.Top = 4560
            otcForm1(0).Height = 5520
            otcForm1(2).Top = 5160
            Me.Height = 6615
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)      '*Ao 2013/05/02  CommSheet TK-ISS3000-155-IT2 FSBT Explained-Aoo.xlsx
On Error GoTo ErrHandle

    If Not oOpos Is Nothing Then
        If bVB_OPOSCashDrawerEnable Then
            If oOpos.DrawerOpen Then
                wMsgDrawer.Show vbModal
            End If
        End If
    End If
    
    Call SP_OBJxSetNothing(Me)
    Exit Sub
ErrHandle:
    Call SP_CRTxLogPrg("ErrorNumber:" & Err.Number & " ErrorDesc:" & Err.Description & " Form:" & Me.Name & " Control:Form" & " Events:Form_Unload")
End Sub

Private Sub ocmCancel_Click()
    Select Case SP_VALnText2Double(tVB_TransactionSub)
        Case tEN_TrnStartBank
            Call ocmOK_Click
        Case tEN_TrnDepositMoney, tEN_TrnPickUp
            Unload Me
            wCNBegin.Show vbModal
        Case tEN_TrnOverShortReport, tEN_TrnOverShortReportTaxInc, tEN_TrnOverShortReportManualTax
            bVB_CNOverShortSta = False
            Unload Me
    
    End Select
End Sub

Private Sub ocmDetail_Click()
    Select Case SP_VALnText2Double(tVB_TransactionSub)
        Case tEN_TrnStartBank
            bVB_CNStartBankDT = False
            wCNStartBankDT.Show vbModal
            If bVB_CNStartBankDT Then
                Unload Me
                Call wBackGrand.W_PRVxFormActivate
            Else
                Call SP_OBJxSetFocus(otbCoin)
            End If
        Case tEN_TrnDepositMoney, tEN_TrnPickUp, tEN_TrnOverShortReport, tEN_TrnOverShortReportTaxInc, tEN_TrnOverShortReportManualTax
            bVB_CNStartBankDT = False
            wTender.Show vbModal
            If bVB_CNStartBankDT Then
                Unload Me
                If SP_VALnText2Double(tVB_TransactionSub) <> tEN_TrnOverShortReport And SP_VALnText2Double(tVB_TransactionSub) <> tEN_TrnOverShortReportTaxInc And SP_VALnText2Double(tVB_TransactionSub) <> tEN_TrnOverShortReportManualTax Then
                    wCNBegin.Show vbModal
                End If
            Else
                Call SP_OBJxSetFocus(otbCoin2)
            End If
    End Select
End Sub

Private Sub ocmOK_Click()
On Error GoTo ErrHandle
'    Select Case SP_VALnText2Double(tVB_TransactionSub)
'        Case tEN_TrnOpen:
'        Case tEN_TrnStartBank
'            If Not (W_PRCbStartBank) Then Exit Sub
'        Case tEN_TrnPickUp: Call W_PRCxPickUp: Exit Sub
'        Case tEN_TrnDepositMoney: Call W_PRCxDepositMny: Exit Sub
'    End Select
'    Unload Me
'    Exit Sub
    Select Case SP_VALnText2Double(tVB_TransactionSub)
        Case tEN_TrnOpen:
                Unload Me
                Call wBackGrand.W_PRVxFormActivate '*Ao 2011/03/28 CommSheet ML-PRV-046- POS Front Terminate Reboot and Not responding.xls
        Case tEN_TrnStartBank
            If W_PRCbStartBank Then
                Unload Me
                Call wBackGrand.W_PRVxFormActivate '*Ao 2011/03/28 CommSheet ML-PRV-046- POS Front Terminate Reboot and Not responding.xls
            Else
                Exit Sub
            End If
        Case tEN_TrnDepositMoney
            Call W_PRCxDepositMny
        Case tEN_TrnPickUp
            Call W_PRCxPickUp
        Case tEN_TrnOverShortReport, tEN_TrnOverShortReportTaxInc, tEN_TrnOverShortReportManualTax
            '*BG 56-09-10  Comsheet-285
            If W_CHKbOverShortReport(otbSumTender2.Text) Then
                Exit Sub
            End If
            '****
            Call W_PRCxCashOver
            bVB_CNOverShortSta = True
    End Select
    Exit Sub
ErrHandle:
    
End Sub

Private Function W_PRCbStartBank() As Boolean
On Error GoTo ErrHandle '*KT 54-04-12 V.3.0.57 ป้องกันโปรแกรมหลุด
    W_PRCbStartBank = False
    '*KT 51-06-23 Comsheet ML-SF-361
'    If Not (SP_CHKbBanknoteAlw(SP_VALcText2Double(otbSumTender.Text))) Then Exit Function      '*AO51-03-26   Comm.Sheet  ML-SF-100
    Call SP_INSxTransHD(tVB_CNTblHD)
    Call W_INSxTranDT
    Call W_INSxTranRC
    Call W_UPDxTransHD
    Call SP_UPDxSetCompleteTerminalMtn
    'process ตรวจการเปิดลิ้นชัก
    tVB_Transaction = aVB_TransMode(2).ID
    tVB_TransactionSub = tVB_Transaction
    tVB_TransName = aVB_TransMode(2).Name
    tVB_TransNameABB = aVB_NameAbb(tEN_AbbSale - 1).UsrLangPri & ";" & aVB_NameAbb(tEN_AbbSale - 1).UsrLangSec
    tVB_CNTypePrinter = aVB_TransMode(2).PrintType
    W_PRCbStartBank = True  '*AO51-03-26   Comm.Sheet  ML-SF-100 '*KT 54-04-12 V.3.0.57 ป้องกันโปรแกรมหลุด
    Exit Function
ErrHandle:
    Call SP_CRTxLogPrg("ErrorNumber:" & Err.Number & " ErrorDesc:" & Err.Description & " Form:" & Me.Name & " Control:Form" & " Events:W_PRCbStartBank") '*KT 54-04-12 V.3.0.57 ป้องกันโปรแกรมหลุด
End Function

Private Sub W_PRCxDepositMny()
On Error GoTo ErrHandle '*KT 54-04-12 V.3.0.57 ป้องกันโปรแกรมหลุด
'    otbSumTender.Text = SP_CHKtCurPoint(otbSumTender.Text)          '*AO51-03-12
    If SP_VALcText2Double(otbSumTender2.Text) > 0 Then
        Call SP_INSxTransHD(tVB_CNTblHD)
        Select Case SP_VALnText2Double(tVB_TransactionSub)
            Case tEN_TrnDepositMoney, tEN_TrnPickUp
                wMain.olaDocNo.Caption = tVB_Running '*KT 52-02-23 640 or 800
        End Select
        Call W_INSxTranDT
        Call W_INSxTranRC
        Call W_UPDxTransHD
        Call SP_UPDxSetCompleteTerminalMtn
        Unload Me
        wCNBegin.Show vbModal
    Else
'        Call ocmCancel_Click           '*Ao 2013/05/02  CommSheet TK-ISS3000-155-IT2 FSBT Explained-Aoo.xlsx
        Call SP_MSGnShowing("tms_CN005242", nCS_Warning)
    End If
    Exit Sub
ErrHandle:
    Call SP_CRTxLogPrg("ErrorNumber:" & Err.Number & " ErrorDesc:" & Err.Description & " Form:" & Me.Name & " Control:Form" & " Events:W_PRCxPickUp") '*KT 54-04-12 V.3.0.57 ป้องกันโปรแกรมหลุด
End Sub

Private Sub W_PRCxPickUp()
On Error GoTo ErrHandle '*KT 54-04-12 V.3.0.57 ป้องกันโปรแกรมหลุด
'    otbSumTender.Text = SP_CHKtCurPoint(otbSumTender.Text)          '*AO51-03-12
    If SP_VALcText2Double(otbSumTender2.Text) > 0 Then
        Call SP_INSxTransHD(tVB_CNTblHD)
        Select Case SP_VALnText2Double(tVB_TransactionSub)
            Case tEN_TrnDepositMoney, tEN_TrnPickUp
                wMain.olaDocNo.Caption = tVB_Running '*KT 52-02-23 640 or 800
        End Select
        Call W_INSxTranDT
        Call W_INSxTranRC
        Call W_UPDxTransHD
        Call SP_UPDxSetCompleteTerminalMtn
        Unload Me
        wCNBegin.Show vbModal
    Else
        'Call ocmCancel_Click
        Call SP_MSGnShowing("tms_CN005242", nCS_Warning)
    End If
    Exit Sub
ErrHandle:
    Call SP_CRTxLogPrg("ErrorNumber:" & Err.Number & " ErrorDesc:" & Err.Description & " Form:" & Me.Name & " Control:Form" & " Events:W_PRCxPickUp") '*KT 54-04-12 V.3.0.57 ป้องกันโปรแกรมหลุด
End Sub

Private Sub W_PRCxCashOver()
On Error GoTo ErrHandle '*KT 54-04-12 V.3.0.57 ป้องกันโปรแกรมหลุด
'    otbSumTender.Text = SP_CHKtCurPoint(otbSumTender.Text)          '*AO51-03-12
    If SP_VALcText2Double(otbSumTender2.Text) >= 0 Then
        Call SP_INSxTransHD(tVB_CNTblHD)
        Select Case SP_VALnText2Double(tVB_TransactionSub)
            Case tEN_TrnDepositMoney, tEN_TrnPickUp, tEN_TrnOverShortReport, tEN_TrnOverShortReportTaxInc, tEN_TrnOverShortReportManualTax
                wMain.olaDocNo.Caption = tVB_Running '*KT 52-02-23 640 or 800
        End Select
        Call W_INSxTranDT
        Call W_INSxTranRC
        Call W_UPDxTransHD
        Call SP_UPDxSetCompleteTerminalMtn
        Unload Me
        Select Case SP_VALnText2Double(tVB_TransactionSub)
            Case tEN_TrnOverShortReport, tEN_TrnOverShortReportTaxInc, tEN_TrnOverShortReportManualTax
                'none
            Case Else
                wCNBegin.Show vbModal
        End Select
    Else
        Call ocmCancel_Click
    End If
    Exit Sub
ErrHandle:
    Call SP_CRTxLogPrg("ErrorNumber:" & Err.Number & " ErrorDesc:" & Err.Description & " Form:" & Me.Name & " Control:Form" & " Events:W_PRCxPickUp") '*KT 54-04-12 V.3.0.57 ป้องกันโปรแกรมหลุด
End Sub

Public Sub W_INSxTranRC()
'------------------------------------------------------------------
'Call:
'Cmt: *KT 51-08-18 ทำการบันทึกข้อมูล RC ของ Start Bank
'------------------------------------------------------------------
    Dim tSql$
    Dim cFAmt As Double
    Dim cVB_Pay As Double
    Dim cNet As Double
    Dim cVB_Change As Double
    Dim nI As Long
    Dim nSeqNo As Long
    Dim tMsg As String
    
On Error GoTo ErrHandle

    cVB_Change = 0
    Select Case SP_VALnText2Double(tVB_TransactionSub)
        Case tEN_TrnStartBank
            cFAmt = CDbl(Val(Format(otbSumTender.Text)))
        Case tEN_TrnDepositMoney, tEN_TrnPickUp, tEN_TrnOverShortReport
            cFAmt = CDbl(Val(Format(otbCoin2.Text))) + CDbl(Val(Format(otbCurrency2.Text)))
    End Select
    
    cVB_Pay = cFAmt
    cNet = cFAmt
    tVB_CrdCardNo = ""
    cVB_FCSrcRteAmt = 0
    cVB_FCSrcRteFac = 0
    cVB_RoundChg = 0
    tVB_CrdApprCode = ""
    nUT_SeqRC = 1
    tVB_FTRteCode = ""
    tVB_CrdExpM = "": tVB_CrdExpY = ""
'    tVB_TenderCode = aVB_TenderMnu(0).ID
'    tVB_TenderType = aVB_TenderMnu(0).Name '*KT 53-10-15
'    tVB_TenderName = aVB_TenderMnu(0).DescSale
    For nI = 0 To UBound(aVB_TenderMnu)                                                             '*Ao 2013/05/31
        If UCase(aVB_TenderMnu(nI).ID) = UCase("T001") Then
            nSeqNo = nI
            Exit For
        End If
    Next nI
    tVB_TenderCode = aVB_TenderMnu(nSeqNo).ID
    tVB_TenderType = aVB_TenderMnu(nSeqNo).Name
    tVB_TenderName = aVB_TenderMnu(nSeqNo).DescSale
    
    tSql = "DELETE " & tVB_CNTblRC
    tSql = tSql & vbCrLf & "WHERE FTTmnNum = '" & tVB_CNTerminalNum & "'"
    tSql = tSql & vbCrLf & "AND FTShdTransNo = '" & tVB_Running & "'"
    tSql = tSql & vbCrLf & "AND FDShdTransDate = '" & tUT_SaleDate & "'"
    tSql = tSql & vbCrLf & "AND FTShdTransType = '" & tVB_Transaction & "'"
    'Call SP_SQLvExecute(tSql) '*Tao 59-01-14 ของเดิม
            
    '*Tao 59-01-14 ถ้ามีข้อผิดพลาด ให้แจ้ง error และเก็บ log
    If SP_SQLvExecute(tSql) <> 0 Then
        tMsg = cUT.UT_STRtToken(tMS_CN502, ";", CInt(nVB_CNCutLang))
        Call SP_MSGnShowing(tMsg, nCS_Warning, True)
        Call SP_CRTxWriteLogTerminated("Form:" & Me.Name & " Events:W_INSxTranRC " & "(" & tSql & ")")
    End If
    '========================================
    
    'save RC Cash
    tSql = "INSERT INTO " & tVB_CNTblRC
    tSql = tSql & vbCrLf & "(FTTmnNum,FTShdTransNo,FNSrcSeqNo,"
    tSql = tSql & vbCrLf & "FTShdTransType,FDShdTransDate,FTTdmCode,"
    tSql = tSql & vbCrLf & "FTTdmType,FTSrcCardNo,FTApprCode,"
    tSql = tSql & vbCrLf & "FCSrcCardChg,FCSrcFAmt,FCSrcAmt,"
    tSql = tSql & vbCrLf & "FCSrcNet,FCSrcChg,FTSrcStaPrc,"
    tSql = tSql & vbCrLf & "FTSrcRmk,FTSrcRetDocRef,FCSrcRetAmt,"
    tSql = tSql & vbCrLf & "FNShdStaSend,"
    tSql = tSql & vbCrLf & "FTRteCode,"
    tSql = tSql & vbCrLf & "FCSrcRteAmt ,"
    tSql = tSql & vbCrLf & "FCSrcRteFac,"
    tSql = tSql & vbCrLf & "FCSrcRndMnyChg,"
    tSql = tSql & vbCrLf & "FTSrcAbb,"
    tSql = tSql & vbCrLf & "FTSrcCardExp,"
    tSql = tSql & vbCrLf & "FCSrcOntopAmt,"
    tSql = tSql & vbCrLf & "FNSrcCrdPoint,"
    tSql = tSql & vbCrLf & "FTREFNO,"  '*Eaw 56-11-05
    tSql = tSql & vbCrLf & tCS_CNLastFedIns
    tSql = tSql & vbCrLf & " VALUES('" & tVB_CNTerminalNum & "','" & tVB_Running & "'," & nUT_SeqRC & ","
    tSql = tSql & vbCrLf & "'" & tVB_Transaction & "','" & tUT_SaleDate & "','" & tVB_TenderCode & "',"
    tSql = tSql & vbCrLf & "'" & tVB_TenderType & "','" & tVB_CrdCardNo & "','" & tVB_CrdApprCode & "',"
    tSql = tSql & vbCrLf & 0 & "," & SP_FMTcCurPoint(cFAmt, True, False) & "," & SP_FMTcCurPoint(cVB_Pay, True, False) & ","
    tSql = tSql & vbCrLf & SP_FMTcCurPoint(cNet, True, False) & "," & SP_FMTcCurPoint(cVB_Change, True, False) & ",'',"
    tSql = tSql & vbCrLf & "'',''," & 0 & ","
    tSql = tSql & vbCrLf & "0,"
    tSql = tSql & vbCrLf & "'" & tVB_FTRteCode & "',"
    tSql = tSql & vbCrLf & SP_FMTcCurPoint(cVB_FCSrcRteAmt, True, False) & ","
    tSql = tSql & vbCrLf & SP_FMTcCurPoint(cVB_FCSrcRteFac, True, False) & ","
    tSql = tSql & vbCrLf & SP_FMTcCurPoint(cVB_RoundChg, True, False) & ","
    'tSql = tSql & vbCrLf & "'" & tVB_TenderName & "',"
    tSql = tSql & vbCrLf & "'" & "Cash" & "',"
    If tVB_CrdExpM = "" And tVB_CrdExpY = "" Then
        tSql = tSql & "'" & "" & "',"
    Else
        tSql = tSql & "'" & tVB_CrdExpM & "/" & tVB_CrdExpY & "',"
    End If
    tSql = tSql & 0 & ","
    tSql = tSql & 0 & ","
    tSql = tSql & vbCrLf & "'" & tVB_TenderType & "',"  '*Eaw 56-11-05
    tSql = tSql & UT_tLastInsSql & ")"
    'Call SP_SQLvExecute(tSql) '*Tao 59-01-14 ของเดิม
            
    '*Tao 59-01-14 ถ้ามีข้อผิดพลาด ให้แจ้ง error และเก็บ log
    If SP_SQLvExecute(tSql) <> 0 Then
        tMsg = cUT.UT_STRtToken(tMS_CN502, ";", CInt(nVB_CNCutLang))
        Call SP_MSGnShowing(tMsg, nCS_Warning, True)
        Call SP_CRTxWriteLogTerminated("Form:" & Me.Name & " Events:W_INSxTranRC " & "(" & tSql & ")")
    End If
    '========================================
    
     'save RC Gift Voucher/Foreign Currency
    Select Case SP_VALnText2Double(tVB_TransactionSub)
        Case tEN_TrnDepositMoney, tEN_TrnPickUp, tEN_TrnOverShortReport
            'save Gift Voucher
            cVB_Change = 0
            cFAmt = CDbl(Val(Format(otbGV2.Text)))
            cVB_Pay = cFAmt
            cNet = cFAmt
            tVB_CrdCardNo = ""
            cVB_FCSrcRteAmt = 0
            cVB_FCSrcRteFac = 0
            cVB_RoundChg = 0
            tVB_CrdApprCode = ""
            nUT_SeqRC = 2
            tVB_FTRteCode = ""
            tVB_CrdExpM = "": tVB_CrdExpY = ""
'            tVB_TenderCode = aVB_TenderMnu(8).ID
'            tVB_TenderType = aVB_TenderMnu(8).Name '*KT 53-10-15
'            tVB_TenderName = aVB_TenderMnu(8).DescSale
            For nI = 0 To UBound(aVB_TenderMnu)
                If UCase(aVB_TenderMnu(nI).ID) = UCase("T005") Then
                    nSeqNo = nI
                    Exit For
                End If
            Next nI
            tVB_TenderCode = aVB_TenderMnu(nSeqNo).ID
            tVB_TenderType = aVB_TenderMnu(nSeqNo).Name
            tVB_TenderName = aVB_TenderMnu(nSeqNo).DescSale
            
            tSql = "INSERT INTO " & tVB_CNTblRC
            tSql = tSql & vbCrLf & "(FTTmnNum,FTShdTransNo,FNSrcSeqNo,"
            tSql = tSql & vbCrLf & "FTShdTransType,FDShdTransDate,FTTdmCode,"
            tSql = tSql & vbCrLf & "FTTdmType,FTSrcCardNo,FTApprCode,"
            tSql = tSql & vbCrLf & "FCSrcCardChg,FCSrcFAmt,FCSrcAmt,"
            tSql = tSql & vbCrLf & "FCSrcNet,FCSrcChg,FTSrcStaPrc,"
            tSql = tSql & vbCrLf & "FTSrcRmk,FTSrcRetDocRef,FCSrcRetAmt,"
            tSql = tSql & vbCrLf & "FNShdStaSend,"
            tSql = tSql & vbCrLf & "FTRteCode,"
            tSql = tSql & vbCrLf & "FCSrcRteAmt ,"
            tSql = tSql & vbCrLf & "FCSrcRteFac,"
            tSql = tSql & vbCrLf & "FCSrcRndMnyChg,"
            tSql = tSql & vbCrLf & "FTSrcAbb,"
            tSql = tSql & vbCrLf & "FTSrcCardExp,"
            tSql = tSql & vbCrLf & "FCSrcOntopAmt,"
            tSql = tSql & vbCrLf & "FNSrcCrdPoint,"
            tSql = tSql & vbCrLf & "FTREFNO,"  '*Eaw 56-11-05
            tSql = tSql & vbCrLf & tCS_CNLastFedIns
            tSql = tSql & vbCrLf & " VALUES('" & tVB_CNTerminalNum & "','" & tVB_Running & "'," & nUT_SeqRC & ","
            tSql = tSql & vbCrLf & "'" & tVB_Transaction & "','" & tUT_SaleDate & "','" & tVB_TenderCode & "',"
            tSql = tSql & vbCrLf & "'" & tVB_TenderType & "','" & tVB_CrdCardNo & "','" & tVB_CrdApprCode & "',"
            tSql = tSql & vbCrLf & 0 & "," & SP_FMTcCurPoint(cFAmt, True, False) & "," & SP_FMTcCurPoint(cVB_Pay, True, False) & ","
            tSql = tSql & vbCrLf & SP_FMTcCurPoint(cNet, True, False) & "," & SP_FMTcCurPoint(cVB_Change, True, False) & ",'',"
            tSql = tSql & vbCrLf & "'',''," & 0 & ","
            tSql = tSql & vbCrLf & "0,"
            tSql = tSql & vbCrLf & "'" & tVB_FTRteCode & "',"
            tSql = tSql & vbCrLf & SP_FMTcCurPoint(cVB_FCSrcRteAmt, True, False) & ","
            tSql = tSql & vbCrLf & SP_FMTcCurPoint(cVB_FCSrcRteFac, True, False) & ","
            tSql = tSql & vbCrLf & SP_FMTcCurPoint(cVB_RoundChg, True, False) & ","
            'tSql = tSql & vbCrLf & "'" & tVB_TenderName & "',"
            tSql = tSql & vbCrLf & "'" & "GVchr" & "',"
            If tVB_CrdExpM = "" And tVB_CrdExpY = "" Then
                tSql = tSql & "'" & "" & "',"
            Else
                tSql = tSql & "'" & tVB_CrdExpM & "/" & tVB_CrdExpY & "',"
            End If
            tSql = tSql & 0 & ","
            tSql = tSql & 0 & ","
            tSql = tSql & vbCrLf & "'" & tVB_TenderType & "',"  '*Eaw 56-11-05
            tSql = tSql & UT_tLastInsSql & ")"
            'Call SP_SQLvExecute(tSql) '*Tao 59-01-14 ของเดิม
                    
            '*Tao 59-01-14 ถ้ามีข้อผิดพลาด ให้แจ้ง error และเก็บ log
            If SP_SQLvExecute(tSql) <> 0 Then
                tMsg = cUT.UT_STRtToken(tMS_CN502, ";", CInt(nVB_CNCutLang))
                Call SP_MSGnShowing(tMsg, nCS_Warning, True)
                Call SP_CRTxWriteLogTerminated("Form:" & Me.Name & " Events:W_INSxTranRC " & "(" & tSql & ")")
            End If
            '========================================
            
            'Foreign
             cVB_Change = 0
            cFAmt = CDbl(Val(Format(otbForeign2.Text)))
            cVB_Pay = cFAmt
            cNet = cFAmt
            tVB_CrdCardNo = ""
            cVB_FCSrcRteAmt = 0
            cVB_FCSrcRteFac = 0
            cVB_RoundChg = 0
            tVB_CrdApprCode = ""
            nUT_SeqRC = 3
            tVB_FTRteCode = ""
            tVB_CrdExpM = "": tVB_CrdExpY = ""
'            tVB_TenderCode = aVB_TenderMnu(11).ID
'            tVB_TenderType = aVB_TenderMnu(11).Name '*KT 53-10-15
'            tVB_TenderName = aVB_TenderMnu(11).DescSale
            For nI = 0 To UBound(aVB_TenderMnu)
                If UCase(aVB_TenderMnu(nI).ID) = UCase("T012") Then
                    nSeqNo = nI
                    Exit For
                End If
            Next nI
            tVB_TenderCode = aVB_TenderMnu(nSeqNo).ID
            tVB_TenderType = aVB_TenderMnu(nSeqNo).Name
            tVB_TenderName = aVB_TenderMnu(nSeqNo).DescSale
            
            tSql = "INSERT INTO " & tVB_CNTblRC
            tSql = tSql & vbCrLf & "(FTTmnNum,FTShdTransNo,FNSrcSeqNo,"
            tSql = tSql & vbCrLf & "FTShdTransType,FDShdTransDate,FTTdmCode,"
            tSql = tSql & vbCrLf & "FTTdmType,FTSrcCardNo,FTApprCode,"
            tSql = tSql & vbCrLf & "FCSrcCardChg,FCSrcFAmt,FCSrcAmt,"
            tSql = tSql & vbCrLf & "FCSrcNet,FCSrcChg,FTSrcStaPrc,"
            tSql = tSql & vbCrLf & "FTSrcRmk,FTSrcRetDocRef,FCSrcRetAmt,"
            tSql = tSql & vbCrLf & "FNShdStaSend,"
            tSql = tSql & vbCrLf & "FTRteCode,"
            tSql = tSql & vbCrLf & "FCSrcRteAmt ,"
            tSql = tSql & vbCrLf & "FCSrcRteFac,"
            tSql = tSql & vbCrLf & "FCSrcRndMnyChg,"
            tSql = tSql & vbCrLf & "FTSrcAbb,"
            tSql = tSql & vbCrLf & "FTSrcCardExp,"
            tSql = tSql & vbCrLf & "FCSrcOntopAmt,"
            tSql = tSql & vbCrLf & "FNSrcCrdPoint,"
            tSql = tSql & vbCrLf & "FTREFNO,"  '*Eaw 56-11-05
            tSql = tSql & vbCrLf & tCS_CNLastFedIns
            tSql = tSql & vbCrLf & " VALUES('" & tVB_CNTerminalNum & "','" & tVB_Running & "'," & nUT_SeqRC & ","
            tSql = tSql & vbCrLf & "'" & tVB_Transaction & "','" & tUT_SaleDate & "','" & tVB_TenderCode & "',"
            tSql = tSql & vbCrLf & "'" & tVB_TenderType & "','" & tVB_CrdCardNo & "','" & tVB_CrdApprCode & "',"
            tSql = tSql & vbCrLf & 0 & "," & SP_FMTcCurPoint(cFAmt, True, False) & "," & SP_FMTcCurPoint(cVB_Pay, True, False) & ","
            tSql = tSql & vbCrLf & SP_FMTcCurPoint(cNet, True, False) & "," & SP_FMTcCurPoint(cVB_Change, True, False) & ",'',"
            tSql = tSql & vbCrLf & "'',''," & 0 & ","
            tSql = tSql & vbCrLf & "0,"
            tSql = tSql & vbCrLf & "'" & tVB_FTRteCode & "',"
            tSql = tSql & vbCrLf & SP_FMTcCurPoint(cVB_FCSrcRteAmt, True, False) & ","
            tSql = tSql & vbCrLf & SP_FMTcCurPoint(cVB_FCSrcRteFac, True, False) & ","
            tSql = tSql & vbCrLf & SP_FMTcCurPoint(cVB_RoundChg, True, False) & ","
            'tSql = tSql & vbCrLf & "'" & tVB_TenderName & "',"
'            tSql = tSql & vbCrLf & "'" & "Foreign" & "',"
            tSql = tSql & vbCrLf & "'" & "Foreign(Base)" & "',"         '*Ao 2013/05/31
            
            If tVB_CrdExpM = "" And tVB_CrdExpY = "" Then
                tSql = tSql & "'" & "" & "',"
            Else
                tSql = tSql & "'" & tVB_CrdExpM & "/" & tVB_CrdExpY & "',"
            End If
            tSql = tSql & 0 & ","
            tSql = tSql & 0 & ","
            tSql = tSql & vbCrLf & "'" & "Foreign(Local)" & "',"  '*Eaw 56-11-05
            tSql = tSql & UT_tLastInsSql & ")"
            'Call SP_SQLvExecute(tSql) '*Tao 59-01-14 ของเดิม
                    
            '*Tao 59-01-14 ถ้ามีข้อผิดพลาด ให้แจ้ง error และเก็บ log
            If SP_SQLvExecute(tSql) <> 0 Then
                tMsg = cUT.UT_STRtToken(tMS_CN502, ";", CInt(nVB_CNCutLang))
                Call SP_MSGnShowing(tMsg, nCS_Warning, True)
                Call SP_CRTxWriteLogTerminated("Form:" & Me.Name & " Events:W_INSxTranRC " & "(" & tSql & ")")
            End If
            '========================================
            
            tVB_FTRteCode = "" '*Eaw 56-11-05
            cVB_FCSrcRteAmt = 0         '*Ao 2013/11/05
            cVB_FCSrcRteFac = 0
            cVB_RoundChg = 0
            tVB_CrdExpM = ""
            tVB_CrdExpY = ""
    End Select
    Call W_UPDxVirtualPosRC '*Eaw 56-10-01
    Exit Sub
ErrHandle:
    '*Tao 59-01-14  ถ้ามีข้อผิดพลาด ให้แจ้ง error และเก็บ log
    Call SP_MSGxShowError(Err.Number, Err.Description)
    Call SP_CRTxLogPrg("ErrorNumber:" & Err.Number & " ErrorDesc:" & Err.Description & " Form:" & Me.Name & " Events:W_INSxTranRC")
    Call SP_CRTxWriteLogTerminated("ErrorNumber:" & Err.Number & " ErrorDesc:" & Err.Description & " Form:" & Me.Name & " Events:W_INSxTranRC")
End Sub

Private Sub W_UPDxTransHD()
'------------------------------------------------------
'Call:
'Cmt:   *KT 50-12-24 ปรับปรุงข้อมูล Transaction HD
'-----------------------------------------------------
    Dim tSql$
    tSql = "UPDATE " & tVB_CNTblHD
    Select Case SP_VALnText2Double(tVB_TransactionSub)
        Case tEN_TrnStartBank
            tSql = tSql & vbCrLf & " SET FCShdTotal =" & SP_FMTcCurPoint(SP_VALcText2Double(otbSumTender.Text), True, False)
            tSql = tSql & vbCrLf & ",FCShdGrand=" & SP_FMTcCurPoint(SP_VALcText2Double(otbSumTender.Text), True, False)
        Case tEN_TrnDepositMoney, tEN_TrnPickUp, tEN_TrnOverShortReport, tEN_TrnOverShortReportTaxInc, tEN_TrnOverShortReportManualTax
            tSql = tSql & vbCrLf & " SET FCShdTotal =" & SP_FMTcCurPoint(SP_VALcText2Double(otbSumTender2.Text), True, False)
            tSql = tSql & vbCrLf & ",FCShdGrand=" & SP_FMTcCurPoint(SP_VALcText2Double(otbSumTender2.Text), True, False)
    End Select
    tSql = tSql & vbCrLf & ",FTShdStaDoc  = '1'"
    tSql = tSql & vbCrLf & ",FTShdStaPaid='3'"               '3:จ่ายครบ
    tSql = tSql & vbCrLf & ",FTShdStaPrcDoc='1'"
    '*KT 52-06-25 เก็บ Flag ว่ามีการทำงานในสถานะ Online หรือ Offline (สถานะ Online/Offline   ว่าง:Online  1:Offline) ส่วนการ Connect Server
    If oDB.SQLbAvailableServer(oDB.OffServerName, oDB.OnComIP) Then
        tSql = tSql & ",FTShdStaType = '1'"         'Online
    Else
        tSql = tSql & ",FTShdStaType = '0'"         'Offline
    End If
    '*KT 51-09-22 เก็บ Flag ว่ามีการทำงานในสถานะ Online หรือ Offline (สถานะ Online/Offline   ว่าง:Online  1:Offline) ส่วนการใช้ DB
    tSql = tSql & ",FTShdStaPrcGL='1'"         'Offline
    tSql = tSql & ",FTCcyCode ='0'"  '1: แบบ Detail  0: แบบก้อน
    tSql = tSql & ",FTXihDCStatusMALL = '" & tVB_OverSh & "' "      '0 = Menu 17 ,1 = Menu 18     '*Ao 2014/07/31     CommSheet TK-ISS3000-481.0.0.1   V 4.1.129
    
    tSql = tSql & vbCrLf & " WHERE FTTmnNum='" & tVB_CNTerminalNum & "'"
    tSql = tSql & vbCrLf & " AND FTShdTransNo='" & tVB_Running & "'"
    tSql = tSql & vbCrLf & " AND FTShdTransType='" & tVB_Transaction & "'"
    tSql = tSql & vbCrLf & " AND FDShdTransDate ='" & tUT_SaleDate & "'"
    tSql = tSql & vbCrLf & " AND ISNULL(FTShdStaDoc,'0')  <> '1'"
    'Call DB_ExecuteSQL(tSql)
    Call SP_SQLvExecute(tSql)
    tVB_OverSh = "0"
    cVB_CoinsStartBank = SP_VALcText2Double(otbCoin.Text)
    cVB_CurrencyStartBank = SP_VALcText2Double(otbCurrency.Text)
    cVB_TotalStartBank = SP_VALcText2Double(otbSumTender.Text)
End Sub

Private Function W_INSxTranDT() As Boolean
'------------------------------------------------
'Call:
'Cmt: *Insert Data To DT
'-----------------------------------------------
    Dim nSeq%, cNet As Double
    Dim tSql$, tSkuBarC$, tSKUName$
    Dim tSKUName2$, tSKUNameSrvDoc$, tSKUNameSrvDoc2$
    Dim cVat As Double, cVatable As Double
    Dim cCostIn As Double, cQtyAll As Double
    Dim cQty As Double, cRegPrice As Double, cSalePrice As Double, cB4DisChg As Double, cSaleAmt As Double, cSaleUnit As Double, cSdtStkFac As Double
    Dim tWhr$
    Dim nEff As Long
    Dim nI As Integer, nR As Integer
    Dim tSdtStaPdt As Integer
    Dim nRow As Long
On Error GoTo ErrHandle

        W_INSxTranDT = False
        Select Case SP_VALnText2Double(tVB_TransactionSub)
            Case tEN_TrnStartBank
                For nI = 0 To 1
                      nR = nR + 1
                       If nI = 0 Then
                               tUT_SKUCode = "Coin"
                               tSKUName = "Coin"    'ชื่อสินค้า(Default ENG)1
                               tSKUName2 = "Coin"   'ชื่อสินค้า(Local)2
                               tSKUNameSrvDoc = "Coin"  'ชื่อสินค้า Service Doc(Default ENG) 1
                               tSKUNameSrvDoc2 = "Coin"    'ชื่อสินค้า Service Doc (Local)2
                                cQty = 1
                               cQtyAll = 1
                               cRegPrice = Val(Format(otbCoin.Text))
                               cSalePrice = Val(Format(otbCoin.Text))
                               cSaleUnit = Val(Format(otbCoin.Text))
                               cB4DisChg = Val(Format(otbCoin.Text))
                               cSdtStkFac = 1
                               cSaleAmt = Val(Format(otbCoin.Text))
                               tSdtStaPdt = "0"
                       Else
                                tUT_SKUCode = "Currency"
                               tSKUName = "Currency"  'ชื่อสินค้า(Default ENG)1
                               tSKUName2 = "Currency"   'ชื่อสินค้า(Local)2
                               tSKUNameSrvDoc = "Currency"   'ชื่อสินค้า Service Doc(Default ENG) 1
                               tSKUNameSrvDoc2 = "Currency"    'ชื่อสินค้า Service Doc (Local)2
                                cQty = 1
                               cQtyAll = 1
                               cRegPrice = Val(Format(otbCurrency.Text))
                               cSalePrice = Val(Format(otbCurrency.Text))
                               cSaleUnit = Val(Format(otbCurrency.Text))
                               cB4DisChg = Val(Format(otbCurrency.Text))
                               cSdtStkFac = 1
                               cSaleAmt = Val(Format(otbCurrency.Text))
                               tSdtStaPdt = "1"
                       End If
                       
                       tSql = "INSERT INTO " & tVB_CNTblDT
                       tSql = tSql & vbCrLf & "(FTStmCode,FTTmnNum,FTShdTransNo,FNSdtSeqNo,FTSkuCode"
                       tSql = tSql & vbCrLf & ",FTSkuAbbName,FTSkuAbbNameSnd,FTSkuAbbNameSrvDoc,FTSkuAbbNameSndSrvDoc"
                       tSql = tSql & vbCrLf & ",FTShdTransType,FDShdTransDate,FTShdTransTime,FCSdtQty,FCSdtQtyAll"
                       tSql = tSql & vbCrLf & ",FCSdtSalePrice,FCSdtSaleAmt,FCSdtRegPrice,FCSdtSaleUnit,FCSdtB4DisChg"
                       tSql = tSql & vbCrLf & ",FTSdtStaPdt,FCSdtStkFac,"
                       tSql = tSql & vbCrLf & tCS_CNLastFedIns
                       tSql = tSql & vbCrLf & " VALUES ("
                       tSql = tSql & vbCrLf & "'" & tVB_CNStoreCode & "','" & tVB_CNTerminalNum & "','" & tVB_Running & "'," & nR & ",'" & tUT_SKUCode & "'"
                       tSql = tSql & vbCrLf & ",'" & SP_STRtSiggleCode(tSKUName) & "'"
                       tSql = tSql & vbCrLf & "," & "'" & SP_STRtSiggleCode(tSKUName2) & "','" & SP_STRtSiggleCode(tSKUNameSrvDoc) & "','" & SP_STRtSiggleCode(tSKUNameSrvDoc2) & "'"
                       tSql = tSql & vbCrLf & ",'" & tVB_Transaction & "','" & tUT_SaleDate & "','" & Time$ & "'"
                       tSql = tSql & vbCrLf & "," & SP_FMTcCurPoint(cQty, True, False) & "," & SP_FMTcCurPoint(cQtyAll, True, False)
                       tSql = tSql & vbCrLf & "," & cRegPrice
                       tSql = tSql & vbCrLf & "," & cSaleAmt
                       tSql = tSql & vbCrLf & "," & cSalePrice
                       tSql = tSql & vbCrLf & "," & cSaleUnit
                       tSql = tSql & vbCrLf & "," & cB4DisChg
                       tSql = tSql & vbCrLf & ",'" & tSdtStaPdt & "'"
                       tSql = tSql & vbCrLf & "," & cSdtStkFac
                    
                       tSql = tSql & vbCrLf & "," & UT_tLastInsSql & ")"
                       If SP_SQLbExecute(tSql, , , nEff) Then
                           If nEff = 1 Then
                               W_INSxTranDT = True
                           End If
                       End If
                Next nI
            Case tEN_TrnDepositMoney, tEN_TrnPickUp
                For nI = 0 To 3
                       nR = nR + 1
                       Select Case nI
                            Case 0
                               tUT_SKUCode = "Coin"
                               tSKUName = "Coin"    'ชื่อสินค้า(Default ENG)1
                               tSKUName2 = "Coin"   'ชื่อสินค้า(Local)2
                               tSKUNameSrvDoc = "Coin"  'ชื่อสินค้า Service Doc(Default ENG) 1
                               tSKUNameSrvDoc2 = "Coin"    'ชื่อสินค้า Service Doc (Local)2
                                cQty = 1
                               cQtyAll = 1
                               cRegPrice = Val(Format(otbCoin2.Text))
                               cSalePrice = Val(Format(otbCoin2.Text))
                               cSaleUnit = Val(Format(otbCoin2.Text))
                               cB4DisChg = Val(Format(otbCoin2.Text))
                                If cSalePrice = 0 Then
                                    cQty = 0
                                    cQtyAll = 0
                               End If
                               cSdtStkFac = 1
                               cSaleAmt = Val(Format(otbCoin2.Text))
                               tSdtStaPdt = "0"
                            Case 1
                                tUT_SKUCode = "Currency"
                               tSKUName = "Currency"  'ชื่อสินค้า(Default ENG)1
                               tSKUName2 = "Currency"   'ชื่อสินค้า(Local)2
                               tSKUNameSrvDoc = "Currency"   'ชื่อสินค้า Service Doc(Default ENG) 1
                               tSKUNameSrvDoc2 = "Currency"    'ชื่อสินค้า Service Doc (Local)2
                                cQty = 1
                               cQtyAll = 1
                               cRegPrice = Val(Format(otbCurrency2.Text))
                               cSalePrice = Val(Format(otbCurrency2.Text))
                               cSaleUnit = Val(Format(otbCurrency2.Text))
                               cB4DisChg = Val(Format(otbCurrency2.Text))
                               If cSalePrice = 0 Then
                                    cQty = 0
                                    cQtyAll = 0
                               End If
                               cSdtStkFac = 1
                               cSaleAmt = Val(Format(otbCurrency2.Text))
                               tSdtStaPdt = "1"
                            Case 2
'                                tUT_SKUCode = "CreditVoucher"
'                                tSKUName = "Credit Voucher"  'ชื่อสินค้า(Default ENG)1
'                                tSKUName2 = "Credit Voucher"   'ชื่อสินค้า(Local)2
'                                tSKUNameSrvDoc = "Credit Voucher"   'ชื่อสินค้า Service Doc(Default ENG) 1   '*Ao 2013/05/30  CommSheet TK-ISS3000-208-ST 00 00 01-Ao.xls
'                                tSKUNameSrvDoc2 = "Credit Voucher"    'ชื่อสินค้า Service Doc (Local)2
                                tUT_SKUCode = "T005"
                                tSKUName = "Gift Voucher"  'ชื่อสินค้า(Default ENG)1
                                tSKUName2 = "Gift Voucher"   'ชื่อสินค้า(Local)2
                                tSKUNameSrvDoc = "Gift Voucher"   'ชื่อสินค้า Service Doc(Default ENG) 1
                                tSKUNameSrvDoc2 = "Gift Voucher"    'ชื่อสินค้า Service Doc (Local)2
                                cQty = 1
                               cQtyAll = 1
                               cRegPrice = Val(Format(otbGV2.Text))
                               cSalePrice = Val(Format(otbGV2.Text))
                               cSaleUnit = Val(Format(otbGV2.Text))
                               cB4DisChg = Val(Format(otbGV2.Text))
                               If cSalePrice = 0 Then
                                    cQty = 0
                                    cQtyAll = 0
                               End If
                               cSdtStkFac = 1
                               cSaleAmt = Val(Format(otbGV2.Text))
                               tSdtStaPdt = "2"
                            Case 3
                               tUT_SKUCode = "T012"
                               tSKUName = "Foreign(Base)"  'ชื่อสินค้า(Default ENG)1
                               tSKUName2 = "Foreign(Base)"   'ชื่อสินค้า(Local)2
                               tSKUNameSrvDoc = "Foreign(Base)"   'ชื่อสินค้า Service Doc(Default ENG) 1
                               tSKUNameSrvDoc2 = "Foreign(Base)"    'ชื่อสินค้า Service Doc (Local)2
                                cQty = 1
                               cQtyAll = 1
                               cRegPrice = Val(Format(otbForeign2.Text))
                               cSalePrice = Val(Format(otbForeign2.Text))
                               cSaleUnit = Val(Format(otbForeign2.Text))
                               cB4DisChg = Val(Format(otbForeign2.Text))
                               If cSalePrice = 0 Then
                                    cQty = 0
                                    cQtyAll = 0
                               End If
                               cSdtStkFac = 1
                               cSaleAmt = Val(Format(otbForeign2.Text))
                               tSdtStaPdt = "3"
                       End Select
                        
                        tSql = "INSERT INTO " & tVB_CNTblDT
                        tSql = tSql & vbCrLf & "(FTStmCode,FTTmnNum,FTShdTransNo,FNSdtSeqNo,FTSkuCode"
                        tSql = tSql & vbCrLf & ",FTSkuAbbName,FTSkuAbbNameSnd,FTSkuAbbNameSrvDoc,FTSkuAbbNameSndSrvDoc"
                        tSql = tSql & vbCrLf & ",FTShdTransType,FDShdTransDate,FTShdTransTime,FCSdtQty,FCSdtQtyAll"
                        tSql = tSql & vbCrLf & ",FCSdtSalePrice,FCSdtSaleAmt,FCSdtRegPrice,FCSdtSaleUnit,FCSdtB4DisChg"
                        tSql = tSql & vbCrLf & ",FTSdtStaPdt,FCSdtStkFac"
                        tSql = tSql & vbCrLf & ",FTSdtBarCode"                          '*Ao 2013/05/30  CommSheet TK-ISS3000-208-ST 00 00 01-Ao.xls
                        tSql = tSql & vbCrLf & "," & tCS_CNLastFedIns
                        
                        tSql = tSql & vbCrLf & " VALUES ("
                        tSql = tSql & vbCrLf & "'" & tVB_CNStoreCode & "','" & tVB_CNTerminalNum & "','" & tVB_Running & "'," & nR & ",'" & tUT_SKUCode & "'"
                        tSql = tSql & vbCrLf & ",'" & SP_STRtSiggleCode(tSKUName) & "'"
                        tSql = tSql & vbCrLf & "," & "'" & SP_STRtSiggleCode(tSKUName2) & "','" & SP_STRtSiggleCode(tSKUNameSrvDoc) & "','" & SP_STRtSiggleCode(tSKUNameSrvDoc2) & "'"
                        tSql = tSql & vbCrLf & ",'" & tVB_Transaction & "','" & tUT_SaleDate & "','" & Time$ & "'"
                        tSql = tSql & vbCrLf & "," & SP_FMTcCurPoint(cQty, True, False) & "," & SP_FMTcCurPoint(cQtyAll, True, False)
                        tSql = tSql & vbCrLf & "," & cRegPrice
                        tSql = tSql & vbCrLf & "," & cSaleAmt
                        tSql = tSql & vbCrLf & "," & cSalePrice
                        tSql = tSql & vbCrLf & "," & cSaleUnit
                        tSql = tSql & vbCrLf & "," & cB4DisChg
                        tSql = tSql & vbCrLf & ",'" & tSdtStaPdt & "'"
                        tSql = tSql & vbCrLf & "," & cSdtStkFac
                        tSql = tSql & vbCrLf & ",'" & tUT_SKUCode & "'"          '*Ao 2013/05/30  CommSheet TK-ISS3000-208-ST 00 00 01-Ao.xls
                        
                        tSql = tSql & vbCrLf & "," & UT_tLastInsSql & ")"
                        If SP_SQLbExecute(tSql, , , nEff) Then
                            If nEff = 1 Then
                                W_INSxTranDT = True
                            End If
                        End If
                Next nI
            Case tEN_TrnOverShortReport, tEN_TrnOverShortReportTaxInc, tEN_TrnOverShortReportManualTax
                For nI = 0 To 4 '*Tao 62-09-11 เพิ่ม Vendor Voucher (VVch)
                       nR = nR + 1
                       Select Case nI
                            Case 0
                               tUT_SKUCode = "Coin"
                               tSKUName = "Coin"    'ชื่อสินค้า(Default ENG)1
                               tSKUName2 = "Coin"   'ชื่อสินค้า(Local)2
                               tSKUNameSrvDoc = "Coin"  'ชื่อสินค้า Service Doc(Default ENG) 1
                               tSKUNameSrvDoc2 = "Coin"    'ชื่อสินค้า Service Doc (Local)2
                                cQty = 1
                               cQtyAll = 1
                               cRegPrice = Val(Format(otbCoin2.Text))
                               cSalePrice = Val(Format(otbCoin2.Text))
                               cSaleUnit = Val(Format(otbCoin2.Text))
                               cB4DisChg = Val(Format(otbCoin2.Text))
                                If cSalePrice = 0 Then
                                    cQty = 0
                                    cQtyAll = 0
                               End If
                               cSdtStkFac = 1
                               cSaleAmt = Val(Format(otbCoin2.Text))
                               tSdtStaPdt = "0"
                            Case 1
                                tUT_SKUCode = "Currency"
                               tSKUName = "Currency"  'ชื่อสินค้า(Default ENG)1
                               tSKUName2 = "Currency"   'ชื่อสินค้า(Local)2
                               tSKUNameSrvDoc = "Currency"   'ชื่อสินค้า Service Doc(Default ENG) 1
                               tSKUNameSrvDoc2 = "Currency"    'ชื่อสินค้า Service Doc (Local)2
                                cQty = 1
                               cQtyAll = 1
                               cRegPrice = Val(Format(otbCurrency2.Text))
                               cSalePrice = Val(Format(otbCurrency2.Text))
                               cSaleUnit = Val(Format(otbCurrency2.Text))
                               cB4DisChg = Val(Format(otbCurrency2.Text))
                               If cSalePrice = 0 Then
                                    cQty = 0
                                    cQtyAll = 0
                               End If
                               cSdtStkFac = 1
                               cSaleAmt = Val(Format(otbCurrency2.Text))
                               tSdtStaPdt = "1"
                            Case 2
'                                tUT_SKUCode = "CreditVoucher"
'                                tSKUName = "Credit Voucher"  'ชื่อสินค้า(Default ENG)1
'                                tSKUName2 = "Credit Voucher"   'ชื่อสินค้า(Local)2
'                                tSKUNameSrvDoc = "Credit Voucher"   'ชื่อสินค้า Service Doc(Default ENG) 1   '*Ao 2013/05/30  CommSheet TK-ISS3000-208-ST 00 00 01-Ao.xls
'                                tSKUNameSrvDoc2 = "Credit Voucher"    'ชื่อสินค้า Service Doc (Local)2
                                tUT_SKUCode = "T005"
                                tSKUName = "Gift Voucher"  'ชื่อสินค้า(Default ENG)1
                                tSKUName2 = "Gift Voucher"   'ชื่อสินค้า(Local)2
                                tSKUNameSrvDoc = "Gift Voucher"   'ชื่อสินค้า Service Doc(Default ENG) 1
                                tSKUNameSrvDoc2 = "Gift Voucher"    'ชื่อสินค้า Service Doc (Local)2
                                cQty = 1
                               cQtyAll = 1
                               cRegPrice = Val(Format(otbGV2.Text))
                               cSalePrice = Val(Format(otbGV2.Text))
                               cSaleUnit = Val(Format(otbGV2.Text))
                               cB4DisChg = Val(Format(otbGV2.Text))
                               If cSalePrice = 0 Then
                                    cQty = 0
                                    cQtyAll = 0
                               End If
                               cSdtStkFac = 1
                               cSaleAmt = Val(Format(otbGV2.Text))
                               tSdtStaPdt = "2"
                            Case 3 '*Tao 62-09-11 Enter Vendor Voucher
                               tUT_SKUCode = "T006"
                               tSKUName = "VVch"  'ชื่อสินค้า(Default ENG)1
                               tSKUName2 = "VVch"   'ชื่อสินค้า(Local)2
                               tSKUNameSrvDoc = tSKUName   'ชื่อสินค้า Service Doc(Default ENG) 1
                               tSKUNameSrvDoc2 = tSKUName2    'ชื่อสินค้า Service Doc (Local)2
                               cQty = 1
                               cQtyAll = 1
                               cRegPrice = Val(Format(otbVVch2.Text))
                               cSalePrice = Val(Format(otbVVch2.Text))
                               cSaleUnit = Val(Format(otbVVch2.Text))
                               cB4DisChg = Val(Format(otbVVch2.Text))
                               If cSalePrice = 0 Then
                                    cQty = 0
                                    cQtyAll = 0
                               End If
                               cSdtStkFac = 1
                               cSaleAmt = Val(Format(otbVVch2.Text))
                               tSdtStaPdt = "6"
                            Case 4
                               tUT_SKUCode = "T012"
                               tSKUName = "Foreign(Base)"  'ชื่อสินค้า(Default ENG)1
                               tSKUName2 = "Foreign(Base)"   'ชื่อสินค้า(Local)2
                               tSKUNameSrvDoc = "Foreign(Base)"   'ชื่อสินค้า Service Doc(Default ENG) 1
                               tSKUNameSrvDoc2 = "Foreign(Base)"    'ชื่อสินค้า Service Doc (Local)2
                                cQty = 1
                               cQtyAll = 1
                               cRegPrice = Val(Format(otbForeign2.Text))
                               cSalePrice = Val(Format(otbForeign2.Text))
                               cSaleUnit = Val(Format(otbForeign2.Text))
                               cB4DisChg = Val(Format(otbForeign2.Text))
                               If cSalePrice = 0 Then
                                    cQty = 0
                                    cQtyAll = 0
                               End If
                               cSdtStkFac = 1
                               cSaleAmt = Val(Format(otbForeign2.Text))
                               tSdtStaPdt = "3"
                       End Select
                       
                        tSql = "INSERT INTO " & tVB_CNTblDT
                        tSql = tSql & vbCrLf & "(FTStmCode,FTTmnNum,FTShdTransNo,FNSdtSeqNo,FTSkuCode"
                        tSql = tSql & vbCrLf & ",FTSkuAbbName,FTSkuAbbNameSnd,FTSkuAbbNameSrvDoc,FTSkuAbbNameSndSrvDoc"
                        tSql = tSql & vbCrLf & ",FTShdTransType,FDShdTransDate,FTShdTransTime,FCSdtQty,FCSdtQtyAll"
                        tSql = tSql & vbCrLf & ",FCSdtSalePrice,FCSdtSaleAmt,FCSdtRegPrice,FCSdtSaleUnit,FCSdtB4DisChg"
                        tSql = tSql & vbCrLf & ",FTSdtStaPdt,FCSdtStkFac"
                        tSql = tSql & vbCrLf & ",FTSdtBarCode"                          '*Ao 2013/05/30  CommSheet TK-ISS3000-208-ST 00 00 01-Ao.xls
                        tSql = tSql & vbCrLf & "," & tCS_CNLastFedIns
                        
                        tSql = tSql & vbCrLf & " VALUES ("
                        tSql = tSql & vbCrLf & "'" & tVB_CNStoreCode & "','" & tVB_CNTerminalNum & "','" & tVB_Running & "'," & nR & ",'" & tUT_SKUCode & "'"
                        tSql = tSql & vbCrLf & ",'" & SP_STRtSiggleCode(tSKUName) & "'"
                        tSql = tSql & vbCrLf & "," & "'" & SP_STRtSiggleCode(tSKUName2) & "','" & SP_STRtSiggleCode(tSKUNameSrvDoc) & "','" & SP_STRtSiggleCode(tSKUNameSrvDoc2) & "'"
                        tSql = tSql & vbCrLf & ",'" & tVB_Transaction & "','" & tUT_SaleDate & "','" & Time$ & "'"
                        tSql = tSql & vbCrLf & "," & SP_FMTcCurPoint(cQty, True, False) & "," & SP_FMTcCurPoint(cQtyAll, True, False)
                        tSql = tSql & vbCrLf & "," & cRegPrice
                        tSql = tSql & vbCrLf & "," & cSaleAmt
                        tSql = tSql & vbCrLf & "," & cSalePrice
                        tSql = tSql & vbCrLf & "," & cSaleUnit
                        tSql = tSql & vbCrLf & "," & cB4DisChg
                        tSql = tSql & vbCrLf & ",'" & tSdtStaPdt & "'"
                        tSql = tSql & vbCrLf & "," & cSdtStkFac
                        tSql = tSql & vbCrLf & ",'" & tUT_SKUCode & "'"          '*Ao 2013/05/30  CommSheet TK-ISS3000-208-ST 00 00 01-Ao.xls
                        
                        tSql = tSql & vbCrLf & "," & UT_tLastInsSql & ")"
                        If SP_SQLbExecute(tSql, , , nEff) Then
                            If nEff = 1 Then
                                W_INSxTranDT = True
                            End If
                        End If
                Next nI
        End Select
        Select Case SP_VALnText2Double(tVB_TransactionSub)
                Case tEN_TrnOverShortReport, tEN_TrnOverShortReportTaxInc, tEN_TrnOverShortReportManualTax
                      Call W_PRNbOverShort(nR + 1)
                      Call W_UPDxVirtualPosDT '*Eaw 56-10-01
        End Select
        Exit Function
ErrHandle:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Function

Private Sub W_PRNbOverShort(ByVal pnSeqNo As Long)

    Dim cACTUAL As Double
    Dim cQtyACTUAL As Double

    Dim nSeq%, cNet As Double
    Dim tSql$, tSkuBarC$, tSKUName$
    Dim tSKUName2$, tSKUNameSrvDoc$, tSKUNameSrvDoc2$
    Dim cVat As Double, cVatable As Double
    Dim cCostIn As Double, cQtyAll As Double
    Dim cQty As Double, cRegPrice As Double, cSalePrice As Double, cB4DisChg As Double, cSaleAmt As Double, cSaleUnit As Double, cSdtStkFac As Double
    Dim tWhr$
    Dim nEff As Long
    Dim nI As Integer, nR As Integer
    Dim tSdtStaPdt As Integer
    Dim nRow As Long

    Dim orsPickup As ADODB.Recordset
    Dim orsTemp As ADODB.Recordset
    Dim orsFor As ADODB.Recordset
    Dim orsCur  As ADODB.Recordset
    Dim tTransWhere As String
    Dim nError As Long
    Dim cReg As Double
    Dim cRate As Double
    Dim cACTUALTmp As Double
    Dim cChg As Double
    Dim tTenderType As String '*Eaw 56-06-24 CommSheet TK-ISS3000-FAPL-029.xls
    Dim tTransType As String  '*Eaw 56-06-28
    
On Error GoTo ErrHandler
    nR = pnSeqNo
    tSql = wFunctionKB.W_GETtCloseTenderSQLEx(tVB_CNTblHD, tVB_CNTblRC, "FTShdTransType", True)
    Call SP_SQLvExecute(tSql, orsTemp)

   'cash
    cACTUAL = 0
    cQtyACTUAL = 0
    cChg = 0   '*Eaw 56-06-24 CommSheet TK-ISS3000-FAPL-029.xls
    If Not (orsTemp.BOF And orsTemp.EOF) Then
        orsTemp.Filter = "FTTdmCode = 'T001'"
        If Not (orsTemp.BOF And orsTemp.EOF) Then
'            cACTUAL = SP_FEDvChkNumeric(orsTemp, "FCSrcNet")
'            cQtyACTUAL = SP_FEDvChkNumeric(orsTemp, "FCQty")
'            cChg = SP_FEDvChkNumeric(orsTemp, "FCSrcChg")                       '*Ao 2013/05/31 CommSheet TK-ISS3000-217-ST 00 00 01-Ao.xls
            orsTemp.MoveFirst
            Do While Not orsTemp.EOF
                 tTransType = SP_FEDtChkString(orsTemp, "FTShdTransType")
                cACTUAL = cACTUAL + SP_FEDvChkNumeric(orsTemp, "FCSrcNet")
'                Select Case UCase(Trim(tTransType))
''                    Case "03", "06", "07", "10", "11", "13", "19", "22"
'                    Case "03", "06", "07", "10", "11", "13", "19", "22", "38", "43", "04", "16", "17", "20" '*BG 56-09-04 Comsheet-285
'                        cQtyACTUAL = cQtyACTUAL + SP_FEDvChkNumeric(orsTemp, "FCQty")
'                End Select
                cQtyACTUAL = cQtyACTUAL + SP_FEDvChkNumeric(orsTemp, "FCQty")    '*Ao 2013/09/23 CommSheet TK-ISS3000-350-UAT-FAPL-118_00.00.03 FSBT NG
                orsTemp.MoveNext
            Loop
        End If
         orsTemp.Filter = adFilterNone
    End If
    If Not (orsTemp.BOF And orsTemp.EOF) Then  '*Eaw 56-06-24 CommSheet TK-ISS3000-FAPL-029.xls
        orsTemp.MoveFirst
        Do While Not (orsTemp.EOF)
            tTenderType = SP_FEDtChkString(orsTemp, "FTTdmCode")
            Select Case UCase(Trim(tTenderType))
'                Case  "T004", "T005", "T006", "T007", "T008", "T009", "T010", "T011", "T012", "T014"
                Case "T001", "T004", "T005", "T006", "T007", "T008", "T009", "T010", "T011", "T012", "T014"
                    cChg = cChg + SP_FEDvChkNumeric(orsTemp, "FCSrcChg")
            End Select
            orsTemp.MoveNext
        Loop
   End If
    cACTUAL = cACTUAL - cChg                                '*Ao 2013/05/31 CommSheet TK-ISS3000-217-ST 00 00 01-Ao.xls
    
    tUT_SKUCode = "Cash"
    tSKUName = "Cash"  'ชื่อสินค้า(Default ENG)1
    tSKUName2 = "Cash"   'ชื่อสินค้า(Local)2
    tSKUNameSrvDoc = "Cash"  'ชื่อสินค้า Service Doc(Default ENG) 1
    tSKUNameSrvDoc2 = "Cash"    'ชื่อสินค้า Service Doc (Local)2

    cQty = cQtyACTUAL
    cQtyAll = cQtyACTUAL

    cRegPrice = cACTUAL
    cSalePrice = cACTUAL
    cSaleUnit = cACTUAL
    cB4DisChg = cACTUAL
    cSdtStkFac = cACTUAL
    cSaleAmt = cACTUAL
    cSaleAmt = cACTUAL
    tSdtStaPdt = "4"

    tSql = "INSERT INTO " & tVB_CNTblDT
    tSql = tSql & vbCrLf & "(FTStmCode,FTTmnNum,FTShdTransNo,FNSdtSeqNo,FTSkuCode"
    tSql = tSql & vbCrLf & ",FTSkuAbbName,FTSkuAbbNameSnd,FTSkuAbbNameSrvDoc,FTSkuAbbNameSndSrvDoc"
    tSql = tSql & vbCrLf & ",FTShdTransType,FDShdTransDate,FTShdTransTime,FCSdtQty,FCSdtQtyAll"
    tSql = tSql & vbCrLf & ",FCSdtSalePrice,FCSdtSaleAmt,FCSdtRegPrice,FCSdtSaleUnit,FCSdtB4DisChg"
    tSql = tSql & vbCrLf & ",FTSdtStaPdt,FCSdtStkFac"
    tSql = tSql & vbCrLf & ",FTSdtBarCode"                      '*Ao 2013/05/30  CommSheet TK-ISS3000-208-ST 00 00 01-Ao.xls
    tSql = tSql & vbCrLf & "," & tCS_CNLastFedIns
    tSql = tSql & vbCrLf & " VALUES ("
    tSql = tSql & vbCrLf & "'" & tVB_CNStoreCode & "','" & tVB_CNTerminalNum & "','" & tVB_Running & "'," & nR & ",'" & tUT_SKUCode & "'"
    tSql = tSql & vbCrLf & ",'" & SP_STRtSiggleCode(tSKUName) & "'"
    tSql = tSql & vbCrLf & "," & "'" & SP_STRtSiggleCode(tSKUName2) & "','" & SP_STRtSiggleCode(tSKUNameSrvDoc) & "','" & SP_STRtSiggleCode(tSKUNameSrvDoc2) & "'"
    tSql = tSql & vbCrLf & ",'" & tVB_Transaction & "','" & tUT_SaleDate & "','" & Time$ & "'"
    tSql = tSql & vbCrLf & "," & SP_FMTcCurPoint(cQty, True, False) & "," & SP_FMTcCurPoint(cQtyAll, True, False)
    tSql = tSql & vbCrLf & "," & cSalePrice
    tSql = tSql & vbCrLf & "," & cSaleAmt
    tSql = tSql & vbCrLf & "," & cSalePrice
    tSql = tSql & vbCrLf & "," & cSaleUnit
    tSql = tSql & vbCrLf & "," & cB4DisChg
    tSql = tSql & vbCrLf & ",'" & tSdtStaPdt & "'"
    tSql = tSql & vbCrLf & "," & cSdtStkFac
    tSql = tSql & vbCrLf & ",'T001'"    '*Ao 2013/05/30  CommSheet TK-ISS3000-208-ST 00 00 01-Ao.xls
    
    tSql = tSql & vbCrLf & "," & UT_tLastInsSql & ")"
    If SP_SQLbExecute(tSql, , , nEff) Then
        If nEff = 1 Then
            'none
        End If
    End If

   'GVch
    nR = nR + 1
    cACTUAL = 0
    cQtyACTUAL = 0
    If Not (orsTemp.BOF And orsTemp.EOF) Then
'        orsTemp.Filter = "FTTdmCode = 'T004' OR FTTdmCode = 'T005'  OR FTTdmCode = 'T006' OR FTTdmCode = 'T007' OR FTTdmCode = 'T008' OR FTTdmCode = 'T009'  OR FTTdmCode = 'T010' OR FTTdmCode = 'T011' OR FTTdmCode = 'T018'"
'        orsTemp.Filter = "FTTdmCode = 'T004' OR FTTdmCode = 'T005'  OR FTTdmCode = 'T006' OR FTTdmCode = 'T007' OR FTTdmCode = 'T008' OR FTTdmCode = 'T009'  OR FTTdmCode = 'T010' OR FTTdmCode = 'T011'"  '*Eaw 56-06-25 ไมานับ Credit Voucher เป็น GVch
'        orsTemp.Filter = "FTTdmCode = 'T004' OR FTTdmCode = 'T005'  OR FTTdmCode = 'T006' OR FTTdmCode = 'T007' OR FTTdmCode = 'T008' OR FTTdmCode = 'T009'  OR FTTdmCode = 'T011'"   '*BG 56-08-07  คิดรวมแค่ GV อย่างเดียว" CommSheet TK-ISS3000-303-UAT_Question.xls
        orsTemp.Filter = "FTTdmCode = 'T005'"   '*BG 56-08-07  คิดรวมแค่ GV อย่างเดียว" CommSheet TK-ISS3000-303-UAT_Question.xls
        If Not (orsTemp.BOF And orsTemp.EOF) Then
            Do While Not orsTemp.EOF
                If SP_FEDtChkString(orsTemp, "FTTdmCode") = "T018" And (SP_FEDtChkString(orsTemp, "FTShdTransType") = "04" Or SP_FEDtChkString(orsTemp, "FTShdTransType") = "39" Or SP_FEDtChkString(orsTemp, "FTShdTransType") = "44" Or SP_FEDtChkString(orsTemp, "FTShdTransType") = "14") Then
                    cACTUAL = cACTUAL + 0
                    cQtyACTUAL = cQtyACTUAL + 0
                Else
                    cACTUAL = cACTUAL + SP_FEDvChkNumeric(orsTemp, "FCSrcNet")
                    cQtyACTUAL = cQtyACTUAL + SP_FEDvChkNumeric(orsTemp, "FCQty")
                End If
                orsTemp.MoveNext
            Loop
        End If
         orsTemp.Filter = adFilterNone
    End If

    tUT_SKUCode = "GVch"
    tSKUName = "GVch"  'ชื่อสินค้า(Default ENG)1               '*Ao 2013/05/30  CommSheet TK-ISS3000-208-ST 00 00 01-Ao.xls
    tSKUName2 = "GVch"   'ชื่อสินค้า(Local)2
    tSKUNameSrvDoc = "GVch"  'ชื่อสินค้า Service Doc(Default ENG) 1
    tSKUNameSrvDoc2 = "GVch"    'ชื่อสินค้า Service Doc (Local)2

    cQty = cQtyACTUAL
    cQtyAll = cQtyACTUAL

    cRegPrice = cACTUAL
    cSalePrice = cACTUAL
    cSaleUnit = cACTUAL
    cB4DisChg = cACTUAL
    cSdtStkFac = cACTUAL
    cSaleAmt = cACTUAL
    cSaleAmt = cACTUAL
    tSdtStaPdt = "4"

    tSql = "INSERT INTO " & tVB_CNTblDT
    tSql = tSql & vbCrLf & "(FTStmCode,FTTmnNum,FTShdTransNo,FNSdtSeqNo,FTSkuCode"
    tSql = tSql & vbCrLf & ",FTSkuAbbName,FTSkuAbbNameSnd,FTSkuAbbNameSrvDoc,FTSkuAbbNameSndSrvDoc"
    tSql = tSql & vbCrLf & ",FTShdTransType,FDShdTransDate,FTShdTransTime,FCSdtQty,FCSdtQtyAll"
    tSql = tSql & vbCrLf & ",FCSdtSalePrice,FCSdtSaleAmt,FCSdtRegPrice,FCSdtSaleUnit,FCSdtB4DisChg"
    tSql = tSql & vbCrLf & ",FTSdtStaPdt,FCSdtStkFac"
    tSql = tSql & vbCrLf & ",FTSdtBarCode"                      '*Ao 2013/05/30  CommSheet TK-ISS3000-208-ST 00 00 01-Ao.xls
    tSql = tSql & vbCrLf & "," & tCS_CNLastFedIns
    tSql = tSql & vbCrLf & " VALUES ("
    tSql = tSql & vbCrLf & "'" & tVB_CNStoreCode & "','" & tVB_CNTerminalNum & "','" & tVB_Running & "'," & nR & ",'" & tUT_SKUCode & "'"
    tSql = tSql & vbCrLf & ",'" & SP_STRtSiggleCode(tSKUName) & "'"
    tSql = tSql & vbCrLf & "," & "'" & SP_STRtSiggleCode(tSKUName2) & "','" & SP_STRtSiggleCode(tSKUNameSrvDoc) & "','" & SP_STRtSiggleCode(tSKUNameSrvDoc2) & "'"
    tSql = tSql & vbCrLf & ",'" & tVB_Transaction & "','" & tUT_SaleDate & "','" & Time$ & "'"
    tSql = tSql & vbCrLf & "," & SP_FMTcCurPoint(cQty, True, False) & "," & SP_FMTcCurPoint(cQtyAll, True, False)
    tSql = tSql & vbCrLf & "," & cSalePrice
    tSql = tSql & vbCrLf & "," & cSaleAmt
    tSql = tSql & vbCrLf & "," & cSalePrice
    tSql = tSql & vbCrLf & "," & cSaleUnit
    tSql = tSql & vbCrLf & "," & cB4DisChg
    tSql = tSql & vbCrLf & ",'" & tSdtStaPdt & "'"
    tSql = tSql & vbCrLf & "," & cSdtStkFac
    tSql = tSql & vbCrLf & ",'T005'"    '*Ao 2013/05/30  CommSheet TK-ISS3000-208-ST 00 00 01-Ao.xls
    
    tSql = tSql & vbCrLf & "," & UT_tLastInsSql & ")"
    If SP_SQLbExecute(tSql, , , nEff) Then
        If nEff = 1 Then
            'none
        End If
    End If
   
   '*Tao 62-09-05 Sum Vendor Voucher
   nR = nR + 1
   cACTUAL = 0
   cQtyACTUAL = 0
   If Not (orsTemp.BOF And orsTemp.EOF) Then
        orsTemp.Filter = "FTTdmCode = 'T006'"
        If Not (orsTemp.BOF And orsTemp.EOF) Then
            Do While Not orsTemp.EOF
'                If SP_FEDtChkString(orsTemp, "FTTdmCode") = "T018" And (SP_FEDtChkString(orsTemp, "FTShdTransType") = "04" Or SP_FEDtChkString(orsTemp, "FTShdTransType") = "39" Or SP_FEDtChkString(orsTemp, "FTShdTransType") = "44" Or SP_FEDtChkString(orsTemp, "FTShdTransType") = "14") Then
'                    cACTUAL = cACTUAL + 0
'                    cQtyACTUAL = cQtyACTUAL + 0
'                Else
                    cACTUAL = cACTUAL + SP_FEDvChkNumeric(orsTemp, "FCSrcNet")
                    cQtyACTUAL = cQtyACTUAL + SP_FEDvChkNumeric(orsTemp, "FCQty")
'                End If
                orsTemp.MoveNext
            Loop
        End If
         orsTemp.Filter = adFilterNone
   End If

    tUT_SKUCode = "T006"
    tSKUName = "VVch"  'ชื่อสินค้า(Default ENG)1
    tSKUName2 = "VVch"   'ชื่อสินค้า(Local)2
    tSKUNameSrvDoc = tSKUName  'ชื่อสินค้า Service Doc(Default ENG) 1
    tSKUNameSrvDoc2 = tSKUName2    'ชื่อสินค้า Service Doc (Local)2

    cQty = cQtyACTUAL
    cQtyAll = cQtyACTUAL

    cRegPrice = cACTUAL
    cSalePrice = cACTUAL
    cSaleUnit = cACTUAL
    cB4DisChg = cACTUAL
    cSdtStkFac = cACTUAL
    cSaleAmt = cACTUAL
    cSaleAmt = cACTUAL
    tSdtStaPdt = "4"

    tSql = "INSERT INTO " & tVB_CNTblDT
    tSql = tSql & vbCrLf & "(FTStmCode,FTTmnNum,FTShdTransNo,FNSdtSeqNo,FTSkuCode"
    tSql = tSql & vbCrLf & ",FTSkuAbbName,FTSkuAbbNameSnd,FTSkuAbbNameSrvDoc,FTSkuAbbNameSndSrvDoc"
    tSql = tSql & vbCrLf & ",FTShdTransType,FDShdTransDate,FTShdTransTime,FCSdtQty,FCSdtQtyAll"
    tSql = tSql & vbCrLf & ",FCSdtSalePrice,FCSdtSaleAmt,FCSdtRegPrice,FCSdtSaleUnit,FCSdtB4DisChg"
    tSql = tSql & vbCrLf & ",FTSdtStaPdt,FCSdtStkFac"
    tSql = tSql & vbCrLf & ",FTSdtBarCode"
    tSql = tSql & vbCrLf & "," & tCS_CNLastFedIns
    tSql = tSql & vbCrLf & " VALUES ("
    tSql = tSql & vbCrLf & "'" & tVB_CNStoreCode & "','" & tVB_CNTerminalNum & "','" & tVB_Running & "'," & nR & ",'" & tUT_SKUCode & "'"
    tSql = tSql & vbCrLf & ",'" & SP_STRtSiggleCode(tSKUName) & "'"
    tSql = tSql & vbCrLf & "," & "'" & SP_STRtSiggleCode(tSKUName2) & "','" & SP_STRtSiggleCode(tSKUNameSrvDoc) & "','" & SP_STRtSiggleCode(tSKUNameSrvDoc2) & "'"
    tSql = tSql & vbCrLf & ",'" & tVB_Transaction & "','" & tUT_SaleDate & "','" & Time$ & "'"
    tSql = tSql & vbCrLf & "," & SP_FMTcCurPoint(cQty, True, False) & "," & SP_FMTcCurPoint(cQtyAll, True, False)
    tSql = tSql & vbCrLf & "," & cSalePrice
    tSql = tSql & vbCrLf & "," & cSaleAmt
    tSql = tSql & vbCrLf & "," & cSalePrice
    tSql = tSql & vbCrLf & "," & cSaleUnit
    tSql = tSql & vbCrLf & "," & cB4DisChg
    tSql = tSql & vbCrLf & ",'" & tSdtStaPdt & "'"
    tSql = tSql & vbCrLf & "," & cSdtStkFac
    tSql = tSql & vbCrLf & ",'T006'"
    tSql = tSql & vbCrLf & "," & UT_tLastInsSql & ")"
    If SP_SQLbExecute(tSql, , , nEff) Then
        If nEff = 1 Then
            'none
        End If
    End If

'    orsTemp.Close               '*Ao 2013/10/11      V4.0.78
   Call SP_DATxRsNothing(orsTemp)
   
    tSql = wFunctionKB.W_GETtCloseTenderCurSQL(tVB_CNTblHD, tVB_CNTblRC, "FTShdTransType", True)
    Call SP_SQLvExecute(tSql, orsTemp)
    tSql = "SELECT * FROM TCNMForCurrency"
    tSql = tSql & vbCrLf & " ORDER BY FTFcrCode"
    Call SP_SQLvExecute(tSql, orsFor, oDB.ocnOnLine, , , , , True, True) 'read TCNMForCurrency
    If Not (orsFor.BOF Or orsFor.EOF) Then
            Do While Not orsFor.EOF
                    nR = nR + 1
                    cACTUAL = 0
                    cQtyACTUAL = 0
                    cACTUALTmp = 0
                     If Not (orsTemp.BOF And orsTemp.EOF) Then
                        orsTemp.Filter = "FTRteCode = '" & SP_FEDtChkString(orsFor, "FTFcrCode") & "'"
                        If Not (orsTemp.BOF And orsTemp.EOF) Then
                            'cACTUAL = SP_FEDvChkNumeric(orsTemp, "FCSrcNet") 'TK-ISS3000-164 Over/Short Report - Cash & Foreign Currency problem
'                            cACTUAL = SP_FEDvChkNumeric(orsTemp, "FCSrcRteAmt")
                            cACTUAL = SP_FEDvChkNumeric(orsTemp, "FCSrcNet")               '*Ao แก้ไขทำให้การคำนวณ OVER/SHORT(BASE)
                            cACTUALTmp = SP_FEDvChkNumeric(orsTemp, "FCSrcRteAmt")
                            cQtyACTUAL = SP_FEDvChkNumeric(orsTemp, "FCQty")
                        End If
                        orsTemp.Filter = adFilterNone
                    End If
                    tUT_SKUCode = SP_FEDtChkString(orsFor, "FTFcrCode")
                    tSKUName = Trim(SP_FEDtChkString(orsFor, "FTFcrType")) 'ชื่อสินค้า(Default ENG)1
                    tSKUName2 = Trim(SP_FEDtChkString(orsFor, "FTFcrType"))   'ชื่อสินค้า(Local)2                   '*Ao 2013/03/14  TK-ISS3000-094
                    tSKUNameSrvDoc = tSKUName  'ชื่อสินค้า Service Doc(Default ENG) 1
                    tSKUNameSrvDoc2 = tSKUName2   'ชื่อสินค้า Service Doc (Local)2

                    cQty = cQtyACTUAL
                    cQtyAll = cQtyACTUAL

                    cRegPrice = cACTUAL
'                    cSalePrice = cACTUAL
                    cSalePrice = cACTUALTmp     '*Ao แก้ไขทำให้การคำนวณ OVER/SHORT(BASE)
                    cSaleUnit = cACTUAL
                    cB4DisChg = cACTUAL
                    cSdtStkFac = cACTUAL
                    cSaleAmt = cACTUAL
                    cSaleAmt = cACTUAL
                    tSdtStaPdt = "4"

                    tSql = "INSERT INTO " & tVB_CNTblDT
                    tSql = tSql & vbCrLf & "(FTStmCode,FTTmnNum,FTShdTransNo,FNSdtSeqNo,FTSkuCode"
                    tSql = tSql & vbCrLf & ",FTSkuAbbName,FTSkuAbbNameSnd,FTSkuAbbNameSrvDoc,FTSkuAbbNameSndSrvDoc"
                    tSql = tSql & vbCrLf & ",FTShdTransType,FDShdTransDate,FTShdTransTime,FCSdtQty,FCSdtQtyAll"
                    tSql = tSql & vbCrLf & ",FCSdtSalePrice,FCSdtSaleAmt,FCSdtRegPrice,FCSdtSaleUnit,FCSdtB4DisChg"
                    tSql = tSql & vbCrLf & ",FTSdtStaPdt,FCSdtStkFac"
                    tSql = tSql & vbCrLf & ",FTSdtBarCode"                      '*Ao 2013/05/30  CommSheet TK-ISS3000-208-ST 00 00 01-Ao.xls
                    tSql = tSql & vbCrLf & "," & tCS_CNLastFedIns
                    tSql = tSql & vbCrLf & " VALUES ("
                    tSql = tSql & vbCrLf & "'" & tVB_CNStoreCode & "','" & tVB_CNTerminalNum & "','" & tVB_Running & "'," & nR & ",'" & tUT_SKUCode & "'"
                    tSql = tSql & vbCrLf & ",'" & SP_STRtSiggleCode(tSKUName) & "'"
                    tSql = tSql & vbCrLf & "," & "'" & SP_STRtSiggleCode(tSKUName2) & "','" & SP_STRtSiggleCode(tSKUNameSrvDoc) & "','" & SP_STRtSiggleCode(tSKUNameSrvDoc2) & "'"
                    tSql = tSql & vbCrLf & ",'" & tVB_Transaction & "','" & tUT_SaleDate & "','" & Time$ & "'"
                    tSql = tSql & vbCrLf & "," & SP_FMTcCurPoint(cQty, True, False) & "," & SP_FMTcCurPoint(cQtyAll, True, False)
                    tSql = tSql & vbCrLf & "," & cSalePrice
                    tSql = tSql & vbCrLf & "," & cSaleAmt
                    tSql = tSql & vbCrLf & "," & cSalePrice
                    tSql = tSql & vbCrLf & "," & cSaleUnit
                    tSql = tSql & vbCrLf & "," & cB4DisChg
                    tSql = tSql & vbCrLf & ",'" & tSdtStaPdt & "'"
                    tSql = tSql & vbCrLf & "," & cSdtStkFac
                    'tSql = tSql & vbCrLf & "," & "'T012'"                           '*Ao 2013/05/30  CommSheet TK-ISS3000-208-ST 00 00 01-Ao.xls
                    tSql = tSql & vbCrLf & ",'" & tUT_SKUCode & "'"   '*BG 56-07-09  CommSheet TK-ISS3000-208-ST 00 00 01-Ao.xls
                    tSql = tSql & vbCrLf & "," & UT_tLastInsSql & ")"
                    If SP_SQLbExecute(tSql, , , nEff) Then
                        If nEff = 1 Then
                            'none
                        End If
                    End If
            orsFor.MoveNext
        Loop
        '
        nR = nR + 1
        tUT_SKUCode = "T012"
        tSKUName = "Foreign(Base)"  'ชื่อสินค้า(Default ENG)1
        tSKUName2 = "Foreign(Base)"  'ชื่อสินค้า(Local)2
        tSKUNameSrvDoc = tSKUName  'ชื่อสินค้า Service Doc(Default ENG) 1
        tSKUNameSrvDoc2 = tSKUName2   'ชื่อสินค้า Service Doc (Local)2
        
        cRegPrice = 0
        cSalePrice = 0
        cSaleUnit = 0
        cB4DisChg = 0
        cSdtStkFac = 0
        cSaleAmt = 0
        cSaleAmt = 0
        cQty = 0
        cQtyAll = 0
        tSdtStaPdt = "4"

        tSql = "INSERT INTO " & tVB_CNTblDT
        tSql = tSql & vbCrLf & "(FTStmCode,FTTmnNum,FTShdTransNo,FNSdtSeqNo,FTSkuCode"
        tSql = tSql & vbCrLf & ",FTSkuAbbName,FTSkuAbbNameSnd,FTSkuAbbNameSrvDoc,FTSkuAbbNameSndSrvDoc"
        tSql = tSql & vbCrLf & ",FTShdTransType,FDShdTransDate,FTShdTransTime,FCSdtQty,FCSdtQtyAll"
        tSql = tSql & vbCrLf & ",FCSdtSalePrice,FCSdtSaleAmt,FCSdtRegPrice,FCSdtSaleUnit,FCSdtB4DisChg"
        tSql = tSql & vbCrLf & ",FTSdtStaPdt,FCSdtStkFac"
        tSql = tSql & vbCrLf & ",FTSdtBarCode"                      '*Ao 2013/05/30  CommSheet TK-ISS3000-208-ST 00 00 01-Ao.xls
        tSql = tSql & vbCrLf & "," & tCS_CNLastFedIns
        tSql = tSql & vbCrLf & " VALUES ("
        tSql = tSql & vbCrLf & "'" & tVB_CNStoreCode & "','" & tVB_CNTerminalNum & "','" & tVB_Running & "'," & nR & ",'" & tUT_SKUCode & "'"
        tSql = tSql & vbCrLf & ",'" & SP_STRtSiggleCode(tSKUName) & "'"
        tSql = tSql & vbCrLf & "," & "'" & SP_STRtSiggleCode(tSKUName2) & "','" & SP_STRtSiggleCode(tSKUNameSrvDoc) & "','" & SP_STRtSiggleCode(tSKUNameSrvDoc2) & "'"
        tSql = tSql & vbCrLf & ",'" & tVB_Transaction & "','" & tUT_SaleDate & "','" & Time$ & "'"
        tSql = tSql & vbCrLf & "," & SP_FMTcCurPoint(cQty, True, False) & "," & SP_FMTcCurPoint(cQtyAll, True, False)
        tSql = tSql & vbCrLf & "," & cSalePrice
        tSql = tSql & vbCrLf & "," & cSaleAmt
        tSql = tSql & vbCrLf & "," & cSalePrice
        tSql = tSql & vbCrLf & "," & cSaleUnit
        tSql = tSql & vbCrLf & "," & cB4DisChg
        tSql = tSql & vbCrLf & ",'" & tSdtStaPdt & "'"
        tSql = tSql & vbCrLf & "," & cSdtStkFac
        tSql = tSql & vbCrLf & ",'" & tUT_SKUCode & "'"         '*Ao 2013/05/30  CommSheet TK-ISS3000-208-ST 00 00 01-Ao.xls
        tSql = tSql & vbCrLf & "," & UT_tLastInsSql & ")"
        If SP_SQLbExecute(tSql, , , nEff) Then
            If nEff = 1 Then
                'none
            End If
        End If
       

'   '***CASH TRANSFER REPORT***
'   'summary pickup money
'        For nI = 0 To 2
'            Select Case nI
'                Case 0
'                    tUT_SKUCode = "T001"
'                    tSKUName = "Cash" 'ชื่อสินค้า(Default ENG)1
'                    tSKUName2 = "Cash" 'ชื่อสินค้า(Local)2
'                    tSKUNameSrvDoc = "Cash" 'ชื่อสินค้า Service Doc(Default ENG) 1
'                    tSKUNameSrvDoc2 = "Cash"   'ชื่อสินค้า Service Doc (Local)2
'                    cNet = Val(Format(otbCoin2.Text)) + Val(Format(otbCurrency2.Text))
'                Case 1
'                    tUT_SKUCode = "T018"
'                    tSKUName = "GVch" 'ชื่อสินค้า(Default ENG)1
'                    tSKUName2 = "GVch" 'ชื่อสินค้า(Local)2
'                    tSKUNameSrvDoc = "GVch" 'ชื่อสินค้า Service Doc(Default ENG) 1
'                    tSKUNameSrvDoc2 = "GVch"   'ชื่อสินค้า Service Doc (Local)2
'                    cNet = Val(Format(otbGV2.Text))
'                Case 2
'                    tUT_SKUCode = "T012"
'                    tSKUName = "Foreign(Base)" 'ชื่อสินค้า(Default ENG)1
'                    tSKUName2 = "Foreign(Base)" 'ชื่อสินค้า(Local)2
'                    tSKUNameSrvDoc = "Foreign(Base)" 'ชื่อสินค้า Service Doc(Default ENG) 1
'                    tSKUNameSrvDoc2 = "Foreign(Base)"   'ชื่อสินค้า Service Doc (Local)2
'                    cNet = Val(Format(otbForeign2.Text))
'            End Select
'            nR = nR + 1
'            cQty = 1
'            cQtyAll = 1
'
'            cRegPrice = cNet
'            cSalePrice = cNet
'            cSaleUnit = cNet
'            cB4DisChg = cNet
'            cSdtStkFac = cNet
'            cSaleAmt = cNet
'            cSaleAmt = cNet
'            tSdtStaPdt = "5"
'
'            tSql = "INSERT INTO " & tVB_CNTblDT
'            tSql = tSql & vbCrLf & "(FTStmCode,FTTmnNum,FTShdTransNo,FNSdtSeqNo,FTSkuCode"
'            tSql = tSql & vbCrLf & ",FTSkuAbbName,FTSkuAbbNameSnd,FTSkuAbbNameSrvDoc,FTSkuAbbNameSndSrvDoc"
'            tSql = tSql & vbCrLf & ",FTShdTransType,FDShdTransDate,FTShdTransTime,FCSdtQty,FCSdtQtyAll"
'            tSql = tSql & vbCrLf & ",FCSdtSalePrice,FCSdtSaleAmt,FCSdtRegPrice,FCSdtSaleUnit,FCSdtB4DisChg"
'            tSql = tSql & vbCrLf & ",FTSdtStaPdt,FCSdtStkFac,"
'            tSql = tSql & vbCrLf & tCS_CNLastFedIns
'            tSql = tSql & vbCrLf & " VALUES ("
'            tSql = tSql & vbCrLf & "'" & tVB_CNStoreCode & "','" & tVB_CNTerminalNum & "','" & tVB_Running & "'," & nR & ",'" & tUT_SKUCode & "'"
'            tSql = tSql & vbCrLf & ",'" & SP_STRtSiggleCode(tSKUName) & "'"
'            tSql = tSql & vbCrLf & "," & "'" & SP_STRtSiggleCode(tSKUName2) & "','" & SP_STRtSiggleCode(tSKUNameSrvDoc) & "','" & SP_STRtSiggleCode(tSKUNameSrvDoc2) & "'"
'            tSql = tSql & vbCrLf & ",'" & tVB_Transaction & "','" & tUT_SaleDate & "','" & Time$ & "'"
'            tSql = tSql & vbCrLf & "," & SP_FMTcCurPoint(cQty, True, False) & "," & SP_FMTcCurPoint(cQtyAll, True, False)
'            tSql = tSql & vbCrLf & "," & cSalePrice
'            tSql = tSql & vbCrLf & "," & cSaleAmt
'            tSql = tSql & vbCrLf & "," & cSalePrice
'            tSql = tSql & vbCrLf & "," & cSaleUnit
'            tSql = tSql & vbCrLf & "," & cB4DisChg
'            tSql = tSql & vbCrLf & ",'" & tSdtStaPdt & "'"
'            tSql = tSql & vbCrLf & "," & cSdtStkFac
'            tSql = tSql & vbCrLf & "," & UT_tLastInsSql & ")"
'            If SP_SQLbExecute(tSql, , , nEff) Then
'                If nEff = 1 Then
'                   'none
'                End If
'            End If
'        Next nI
    '***CASH TRANSFER REPORT***
     'summary pickup money
     '57-05-08   20130504-ReEntry.xls Re-Entry ไม่ต้องแสดง ASH TRANSFER REPORT
    Select Case SP_VALnText2Double(tVB_TransactionSub)
        Case tEN_TrnOverShortReport
            tTransWhere = tVB_CNTblHD & ".FDShdTransDate = '" & tUT_SaleDate & "' AND " & tVB_CNTblHD & ".FTTmnNum ='" & tVB_CNTerminalNum & "'"
            tTransWhere = tTransWhere & vbCrLf & " AND " & tVB_CNTblHD & ".FTShdTransType = '12'"
'            tTransWhere = tTransWhere & vbCrLf & " AND " & tVB_CNTblHD & ".FTLogCode  = '" & tUT_LoginCode & "'"
            tTransWhere = tTransWhere & vbCrLf & " GROUP BY FTSkuCode,FTSkuAbbName,FTSkuAbbNameSnd,FTSdtStaPdt,FCSdtStkFac"
            tSql = "SELECT FTSkuCode,FTSkuAbbName,FTSkuAbbNameSnd,FTSdtStaPdt,FCSdtStkFac,SUM(FCSdtSaleAmt) AS FCSum ,SUM(FCSdtRegPrice) AS FCSum2 "
            tSql = tSql & vbCrLf & "FROM   " & tVB_CNTblHD & " INNER JOIN"
            tSql = tSql & vbCrLf & "               " & tVB_CNTblDT & " ON " & tVB_CNTblHD & ".FTTmnNum = " & tVB_CNTblDT & ".FTTmnNum AND " & tVB_CNTblHD & ".FTShdTransNo = " & tVB_CNTblDT & ".FTShdTransNo AND "
            tSql = tSql & vbCrLf & "               " & tVB_CNTblHD & ".FTShdTransType = " & tVB_CNTblDT & ".FTShdTransType And " & tVB_CNTblHD & ".FDShdTransDate = " & tVB_CNTblDT & ".FDShdTransDate"
            tSql = tSql & vbCrLf & " WHERE " & tTransWhere
            nError = SP_SQLvExecute(tSql, orsPickup, oDB.ocnOffLine, , , , , False)
            If nError = 0 Then
                cNet = 0
                nR = nR + 1
                If Not (orsPickup.BOF And orsPickup.EOF) Then
                       orsPickup.Filter = "FTSdtStaPdt = '0' OR FTSdtStaPdt = '1'"
                       If Not (orsPickup.BOF And orsPickup.EOF) Then
                           Do While Not orsPickup.EOF
                               cNet = cNet + SP_FEDvChkNumeric(orsPickup, "FCSum") 'TK-ISS3000-164 Over/Short Report - Cash & Foreign Currency problem
                               orsPickup.MoveNext
                           Loop
                       End If
                       orsPickup.Filter = adFilterNone
                End If
                tUT_SKUCode = "T001" '"Cash"
                tSKUName = "Cash" 'ชื่อสินค้า(Default ENG)1
                tSKUName2 = "Cash" 'ชื่อสินค้า(Local)2
                tSKUNameSrvDoc = tSKUName 'ชื่อสินค้า Service Doc(Default ENG) 1
                tSKUNameSrvDoc2 = tSKUName2   'ชื่อสินค้า Service Doc (Local)2
                If cNet = 0 Then
                  cQty = 0
                  cQtyAll = 0
                Else
                  cQty = 1
                  cQtyAll = 1
                End If
                cRegPrice = cNet
                cSalePrice = cNet
                cSaleUnit = cNet
                cB4DisChg = cNet
                cSdtStkFac = cNet
                cSaleAmt = cNet
                cSaleAmt = cNet
                tSdtStaPdt = "5"
              
                tSql = "INSERT INTO " & tVB_CNTblDT
                tSql = tSql & vbCrLf & "(FTStmCode,FTTmnNum,FTShdTransNo,FNSdtSeqNo,FTSkuCode"
                tSql = tSql & vbCrLf & ",FTSkuAbbName,FTSkuAbbNameSnd,FTSkuAbbNameSrvDoc,FTSkuAbbNameSndSrvDoc"
                tSql = tSql & vbCrLf & ",FTShdTransType,FDShdTransDate,FTShdTransTime,FCSdtQty,FCSdtQtyAll"
                tSql = tSql & vbCrLf & ",FCSdtSalePrice,FCSdtSaleAmt,FCSdtRegPrice,FCSdtSaleUnit,FCSdtB4DisChg"
                tSql = tSql & vbCrLf & ",FTSdtStaPdt,FCSdtStkFac"
                tSql = tSql & vbCrLf & ",FTSdtBarCode"                      '*Ao 2013/05/30  CommSheet TK-ISS3000-208-ST 00 00 01-Ao.xls
                tSql = tSql & vbCrLf & "," & tCS_CNLastFedIns
                tSql = tSql & vbCrLf & " VALUES ("
                tSql = tSql & vbCrLf & "'" & tVB_CNStoreCode & "','" & tVB_CNTerminalNum & "','" & tVB_Running & "'," & nR & ",'" & tUT_SKUCode & "'"
                tSql = tSql & vbCrLf & ",'" & SP_STRtSiggleCode(tSKUName) & "'"
                tSql = tSql & vbCrLf & "," & "'" & SP_STRtSiggleCode(tSKUName2) & "','" & SP_STRtSiggleCode(tSKUNameSrvDoc) & "','" & SP_STRtSiggleCode(tSKUNameSrvDoc2) & "'"
                tSql = tSql & vbCrLf & ",'" & tVB_Transaction & "','" & tUT_SaleDate & "','" & Time$ & "'"
                tSql = tSql & vbCrLf & "," & SP_FMTcCurPoint(cQty, True, False) & "," & SP_FMTcCurPoint(cQtyAll, True, False)
                tSql = tSql & vbCrLf & "," & cSalePrice
                tSql = tSql & vbCrLf & "," & cSaleAmt
                tSql = tSql & vbCrLf & "," & cSalePrice
                tSql = tSql & vbCrLf & "," & cSaleUnit
                tSql = tSql & vbCrLf & "," & cB4DisChg
                tSql = tSql & vbCrLf & ",'" & tSdtStaPdt & "'"
                tSql = tSql & vbCrLf & "," & cSdtStkFac
                tSql = tSql & vbCrLf & ",'" & tUT_SKUCode & "'"           '*Ao 2013/05/30  CommSheet TK-ISS3000-208-ST 00 00 01-Ao.xls
                tSql = tSql & vbCrLf & "," & UT_tLastInsSql & ")"
                If SP_SQLbExecute(tSql, , , nEff) Then
                    If nEff = 1 Then
                       'none
                    End If
                End If
                  
                  'CV
                nR = nR + 1
                cNet = 0
                 If Not (orsPickup.BOF And orsPickup.EOF) Then
                        orsPickup.Filter = "FTSdtStaPdt = '2'"
                        If Not (orsPickup.BOF And orsPickup.EOF) Then
                            Do While Not orsPickup.EOF
'                                  cNet = SP_FEDvChkNumeric(orsPickup, "FCSum")
                                cNet = cNet + SP_FEDvChkNumeric(orsPickup, "FCSum")  '*BG 56-09-17 Comsheet-355 รวม GV
                                orsPickup.MoveNext
                            Loop
                        End If
                        orsPickup.Filter = adFilterNone
                 End If
'                  tUT_SKUCode = "T018" '"GVch"
'                  tSKUName = "GVch" 'ชื่อสินค้า(Default ENG)1
'                  tSKUName2 = "GVch" 'ชื่อสินค้า(Local)2
'                  tSKUNameSrvDoc = tSKUName 'ชื่อสินค้า Service Doc(Default ENG) 1
'                  tSKUNameSrvDoc2 = tSKUName2   'ชื่อสินค้า Service Doc (Local)2
                tUT_SKUCode = "T005" '"GVch"
                tSKUName = "GVch" 'ชื่อสินค้า(Default ENG)1
                tSKUName2 = "GVch" 'ชื่อสินค้า(Local)2
                tSKUNameSrvDoc = tSKUName 'ชื่อสินค้า Service Doc(Default ENG) 1
                tSKUNameSrvDoc2 = tSKUName2   'ชื่อสินค้า Service Doc (Local)2
                  
                If cNet = 0 Then
                  cQty = 0
                  cQtyAll = 0
                Else
                  cQty = 1
                  cQtyAll = 1
                End If
                cRegPrice = cNet
                cSalePrice = cNet
                cSaleUnit = cNet
                cB4DisChg = cNet
                cSdtStkFac = cNet
                cSaleAmt = cNet
                cSaleAmt = cNet
                tSdtStaPdt = "5"
              
                tSql = "INSERT INTO " & tVB_CNTblDT
                tSql = tSql & vbCrLf & "(FTStmCode,FTTmnNum,FTShdTransNo,FNSdtSeqNo,FTSkuCode"
                tSql = tSql & vbCrLf & ",FTSkuAbbName,FTSkuAbbNameSnd,FTSkuAbbNameSrvDoc,FTSkuAbbNameSndSrvDoc"
                tSql = tSql & vbCrLf & ",FTShdTransType,FDShdTransDate,FTShdTransTime,FCSdtQty,FCSdtQtyAll"
                tSql = tSql & vbCrLf & ",FCSdtSalePrice,FCSdtSaleAmt,FCSdtRegPrice,FCSdtSaleUnit,FCSdtB4DisChg"
                tSql = tSql & vbCrLf & ",FTSdtStaPdt,FCSdtStkFac"
                tSql = tSql & vbCrLf & ",FTSdtBarCode"                      '*Ao 2013/05/30  CommSheet TK-ISS3000-208-ST 00 00 01-Ao.xls
                tSql = tSql & vbCrLf & "," & tCS_CNLastFedIns
                tSql = tSql & vbCrLf & " VALUES ("
                tSql = tSql & vbCrLf & "'" & tVB_CNStoreCode & "','" & tVB_CNTerminalNum & "','" & tVB_Running & "'," & nR & ",'" & tUT_SKUCode & "'"
                tSql = tSql & vbCrLf & ",'" & SP_STRtSiggleCode(tSKUName) & "'"
                tSql = tSql & vbCrLf & "," & "'" & SP_STRtSiggleCode(tSKUName2) & "','" & SP_STRtSiggleCode(tSKUNameSrvDoc) & "','" & SP_STRtSiggleCode(tSKUNameSrvDoc2) & "'"
                tSql = tSql & vbCrLf & ",'" & tVB_Transaction & "','" & tUT_SaleDate & "','" & Time$ & "'"
                tSql = tSql & vbCrLf & "," & SP_FMTcCurPoint(cQty, True, False) & "," & SP_FMTcCurPoint(cQtyAll, True, False)
                tSql = tSql & vbCrLf & "," & cSalePrice
                tSql = tSql & vbCrLf & "," & cSaleAmt
                tSql = tSql & vbCrLf & "," & cSalePrice
                tSql = tSql & vbCrLf & "," & cSaleUnit
                tSql = tSql & vbCrLf & "," & cB4DisChg
                tSql = tSql & vbCrLf & ",'" & tSdtStaPdt & "'"
                tSql = tSql & vbCrLf & "," & cSdtStkFac
                tSql = tSql & vbCrLf & ",'" & tUT_SKUCode & "'"           '*Ao 2013/05/30  CommSheet TK-ISS3000-208-ST 00 00 01-Ao.xls
                tSql = tSql & vbCrLf & "," & UT_tLastInsSql & ")"
                If SP_SQLbExecute(tSql, , , nEff) Then
                    If nEff = 1 Then
                       'none
                    End If
                End If
                  
                  'Foreign Currency
                Dim bFound As Boolean
'                  If orsPickup.EOF Then
'                        bFound = True
'                  Else
'                        orsPickup.Filter = "FTSkuCode = 'T012'"
'                        If Not (orsPickup.BOF And orsPickup.EOF) Then          '*Ao 2013/09/23  CommSheet TK-ISS3000-305-UAT_00.00.03 FSBT NG
'                            bFound = True
'                        Else
'                            bFound = False
'                        End If
'                        orsPickup.Filter = adFilterNone
'                  End If
'                    bFound = True               '*Ao 2013/10/08  CommSheet TK-ISS3000-374-FAPL-133.xlsx
                bFound = False
                    
                If bFound Then
                    nR = nR + 1
                    cNet = 0
                     If Not (orsPickup.BOF And orsPickup.EOF) Then
                        orsPickup.Filter = "FTSdtStaPdt = '3'" ' AND FTSkuCode = '" & .TextMatrix(nI, 4) & "'"
                        If Not (orsPickup.BOF And orsPickup.EOF) Then
                            Do While Not orsPickup.EOF
                                cNet = cNet + SP_FEDvChkNumeric(orsPickup, "FCSum")
                                orsPickup.MoveNext
                            Loop
                        End If
                        orsPickup.Filter = adFilterNone
                    End If
                   
                    'tUT_SKUCode = "T012" '.TextMatrix(nI, 4)
                    'tUT_SKUCode = .TextMatrix(nL, 5)    '*BG 56-07-10   CommSheet TK-ISS3000-208-ST 00 00 01-Ao.xls
                    tSKUName = "Foreign(Base)" '.TextMatrix(nI, 1) 'ชื่อสินค้า(Default ENG)1
                    tSKUName2 = "Foreign(Base)" ' .TextMatrix(nI, 5)   'ชื่อสินค้า(Local)2
                    tSKUNameSrvDoc = "Foreign(Base)" ' tSKUName 'ชื่อสินค้า Service Doc(Default ENG) 1
                    tSKUNameSrvDoc2 = "Foreign(Base)" 'tSKUName2   'ชื่อสินค้า Service Doc (Local)2
                    If cNet = 0 Then
                        cQty = 0
                        cQtyAll = 0
                    Else
                        cQty = 1
                        cQtyAll = 1
                    End If
                    cRegPrice = cNet
                    cSalePrice = cNet
                    cSaleUnit = cNet
                    cB4DisChg = cNet
                    cSdtStkFac = cNet
                    cSaleAmt = cNet
                    cSaleAmt = cNet
                    tSdtStaPdt = "5"
                
                    tSql = "INSERT INTO " & tVB_CNTblDT
                    tSql = tSql & vbCrLf & "(FTStmCode,FTTmnNum,FTShdTransNo,FNSdtSeqNo,FTSkuCode"
                    tSql = tSql & vbCrLf & ",FTSkuAbbName,FTSkuAbbNameSnd,FTSkuAbbNameSrvDoc,FTSkuAbbNameSndSrvDoc"
                    tSql = tSql & vbCrLf & ",FTShdTransType,FDShdTransDate,FTShdTransTime,FCSdtQty,FCSdtQtyAll"
                    tSql = tSql & vbCrLf & ",FCSdtSalePrice,FCSdtSaleAmt,FCSdtRegPrice,FCSdtSaleUnit,FCSdtB4DisChg"
                    tSql = tSql & vbCrLf & ",FTSdtStaPdt,FCSdtStkFac"
                    tSql = tSql & vbCrLf & ",FTSdtBarCode"                      '*Ao 2013/05/30  CommSheet TK-ISS3000-208-ST 00 00 01-Ao.xls
                    tSql = tSql & vbCrLf & "," & tCS_CNLastFedIns
                    tSql = tSql & vbCrLf & " VALUES ("
                    tSql = tSql & vbCrLf & "'" & tVB_CNStoreCode & "','" & tVB_CNTerminalNum & "','" & tVB_Running & "'," & nR & ",'" & tUT_SKUCode & "'"
                    tSql = tSql & vbCrLf & ",'" & SP_STRtSiggleCode(tSKUName) & "'"
                    tSql = tSql & vbCrLf & "," & "'" & SP_STRtSiggleCode(tSKUName2) & "','" & SP_STRtSiggleCode(tSKUNameSrvDoc) & "','" & SP_STRtSiggleCode(tSKUNameSrvDoc2) & "'"
                    tSql = tSql & vbCrLf & ",'" & tVB_Transaction & "','" & tUT_SaleDate & "','" & Time$ & "'"
                    tSql = tSql & vbCrLf & "," & SP_FMTcCurPoint(cQty, True, False) & "," & SP_FMTcCurPoint(cQtyAll, True, False)
                    tSql = tSql & vbCrLf & "," & cSalePrice
                    tSql = tSql & vbCrLf & "," & cSaleAmt
                    tSql = tSql & vbCrLf & "," & cSalePrice
                    tSql = tSql & vbCrLf & "," & cSaleUnit
                    tSql = tSql & vbCrLf & "," & cB4DisChg
                    tSql = tSql & vbCrLf & ",'" & tSdtStaPdt & "'"
                    tSql = tSql & vbCrLf & "," & cSdtStkFac
                    tSql = tSql & vbCrLf & ",'" & tUT_SKUCode & "'"           '*Ao 2013/05/30  CommSheet TK-ISS3000-208-ST 00 00 01-Ao.xls
                    tSql = tSql & vbCrLf & "," & UT_tLastInsSql & ")"
                    If SP_SQLbExecute(tSql, , , nEff) Then
                        If nEff = 1 Then
                           'none
                        End If
                    End If
                Else
                         'TK-ISS3000-165 Over/Short Report - Foreign Currency deno did not show
'                          tSql = "SELECT * FROM TCNMForCurrency"
'                          tSql = tSql & vbCrLf & " ORDER BY FTFcrCode"                     '*Ao 2013/10/08 CommSheet TK-ISS3000-374-FAPL-133.xlsx
                    tSql = "SELECT LTRIM(RTRIM(FTFcrCode)) AS FTFcrCode,FTFcrSign,FTFcrType,FCFcrExchang FROM TCNMForCurrency ORDER BY FTFcrCode"
                    Call SP_SQLvExecute(tSql, orsCur, oDB.ocnOnLine, , , , , True, True) 'read TCNMForCurrency
                    If Not (orsCur.BOF Or orsCur.EOF) Then
'                        orsPickup.MoveFirst                         '*Ao 2013/10/11      EOF  V4.0.78
                        Do While Not orsCur.EOF
                            nR = nR + 1
                            cNet = 0
                            cReg = 0
                            cRate = 0
                             If Not (orsPickup.BOF And orsPickup.EOF) Then
                                orsPickup.Filter = "FTSdtStaPdt = '3' AND FTSkuCode = '" & SP_FEDtChkString(orsCur, "FTFcrCode") & "'"
                                If Not (orsPickup.BOF And orsPickup.EOF) Then
                                    Do While Not orsPickup.EOF
                                        cNet = cNet + SP_FEDvChkNumeric(orsPickup, "FCSum")
                                        cReg = cReg + SP_FEDvChkNumeric(orsPickup, "FCSum2")
                                        cRate = SP_FEDvChkNumeric(orsPickup, "FCSdtStkFac")
                                        orsPickup.MoveNext
                                    Loop
                                End If
                                orsPickup.Filter = adFilterNone
                            End If
                                 
'                            tUT_SKUCode = SP_FEDtChkString(orsCur, "FTTdmCode") '.TextMatrix(nI, 4)
                            tUT_SKUCode = SP_FEDtChkString(orsCur, "FTFcrCode") '.TextMatrix(nI, 4)
                            tSKUName = Trim(SP_FEDtChkString(orsCur, "FTFcrType"))  '.TextMatrix(nI, 1) 'ชื่อสินค้า(Default ENG)1
'                            tSKUName2 = SP_FEDtChkString(orsCur, "FTFcrCode")  'ชื่อสินค้า(Local)2                         '*Ao 2013/10/09 CommSheet TK-ISS3000-374-FAPL-133.xlsx
                            tSKUName2 = tSKUName
                            tSKUNameSrvDoc = tSKUName ' tSKUName 'ชื่อสินค้า Service Doc(Default ENG) 1"
                            tSKUNameSrvDoc2 = tSKUName2 'tSKUName2   'ชื่อสินค้า Service Doc (Local)2"
                            If cNet = 0 Then
                                cQty = 0
                                cQtyAll = 0
                            Else
                                cQty = 1
                                cQtyAll = 1
                            End If
                            cRegPrice = cReg
                            cSalePrice = cReg
                            cSaleUnit = cReg
                            cB4DisChg = cReg
'                            cSdtStkFac = cRate
                            cSdtStkFac = SP_FEDvChkNumeric(orsCur, "FCFcrExchang")
                            cSaleAmt = cNet
                            tSdtStaPdt = "5"
                              
                            tSql = "INSERT INTO " & tVB_CNTblDT
                            tSql = tSql & vbCrLf & "(FTStmCode,FTTmnNum,FTShdTransNo,FNSdtSeqNo,FTSkuCode"
                            tSql = tSql & vbCrLf & ",FTSkuAbbName,FTSkuAbbNameSnd,FTSkuAbbNameSrvDoc,FTSkuAbbNameSndSrvDoc"
                            tSql = tSql & vbCrLf & ",FTShdTransType,FDShdTransDate,FTShdTransTime,FCSdtQty,FCSdtQtyAll"
                            tSql = tSql & vbCrLf & ",FCSdtSalePrice,FCSdtSaleAmt,FCSdtRegPrice,FCSdtSaleUnit,FCSdtB4DisChg"
                            tSql = tSql & vbCrLf & ",FTSdtStaPdt,FCSdtStkFac"
                            tSql = tSql & vbCrLf & ",FTSdtBarCode"                      '*Ao 2013/05/30  CommSheet TK-ISS3000-208-ST 00 00 01-Ao.xls
                            tSql = tSql & vbCrLf & "," & tCS_CNLastFedIns
                            tSql = tSql & vbCrLf & " VALUES ("
                            tSql = tSql & vbCrLf & "'" & tVB_CNStoreCode & "','" & tVB_CNTerminalNum & "','" & tVB_Running & "'," & nR & ",'" & tUT_SKUCode & "'"
                            tSql = tSql & vbCrLf & ",'" & SP_STRtSiggleCode(tSKUName) & "'"
                            tSql = tSql & vbCrLf & "," & "'" & SP_STRtSiggleCode(tSKUName2) & "','" & SP_STRtSiggleCode(tSKUNameSrvDoc) & "','" & SP_STRtSiggleCode(tSKUNameSrvDoc2) & "'"
                            tSql = tSql & vbCrLf & ",'" & tVB_Transaction & "','" & tUT_SaleDate & "','" & Time$ & "'"
                            tSql = tSql & vbCrLf & "," & SP_FMTcCurPoint(cQty, True, False) & "," & SP_FMTcCurPoint(cQtyAll, True, False)
                            tSql = tSql & vbCrLf & "," & cSalePrice
                            tSql = tSql & vbCrLf & "," & cSaleAmt
                            tSql = tSql & vbCrLf & "," & cSalePrice
                            tSql = tSql & vbCrLf & "," & cSaleUnit
                            tSql = tSql & vbCrLf & "," & cB4DisChg
                            tSql = tSql & vbCrLf & ",'" & tSdtStaPdt & "'"
                            tSql = tSql & vbCrLf & "," & cSdtStkFac
                            tSql = tSql & vbCrLf & ",'T012'"
                            tSql = tSql & vbCrLf & "," & UT_tLastInsSql & ")"
                            If SP_SQLbExecute(tSql, , , nEff) Then
                                If nEff = 1 Then
                                   'none
                                End If
                            End If
                            orsCur.MoveNext
                        Loop
                    End If
                    '----------------------------------------------------------------------------------------------------
                    '*Ao 2013/10/09 CommSheet TK-ISS3000-374-FAPL-133.xlsx
                    orsPickup.Filter = adFilterNone
                    orsPickup.Filter = "FTSdtStaPdt = '3' AND FTSkuCode = 'T012'"
                    If Not (orsPickup.BOF And orsPickup.EOF) Then
                        nR = nR + 1
                        cNet = 0
                        cReg = 0
'                        cRate = 0
                        orsPickup.MoveFirst
                        Do While Not orsPickup.EOF
                            cNet = cNet + SP_FEDvChkNumeric(orsPickup, "FCSum")
                            cReg = cReg + SP_FEDvChkNumeric(orsPickup, "FCSum2")
'                            cRate = SP_FEDvChkNumeric(orsPickup, "FCSdtStkFac")
                            orsPickup.MoveNext
                        Loop
                    Else
                        nR = nR + 1
                        cNet = 0
                        cReg = 0
'                        cRate = 0
                    End If
                    
                     If True Then                       '*Ao 2013/10/09    เก็บตลอด
                        If cNet = 0 Then
                          cQty = 0
                          cQtyAll = 0
                        Else
                          cQty = 1
                          cQtyAll = 1
                        End If
                        
                        tUT_SKUCode = "T012"
                        tSKUName = "Foreign(Base)"
                        tSKUName2 = "Foreign(Base)"
                        tSKUNameSrvDoc = "Foreign(Base)"
                        tSKUNameSrvDoc2 = "Foreign(Base)"
                                
                        cRegPrice = cReg
                        cSalePrice = cReg
                        cSaleUnit = cReg
                        cB4DisChg = cReg
'                        cSdtStkFac = cRate             '*Ao 2013/10/14
                        cSdtStkFac = 1
                        cSaleAmt = cNet
                        tSdtStaPdt = "5"
                        
                        tSql = "INSERT INTO " & tVB_CNTblDT
                        tSql = tSql & vbCrLf & "(FTStmCode,FTTmnNum,FTShdTransNo,FNSdtSeqNo,FTSkuCode"
                        tSql = tSql & vbCrLf & ",FTSkuAbbName,FTSkuAbbNameSnd,FTSkuAbbNameSrvDoc,FTSkuAbbNameSndSrvDoc"
                        tSql = tSql & vbCrLf & ",FTShdTransType,FDShdTransDate,FTShdTransTime,FCSdtQty,FCSdtQtyAll"
                        tSql = tSql & vbCrLf & ",FCSdtSalePrice,FCSdtSaleAmt,FCSdtRegPrice,FCSdtSaleUnit,FCSdtB4DisChg"
                        tSql = tSql & vbCrLf & ",FTSdtStaPdt,FCSdtStkFac"
                        tSql = tSql & vbCrLf & ",FTSdtBarCode"                      '*Ao 2013/05/30  CommSheet TK-ISS3000-208-ST 00 00 01-Ao.xls
                        tSql = tSql & vbCrLf & "," & tCS_CNLastFedIns
                        tSql = tSql & vbCrLf & " VALUES ("
                        tSql = tSql & vbCrLf & "'" & tVB_CNStoreCode & "','" & tVB_CNTerminalNum & "','" & tVB_Running & "'," & nR & ",'" & tUT_SKUCode & "'"
                        tSql = tSql & vbCrLf & ",'" & SP_STRtSiggleCode(tSKUName) & "'"
                        tSql = tSql & vbCrLf & "," & "'" & SP_STRtSiggleCode(tSKUName2) & "','" & SP_STRtSiggleCode(tSKUNameSrvDoc) & "','" & SP_STRtSiggleCode(tSKUNameSrvDoc2) & "'"
                        tSql = tSql & vbCrLf & ",'" & tVB_Transaction & "','" & tUT_SaleDate & "','" & Time$ & "'"
                        tSql = tSql & vbCrLf & "," & SP_FMTcCurPoint(cQty, True, False) & "," & SP_FMTcCurPoint(cQtyAll, True, False)
                        tSql = tSql & vbCrLf & "," & cSalePrice
                        tSql = tSql & vbCrLf & "," & cSaleAmt
                        tSql = tSql & vbCrLf & "," & cSalePrice
                        tSql = tSql & vbCrLf & "," & cSaleUnit
                        tSql = tSql & vbCrLf & "," & cB4DisChg
                        tSql = tSql & vbCrLf & ",'" & tSdtStaPdt & "'"
                        tSql = tSql & vbCrLf & "," & cSdtStkFac
                        tSql = tSql & vbCrLf & ",'" & tUT_SKUCode & "'"           '*Ao 2013/05/30  CommSheet TK-ISS3000-208-ST 00 00 01-Ao.xls
                        tSql = tSql & vbCrLf & "," & UT_tLastInsSql & ")"
                        If SP_SQLbExecute(tSql, , , nEff) Then
                            If nEff = 1 Then
                               'none
                            End If
                        End If
                    End If
                    '----------------------------------------------------------------------------------------------------
                End If
            End If
        End Select
    End If
    Call SP_DATxRsNothing(orsFor)
    Call SP_DATxRsNothing(orsTemp)
    Call SP_DATxRsNothing(orsPickup)
    Call SP_DATxRsNothing(orsCur)
   Exit Sub
ErrHandler:
    Call SP_MSGxShowError(Err.Number, Err.Description)
    Call SP_DATxRsNothing(orsFor)
    Call SP_DATxRsNothing(orsTemp)
    Call SP_DATxRsNothing(orsPickup)
    Call SP_DATxRsNothing(orsCur)
End Sub

Private Sub otbCoin_Change()
    otbCoin.Text = SP_DEFtChangeNum(Trim(otbCoin.Text))
    otbCoin.SelStart = Len(otbCoin.Text)
    otbSumTender.Text = SP_FMTcCurPoint(Val(Format(otbCoin.Text)) + Val(Format(otbCurrency.Text)), True, False)
End Sub

Private Sub otbCoin_GotFocus()
   With otbCoin
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub otbCoin_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDelete         'delete
            otbCoin.Text = ""
    End Select
End Sub

Private Sub otbCoin_KeyPress(KeyAscii As Integer)
    Call UT_ChkKeyAscii(KeyAscii)
    Select Case KeyAscii
        Case 8, 6           'backspace, delete
            'do nothing
        Case 46             'ทศนิยม
            KeyAscii = 0
'            If InStr(Trim(otbSumTender.Text), ".") > 0 Then
'                KeyAscii = 0
'            End If
        Case 13
            'Call ocmOK_Click
            SendKeys "{TAB}"
        Case Else
            If Not IsNumeric(Chr(KeyAscii)) Then
                KeyAscii = 0
            End If
    End Select
End Sub

Private Sub otbCoin2_Change()
    otbCoin2.Text = SP_DEFtChangeNum(Trim(otbCoin2.Text))
    otbCoin2.SelStart = Len(otbCoin2.Text)
    'otbSumTender2.Text = SP_FMTcCurPoint(Val(Format(otbCoin2.Text)) + Val(Format(otbCurrency2.Text)) + Val(Format(otbGV2.Text)) + Val(Format(otbForeign2.Text)), True, False) '*Tao 62-09-11 ของเดิม
    otbSumTender2.Text = SP_FMTcCurPoint(Val(Format(otbCoin2.Text)) + Val(Format(otbCurrency2.Text)) + Val(Format(otbGV2.Text)) + Val(Format(otbForeign2.Text) + Val(Format(otbVVch2.Text))), True, False) '*Tao 62-09-11 เพิ่ม otbVVch2
End Sub

Private Sub otbCoin2_GotFocus()
   With otbCoin2
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub otbCoin2_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDelete         'delete
            otbCoin2.Text = ""
    End Select
End Sub

Private Sub otbCoin2_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 8, 6           'backspace, delete
            'do nothing
        Case 46             'ทศนิยม
            KeyAscii = 0
'            If InStr(Trim(otbSumTender.Text), ".") > 0 Then
'                KeyAscii = 0
'            End If
        Case 13
            SendKeys "{TAB}"
        Case Else
            If Not IsNumeric(Chr(KeyAscii)) Then
                KeyAscii = 0
            End If
    End Select
End Sub

Private Sub otbCurrency_Change()
    otbCurrency.Text = SP_DEFtChangeNum(Trim(otbCurrency.Text))
    otbCurrency.SelStart = Len(otbCurrency.Text)
    otbSumTender.Text = SP_FMTcCurPoint(Val(Format(otbCoin.Text)) + Val(Format(otbCurrency.Text)), True, False)
End Sub


Private Sub otbCurrency_GotFocus()
   With otbCurrency
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub otbCurrency_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDelete         'delete
            otbCurrency.Text = ""
    End Select
End Sub

Private Sub otbCurrency_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 8, 6           'backspace, delete
            'do nothing
        Case 46             'ทศนิยม
            KeyAscii = 0
'            If InStr(Trim(otbSumTender.Text), ".") > 0 Then
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

Private Sub otbCurrency2_Change()
    otbCurrency2.Text = SP_DEFtChangeNum(Trim(otbCurrency2.Text))
    otbCurrency2.SelStart = Len(otbCurrency2.Text)
    'otbSumTender2.Text = SP_FMTcCurPoint(Val(Format(otbCoin2.Text)) + Val(Format(otbCurrency2.Text)) + Val(Format(otbGV2.Text)) + Val(Format(otbForeign2.Text)), True, False) '*Tao 62-09-11 ของเดิม
    otbSumTender2.Text = SP_FMTcCurPoint(Val(Format(otbCoin2.Text)) + Val(Format(otbCurrency2.Text)) + Val(Format(otbGV2.Text)) + Val(Format(otbForeign2.Text) + Val(Format(otbVVch2.Text))), True, False) '*Tao 62-09-11 เพิ่ม otbVVch2
End Sub

Private Sub otbCurrency2_GotFocus()
   With otbCurrency2
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub otbCurrency2_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDelete         'delete
            otbCurrency2.Text = ""
    End Select
End Sub

Private Sub otbCurrency2_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 8, 6           'backspace, delete
            'do nothing
        Case 46             'ทศนิยม
            KeyAscii = 0
'            If InStr(Trim(otbSumTender.Text), ".") > 0 Then
'                KeyAscii = 0
'            End If
        Case 13
            SendKeys "{TAB}"
        Case Else
            If Not IsNumeric(Chr(KeyAscii)) Then
                KeyAscii = 0
            End If
    End Select
End Sub

Private Sub otbGV2_Change()
    otbGV2.Text = SP_DEFtChangeNum(Trim(otbGV2.Text))
    otbGV2.SelStart = Len(otbGV2.Text)
    'otbSumTender2.Text = SP_FMTcCurPoint(Val(Format(otbCoin2.Text)) + Val(Format(otbCurrency2.Text)) + Val(Format(otbGV2.Text)) + Val(Format(otbForeign2.Text)), True, False) '*Tao 62-09-11 ของเดิม
    otbSumTender2.Text = SP_FMTcCurPoint(Val(Format(otbCoin2.Text)) + Val(Format(otbCurrency2.Text)) + Val(Format(otbGV2.Text)) + Val(Format(otbForeign2.Text) + Val(Format(otbVVch2.Text))), True, False) '*Tao 62-09-11 เพิ่ม otbVVch2
End Sub

Private Sub otbGV2_GotFocus()
   With otbGV2
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub otbGV2_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDelete         'delete
            otbGV2.Text = ""
    End Select
End Sub

Private Sub otbGV2_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 8, 6           'backspace, delete
            'do nothing
        Case 46             'ทศนิยม
            KeyAscii = 0
'            If InStr(Trim(otbSumTender.Text), ".") > 0 Then
'                KeyAscii = 0
'            End If
        Case 13
            SendKeys "{TAB}"
        Case Else
            If Not IsNumeric(Chr(KeyAscii)) Then
                KeyAscii = 0
            End If
    End Select
End Sub

Private Sub otbForeign2_Change()
    otbForeign2.Text = SP_DEFtChangeNum(Trim(otbForeign2.Text))
    otbForeign2.SelStart = Len(otbForeign2.Text)
    'otbSumTender2.Text = SP_FMTcCurPoint(Val(Format(otbCoin2.Text)) + Val(Format(otbCurrency2.Text)) + Val(Format(otbGV2.Text)) + Val(Format(otbForeign2.Text)), True, False) '*Tao 62-09-11 ของเดิม
    otbSumTender2.Text = SP_FMTcCurPoint(Val(Format(otbCoin2.Text)) + Val(Format(otbCurrency2.Text)) + Val(Format(otbGV2.Text)) + Val(Format(otbForeign2.Text) + Val(Format(otbVVch2.Text))), True, False) '*Tao 62-09-11 เพิ่ม otbVVch2
End Sub

Private Sub otbForeign2_GotFocus()
   With otbForeign2
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub otbForeign2_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDelete         'delete
            otbForeign2.Text = ""
    End Select
End Sub

Private Sub otbForeign2_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 8, 6           'backspace, delete
            'do nothing
        Case 46             'ทศนิยม
            KeyAscii = 0
'            If InStr(Trim(otbSumTender.Text), ".") > 0 Then
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

Private Sub otbVVch2_Change()
    otbVVch2.Text = SP_DEFtChangeNum(Trim(otbVVch2.Text))
    otbVVch2.SelStart = Len(otbVVch2.Text)
    'otbSumTender2.Text = SP_FMTcCurPoint(Val(Format(otbCoin2.Text)) + Val(Format(otbCurrency2.Text)) + Val(Format(otbGV2.Text)) + Val(Format(otbForeign2.Text)), True, False) '*Tao 62-09-11 ของเดิม
    otbSumTender2.Text = SP_FMTcCurPoint(Val(Format(otbCoin2.Text)) + Val(Format(otbCurrency2.Text)) + Val(Format(otbGV2.Text)) + Val(Format(otbForeign2.Text) + Val(Format(otbVVch2.Text))), True, False) '*Tao 62-09-11 เพิ่ม otbVVch2
End Sub

Private Sub otbVVch2_GotFocus()
   With otbVVch2
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub otbVVch2_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDelete         'delete
            otbVVch2.Text = ""
    End Select
End Sub

Private Sub otbVVch2_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 8, 6           'backspace, delete
            'do nothing
        Case 46             'ทศนิยม
            KeyAscii = 0
'            If InStr(Trim(otbSumTender.Text), ".") > 0 Then
'                KeyAscii = 0
'            End If
        Case 13
            SendKeys "{TAB}"
        Case Else
            If Not IsNumeric(Chr(KeyAscii)) Then
                KeyAscii = 0
            End If
    End Select
End Sub
