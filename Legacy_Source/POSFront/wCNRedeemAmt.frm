VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.1#0"; "Codejock.SuiteCtrls.Unicode.v11.1.3.ocx"
Begin VB.Form wCNRedeemAmt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Deposit Balance;klg_Title"
   ClientHeight    =   3180
   ClientLeft      =   4290
   ClientTop       =   3600
   ClientWidth     =   7845
   ControlBox      =   0   'False
   Icon            =   "wCNRedeemAmt.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   7845
   Tag             =   "2;"
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   2520
      Index           =   0
      Left            =   0
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   -240
      Width           =   7845
      _Version        =   720897
      _ExtentX        =   13838
      _ExtentY        =   4445
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
      Begin VB.TextBox otbAccBal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0E0FF&
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
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Text            =   "0.00"
         Top             =   1680
         Width           =   2415
      End
      Begin VB.TextBox otbDepositNo 
         BackColor       =   &H00C0E0FF&
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
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   480
         Width           =   3975
      End
      Begin VB.TextBox otbDepositAmt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0E0FF&
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
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Text            =   "0.00"
         Top             =   1080
         Width           =   2415
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
         TabIndex        =   8
         Tag             =   "2;"
         Top             =   480
         Width           =   2235
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
         TabIndex        =   7
         Tag             =   "2;"
         Top             =   1080
         Width           =   2220
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Account Balance:;klg_Bal"
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
         TabIndex        =   6
         Tag             =   "2;"
         Top             =   1800
         Width           =   2205
      End
   End
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   1200
      Index           =   1
      Left            =   0
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1980
      Width           =   7845
      _Version        =   720897
      _ExtentX        =   13838
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
         Left            =   6240
         TabIndex        =   1
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
         Left            =   4680
         TabIndex        =   0
         Tag             =   "2;"
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label olaForm 
         BackStyle       =   0  'Transparent
         Caption         =   "Deposit GST Rate Information"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   735
         Index           =   3
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   4455
      End
   End
End
Attribute VB_Name = "wCNRedeemAmt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private orsHD  As ADODB.Recordset
Private tW_Whe As String

Private tW_TmnNumTmp As String
Private tW_TrnNoTmp As String
Private tW_TrnTypeTmp As String
Private dW_TrnDateTmp As Date
Private dW_DateTime As Date
Private tW_Time As String
Private bW_CANNOT_do As Boolean '*Bump 65-10-13 STK-POS ComSheet CR803 - 006.xlsx
Private Sub W_UPDxPhoneNo(ByVal ptPhone As String)
'-------------------------------------------------------------
'Call:
'Cmt: '*Eaw 56-10-28 CommSheet TK-ISS3000-401.xlsx
'--------------------------------------------------------------
Dim tSql As String

On Error GoTo ErrHandler
    If ptPhone <> "" Then
        tSql = "UPDATE " & tVB_CNTblHD
        tSql = tSql & vbCrLf & "SET FTPmhDocNoBill = '" & ptPhone & "'"
        tSql = tSql & vbCrLf & " WHERE  FTTmnNum='" & tVB_CNTerminalNum & "' AND FTShdTransNo='" & tVB_Running & "'"
        tSql = tSql & vbCrLf & " AND  FTShdTransType='" & tVB_Transaction & "' AND FDShdTransDate='" & tUT_SaleDate & "'"
        Call SP_SQLbExecute(tSql)
    End If
    Exit Sub
ErrHandler:
    
End Sub
Private Function W_PRCtPhone() As String
'-------------------------------------------------------------
'Call:
'Cmt: '*Eaw 56-10-28 CommSheet TK-ISS3000-401.xlsx
'--------------------------------------------------------------
Dim orsTemp As ADODB.Recordset
Dim tSql As String, tTmp As String

On Error GoTo ErrHandler
    tTmp = ""
    tSql = "SELECT ISNULL(FTPmhDocNoBill,'') AS FTPmhDocNoBill FROM TPSTDepositHD"
    tSql = tSql & vbCrLf & "WHERE FTCdpDepositNo = '" & Trim(otbDepositNo.Text) & "'"
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

Private Sub Form_Load()
      Call SP_FRMxChildCenter(Me)
      Call SP_FRMxGetCaptionFrmDB2Tag(Me)
      Call SP_FRMxSetCapForm(Me)
      Call W_FRMxSetInitial
      Call W_FRMxShowData
End Sub

Private Sub W_FRMxSetInitial()
'-------------------------------------------------------------
'Call:
'Cmt: set ค่าเบื้องสำหรับการเริ่มต้นการทำงานขิงหน้าจอนี้
'--------------------------------------------------------------
    'Set valuse initial
    olaForm(3).Visible = False
    bVB_TrnRedeem = False
    bVB_TrnAddDeposit = False
    otbDepositNo.Text = ""
    otbDepositAmt.Text = ""
    otbAccBal.Text = ""
    'Lock control
    otbDepositNo.Locked = True
    otbDepositNo.BackColor = nVB_CNClrBLock
    otbDepositAmt.Locked = True
    otbDepositAmt.BackColor = nVB_CNClrBLock
    otbAccBal.Locked = True
    otbAccBal.BackColor = nVB_CNClrBLock
End Sub

Private Sub W_FRMxShowData()
'------------------------------------------------------------------------
'Call:
'Cmt:
'-----------------------------------------------------------------------
'PH 2.0.0
    Dim tSql$
    Dim nSeqNo As Long
    Dim tTmnNumTmp As String
    Dim tTrnNoTmp As String
    Dim tTrnTypeTmp As String
    Dim dTrnDateTmp As Date
    Dim dDateTime As Date
    Dim tTime As String
    Select Case Val(Format(tVB_TransactionSub))
        Case tEN_TrnRedeem:
            tSql = "STP_CNxSQLSelDepositHDByWhr @ptTmnNum ='" & tVB_TerminalNoRedeem & "',@ptRunning ='" & tVB_TransNoRedeem & "',@ptTransType ='" & tVB_TransTypeRedeem & "',@ptTransDate ='" & SP_DTEtStrSTD(dVB_TransDateRedeem) & "'"
        Case tEN_TrnDepositCancel:
            tSql = "STP_CNxSQLSelDepositHDByWhr @ptTmnNum ='" & tVB_TerminalNoDptCancel & "',@ptRunning ='" & tVB_TransNoDptCancel & "',@ptTransType ='" & tVB_TransTypeDptCancel & "',@ptTransDate ='" & SP_DTEtStrSTD(dVB_TransDateDptCancel) & "'"
        Case tEN_TrnAddDepositCancel:
             tSql = "STP_CNxSQLSelDepositHDByWhr @ptTmnNum ='" & tVB_TerminalNoDptCancel & "',@ptRunning ='" & tVB_TransNoDptCancel & "',@ptTransType ='" & tVB_TransTypeDptCancel & "',@ptTransDate ='" & SP_DTEtStrSTD(dVB_TransDateDptCancel) & "'"
             If SP_SQLvExecute(tSql, orsHD, oDB.ocnOnLine) <> 0 Then Exit Sub 'PH 2.0.0 อ่านข้อมูลจากฐานข้อมูล Server เท่านั้น
             If Not (orsHD.BOF Or orsHD.EOF) Then
                    'เอกสาร Deposit
                     tTmnNumTmp = SP_FEDtChkString(orsHD, "FTShdDepRefTmnNum")
                     tTrnNoTmp = SP_FEDtChkString(orsHD, "FTShdDepReTransNo")
                     tTrnTypeTmp = "06"
                     dTrnDateTmp = SP_FEDdChkDate(orsHD, "FDShdDepReTransDate")
                     dDateTime = SP_FEDdChkDate(orsHD, "FDShdSysDate")
                     tTime = SP_FEDtChkString(orsHD, "FTShdSysTime")
                     
                    tW_TmnNumTmp = tTmnNumTmp
                    tW_TrnNoTmp = tTrnNoTmp
                    tW_TrnTypeTmp = tTrnTypeTmp
                    dW_TrnDateTmp = dTrnDateTmp
                    dW_DateTime = dDateTime
                    tW_Time = tTime
                     
                    Call W_GETxLastAddDepositEx(tTmnNumTmp, tTrnNoTmp, dTrnDateTmp, tTrnTypeTmp, tVB_TerminalNoAddDptLast, tVB_TransNoAddDptLast, dVB_TransDateAddDptLast, tVB_TransTypeAddDptLast, dDateTime)
                    tSql = "STP_CNxSQLSelDepositHDByWhr @ptTmnNum ='" & tVB_TerminalNoAddDptLast & "',@ptRunning ='" & tVB_TransNoAddDptLast & "',@ptTransType ='" & tVB_TransTypeAddDptLast & "',@ptTransDate ='" & SP_DTEtStrSTD(dVB_TransDateAddDptLast) & "'"
             End If
        Case tEN_TrnAddDeposit:
            tSql = "STP_CNxSQLSelDepositHDByWhr @ptTmnNum ='" & tVB_TerminalNoAddDpt & "',@ptRunning ='" & tVB_TransNoAddDpt & "',@ptTransType ='" & tVB_TransTypeAddDpt & "',@ptTransDate ='" & SP_DTEtStrSTD(dVB_TransDateAddDpt) & "'"
    End Select

    If SP_SQLvExecute(tSql, orsHD, oDB.ocnOnLine) <> 0 Then Exit Sub 'PH 2.0.0 อ่านข้อมูลจากฐานข้อมูล Server เท่านั้น
    If Not (orsHD.BOF Or orsHD.EOF) Then
        
        Select Case Val(Format(tVB_TransactionSub))
            Case tEN_TrnRedeem:
                 
                Call SP_GETxSumDepositAmt(tVB_TerminalNoRedeem, tVB_TransNoRedeem, tVB_TransTypeRedeem, dVB_TransDateRedeem, cVB_DepossitAmtBAddDpt, cVB_DepossitVatBAddDpt) '*KT 51-09-22 เก็บค่า Deposit ก่อนหน้าและค่า Vat ของ Deposit ก่อนหน้า
                cVB_AccBalRedeem = SP_FEDvChkNumeric(orsHD, "FCShdDepositBal")
                cVB_DepossitAmtRedeemRnd = SP_FEDvChkNumeric(orsHD, "FCShdRnd")
                cVB_SunTotalAllRedeem = SP_FEDvChkNumeric(orsHD, "FCShdGrand") '+ cVB_DepossitAmtRedeemRnd '*KT 51-11-09
                cVB_SumTotalAllDptCancel = SP_FEDvChkNumeric(orsHD, "FCShdGrand") '- cVB_DepossitAmtRedeemRnd '*Eaw 56-07-20 CommSheet TK-ISS3000-287 TR-FAPL-067.xls
                cVB_SunTotalAllRedeemBDisc = SP_FEDvChkNumeric(orsHD, "FCShdB4DisChg")
                tVB_DepossitNoRedeem = SP_FEDtChkString(orsHD, "FTCdpDepositNo")
                cVB_DepossitAmtRedeem = cVB_DepossitAmtBAddDpt 'SP_FEDvChkNumeric(orsHD, "FCShdDepositAmt") + cVB_DepossitAmtRedeemRnd
                '----
                otbDepositNo.Text = tVB_DepossitNoRedeem
                otbDepositAmt.Text = SP_FMTcCurPoint(cVB_DepossitAmtRedeem, False, True)
                otbAccBal.Text = SP_FMTcCurPoint(IIf(cVB_AccBalRedeem < 0, 0, cVB_AccBalRedeem), False, True)
            Case tEN_TrnDepositCancel:
                cVB_AccBalDptCancel = SP_FEDvChkNumeric(orsHD, "FCShdDepositBal")
                cVB_DepossitAmtDptCancelRnd = SP_FEDvChkNumeric(orsHD, "FCShdRnd")
                cVB_SumTotalAllDptCancel = SP_FEDvChkNumeric(orsHD, "FCShdGrand") '+ cVB_DepossitAmtDptCancelRnd  '*KT 51-11-09
                cVB_SumTotalAllDptCancelBDis = SP_FEDvChkNumeric(orsHD, "FCShdB4DisChg")
                tVB_DepossitNoDptCancel = SP_FEDtChkString(orsHD, "FTCdpDepositNo")
                cVB_DepossitAmtDptCancel = SP_GETcSumDepositAmt(tVB_TerminalNoDptCancel, tVB_TransNoDptCancel, tVB_TransTypeDptCancel, dVB_TransDateDptCancel)  ' SP_FEDvChkNumeric(orsHD, "FCShdDepositAmt") '*KT 51-10-01 Communication Sheet ML-SF-568 Deposit Is Wrong.xls
                '----
                otbDepositNo.Text = tVB_DepossitNoDptCancel
                otbDepositAmt.Text = SP_FMTcCurPoint(cVB_DepossitAmtDptCancel, False, True)  '*KT 51-11-09
                otbAccBal.Text = SP_FMTcCurPoint(IIf(cVB_AccBalDptCancel < 0, 0, cVB_AccBalDptCancel), False, True)
            Case tEN_TrnAddDepositCancel:
                cVB_AccBalDptCancel = SP_FEDvChkNumeric(orsHD, "FCShdDepositBal")
                cVB_DepossitAmtDptCancelRnd = SP_FEDvChkNumeric(orsHD, "FCShdRnd")
                cVB_SumTotalAllDptCancel = SP_FEDvChkNumeric(orsHD, "FCShdGrand") '+ cVB_DepossitAmtDptCancelRnd  '*KT 51-11-09
                cVB_SumTotalAllDptCancelBDis = SP_FEDvChkNumeric(orsHD, "FCShdB4DisChg")
                tVB_DepossitNoDptCancel = SP_FEDtChkString(orsHD, "FTCdpDepositNo")
                If tVB_TerminalNoDptCancel = tVB_TerminalNoAddDptLast And tVB_TransNoDptCancel = tVB_TransNoAddDptLast And tVB_TransTypeDptCancel = tVB_TransTypeAddDptLast And SP_DTEtStrSTD(dVB_TransDateDptCancel) = SP_DTEtStrSTD(dVB_TransDateAddDptLast) Then
                        bVB_CNAddDpsCLast = True
                Else
                        bVB_CNAddDpsCLast = False
                End If
                If bVB_CNAddDpsCLast Then
'                    cVB_DepossitAmtDptCancel = SP_FMTcCurPoint(SP_VALcText2Double(SP_FEDvChkNumeric(orsHD, "FCShdDepositAmt") + SP_FEDvChkNumeric(orsHD, "FCShdRnd")), True, False)
                    cVB_DepossitAmtDptCancel = SP_FMTcCurPoint(SP_VALcText2Double(SP_FEDvChkNumeric(orsHD, "FCShdDepositAmt")), True, False) '*Eaw 56-07-20 CommSheet TK-ISS3000-287 TR-FAPL-067.xls
                Else
                    cVB_DepossitAmtDptCancel = W_GETcAddDepositCancelAmt(tTmnNumTmp, tTrnNoTmp, dTrnDateTmp, tTrnTypeTmp, tVB_TerminalNoAddDptLast, tVB_TransNoAddDptLast, dVB_TransDateAddDptLast, tVB_TransTypeAddDptLast, dDateTime, tTime) 'SP_FEDvChkNumeric(orsHD, "FCShdDepositAmt") + SP_FEDvChkNumeric(orsHD, "FCShdRnd")
                End If
'                cVB_DepossitAmtDptCancel = cVB_DepossitAmtDptCancel - cVB_DepossitAmtDptCancelRnd '*Eaw 56-07-20 CommSheet TK-ISS3000-287 TR-FAPL-067.xls
                '----
                otbDepositNo.Text = tVB_DepossitNoDptCancel
                otbDepositAmt.Text = SP_FMTcCurPoint(cVB_DepossitAmtDptCancel, False, True)  'SP_FMTcCurPoint(cVB_DepossitAmtDptCancel + cVB_DepossitAmtDptCancelRnd, False, True)
                otbAccBal.Text = SP_FMTcCurPoint(cVB_AccBalDptCancel, False, True)
                
                Call W_CHKbDepositVatRate(tVB_DepossitNoDptCancel, "Add-Deposit Cancel") '*Bump 65-10-13 STK-POS ComSheet CR803 - 006.xlsx
                
            Case tEN_TrnAddDeposit:
                cVB_DepossitAmtBAddDpt = SP_GETcSumDepositAmt(tVB_TerminalNoAddDpt, tVB_TransNoAddDpt, tVB_TransTypeAddDpt, dVB_TransDateAddDpt) 'SP_FEDvChkNumeric(orsHD, "FCShdDepositAmt") '*KT 51-10-28 V 3.0.144/V 3.0.200
                cVB_AccBalAddDpt = SP_FEDvChkNumeric(orsHD, "FCShdDepositBal")
                cVB_DepossitAddDptRnd = SP_FEDvChkNumeric(orsHD, "FCShdRnd")
                cVB_SumTotalAllAddDpt = SP_FEDvChkNumeric(orsHD, "FCShdGrand") '+ cVB_DepossitAddDptRnd '*KT 51-11-09
                cVB_SumTotalAllAddDptBDisc = SP_FEDvChkNumeric(orsHD, "FCShdB4DisChg")
                tVB_DepossitNoAddDpt = SP_FEDtChkString(orsHD, "FTCdpDepositNo")
                '----
                otbDepositNo.Text = tVB_DepossitNoAddDpt
                otbDepositAmt.Text = SP_FMTcCurPoint(cVB_DepossitAmtBAddDpt, False, True)
                otbAccBal.Text = SP_FMTcCurPoint(cVB_AccBalAddDpt, False, True)
                
                Call W_CHKbDepositVatRate(tVB_DepossitNoAddDpt, "Add-Deposit") '*Bump 65-10-13 STK-POS ComSheet CR803 - 006.xlsx
                
        End Select
'        DoEvents '*KT 53-01-14 DoEvents
    Else
        ocmOk.Enabled = False
        
    End If
End Sub

Public Function W_PRCxInsRedeemDTCD() As Boolean
'--------------------------------------------------------
'Call:
'Cmt: การทำรายการ Insert DT และ CD of Redeem
'PH 2.0.0
'---------------------------------------------------------
    Dim orsTemp As ADODB.Recordset
    Dim tSql$, tTable$, tWhr$
    Dim tTmnNumTmp As String
    Dim tTrnNoTmp As String
    Dim tTrnTypeTmp As String, tName As String
    Dim dTrnDateTmp As Date
    Dim dDateTime As Date
On Error GoTo ErrHandle
    W_PRCxInsRedeemDTCD = False
    'DT
    tWhr = " WHERE  FTTmnNum=''" & tVB_CNTerminalNum & "'' AND FTShdTransNo=''" & tVB_Running & "''"
    tWhr = tWhr & " AND  FTShdTransType=''" & tVB_Transaction & "'' AND FDShdTransDate=''" & tUT_SaleDate & "''"
    tSql = "STP_CNxSQLFunctionDel @ptTable = '" & tVB_CNTblDT & "',@ptWhere = '" & tWhr & "'"
    Call SP_SQLvExecute(tSql) 'Local
    
    'find last insert DT
    tTable = tVB_CNTblDT
    Select Case Val(Format(tVB_TransactionSub))
        Case tEN_TrnRedeem:
            '*KT 51-11-12
            Call W_GETxLastAddDeposit(tVB_TerminalNoRedeem, tVB_TransNoRedeem, dVB_TransDateRedeem, tVB_TransTypeRedeem, tVB_TerminalNoAddDptLast, tVB_TransNoAddDptLast, dVB_TransDateAddDptLast, tVB_TransTypeAddDptLast) 'PH 2.0.0 อ่านข้อมูลจากฐานข้อมูล Server เท่านั้น
            tSql = "STP_CNxSQLSelDepositDTByWhr @ptTmnNum ='" & tVB_TerminalNoAddDptLast & "',@ptRunning ='" & tVB_TransNoAddDptLast & "',@ptTransType ='" & tVB_TransTypeAddDptLast & "',@ptTransDate ='" & SP_DTEtStrSTD(dVB_TransDateAddDptLast) & "'"
        Case tEN_TrnDepositCancel:
           '*KT 51-11-13
            tSql = "STP_CNxSQLSelDepositHDByWhr @ptTmnNum ='" & tVB_TerminalNoDptCancel & "',@ptRunning ='" & tVB_TransNoDptCancel & "',@ptTransType ='" & tVB_TransTypeDptCancel & "',@ptTransDate ='" & SP_DTEtStrSTD(dVB_TransDateDptCancel) & "'"
            If SP_SQLvExecute(tSql, orsHD, oDB.ocnOnLine) <> 0 Then Exit Function 'PH 2.0.0 อ่านข้อมูลจากฐานข้อมูล Server เท่านั้น
            If Not (orsHD.BOF Or orsHD.EOF) Then
                    'เอกสาร Deposit
                     tTmnNumTmp = SP_FEDtChkString(orsHD, "FTShdDepRefTmnNum")
                     tTrnNoTmp = SP_FEDtChkString(orsHD, "FTShdDepReTransNo")
                     tTrnTypeTmp = "06"
                     dTrnDateTmp = SP_FEDdChkDate(orsHD, "FDShdDepReTransDate")
                     dDateTime = SP_FEDdChkDate(orsHD, "FDShdSysDate")
                     Call W_GETxLastAddDepositEx(tTmnNumTmp, tTrnNoTmp, dTrnDateTmp, tTrnTypeTmp, tVB_TerminalNoAddDptLast, tVB_TransNoAddDptLast, dVB_TransDateAddDptLast, tVB_TransTypeAddDptLast, dDateTime)
                     tSql = "STP_CNxSQLSelDepositDTByWhr @ptTmnNum ='" & tVB_TerminalNoAddDptLast & "',@ptRunning ='" & tVB_TransNoAddDptLast & "',@ptTransType ='" & tVB_TransTypeAddDptLast & "',@ptTransDate ='" & SP_DTEtStrSTD(dVB_TransDateAddDptLast) & "'"
            End If
        Case tEN_TrnAddDepositCancel:
            '*KT 51-11-13
            tSql = "STP_CNxSQLSelDepositHDByWhr @ptTmnNum ='" & tVB_TerminalNoDptCancel & "',@ptRunning ='" & tVB_TransNoDptCancel & "',@ptTransType ='" & tVB_TransTypeDptCancel & "',@ptTransDate ='" & SP_DTEtStrSTD(dVB_TransDateDptCancel) & "'"
            If SP_SQLvExecute(tSql, orsHD, oDB.ocnOnLine) <> 0 Then Exit Function 'PH 2.0.0 อ่านข้อมูลจากฐานข้อมูล Server เท่านั้น
            If Not (orsHD.BOF Or orsHD.EOF) Then
                    'เอกสาร Deposit
                     tTmnNumTmp = SP_FEDtChkString(orsHD, "FTShdDepRefTmnNum")
                     tTrnNoTmp = SP_FEDtChkString(orsHD, "FTShdDepReTransNo")
                     tTrnTypeTmp = "06"
                     dTrnDateTmp = SP_FEDdChkDate(orsHD, "FDShdDepReTransDate")
                     dDateTime = SP_FEDdChkDate(orsHD, "FDShdSysDate")
                     
                     Call W_GETxLastAddDepositEx(tTmnNumTmp, tTrnNoTmp, dTrnDateTmp, tTrnTypeTmp, tVB_TerminalNoAddDptLast, tVB_TransNoAddDptLast, dVB_TransDateAddDptLast, tVB_TransTypeAddDptLast, dDateTime)
                     '*KT 51-11-18 ตรวจสอบว่าใบที่อ้างอิงการทำ add deposit cancel คือ seq  สุดท้าย
                     '-----------------------------------------------------------------------------------------
                     If tVB_TerminalNoDptCancel = tVB_TerminalNoAddDptLast And tVB_TransNoDptCancel = tVB_TransNoAddDptLast And tVB_TransTypeDptCancel = tVB_TransTypeAddDptLast And SP_DTEtStrSTD(dVB_TransDateDptCancel) = SP_DTEtStrSTD(dVB_TransDateAddDptLast) Then
                        bVB_CNAddDpsCLast = True
                    Else
                        bVB_CNAddDpsCLast = False
                     End If
                     '------------------------------------------------------------------------------------------
                     tSql = "STP_CNxSQLSelDepositDTByWhr @ptTmnNum ='" & tVB_TerminalNoAddDptLast & "',@ptRunning ='" & tVB_TransNoAddDptLast & "',@ptTransType ='" & tVB_TransTypeAddDptLast & "',@ptTransDate ='" & SP_DTEtStrSTD(dVB_TransDateAddDptLast) & "'"
            End If
        Case tEN_TrnAddDeposit:
             '*KT 51-11-12
            Call W_GETxLastAddDeposit(tVB_TerminalNoAddDpt, tVB_TransNoAddDpt, dVB_TransDateAddDpt, tVB_TransTypeAddDpt, tVB_TerminalNoAddDptLast, tVB_TransNoAddDptLast, dVB_TransDateAddDptLast, tVB_TransTypeAddDptLast)
            tSql = "STP_CNxSQLSelDepositDTByWhr @ptTmnNum ='" & tVB_TerminalNoAddDptLast & "',@ptRunning ='" & tVB_TransNoAddDptLast & "',@ptTransType ='" & tVB_TransTypeAddDptLast & "',@ptTransDate ='" & SP_DTEtStrSTD(dVB_TransDateAddDptLast) & "'"
    End Select
    
    If SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine) <> 0 Then Call SP_DATxRsNothing(orsTemp): Exit Function 'PH 2.0.0 อ่านข้อมูลจากฐานข้อมูล Server เท่านั้น
    
    'insert into grid     อาจจะไม่ถูกนัก ถ้าแยกเอา promotion แยก table ออกมาต่างหาก, หรือไม่ต้อง return table promotion ด้วย
    Do While Not orsTemp.EOF
            Call wFunctionKB.W_DupTblBtwSleNHold(orsTemp, tTable, 1, False)           'ทำให้ return แล้ว ต้องไม่มีรายการยกเลิก
            '*KT 51-11-12
            '--------------------------------------------------------------------------------------------------------------
            If SP_FEDtChkString(orsTemp, "FTSkuStaMallEmpCard") = "N" Then
                tSql = "UPDATE " & tVB_CNTblDT
                tSql = tSql & vbCrLf & "SET FCSdtBSubReg=0"
                tSql = tSql & vbCrLf & ",FCSdtBSubPro=FCSdtSaleAmt"
                tSql = tSql & vbCrLf & ",FCSdtBluePrice=0"
                tSql = tSql & vbCrLf & ",FCSdtYellowPrice=FCSdtSaleAmt"
                tSql = tSql & vbCrLf & ",FCSdtCostEX=0"
                tSql = tSql & vbCrLf & " WHERE  FTTmnNum='" & SP_FEDtChkString(orsTemp, "FTTmnNum") & "' AND FTShdTransNo='" & tVB_Running & "'"
                tSql = tSql & vbCrLf & " AND  FTShdTransType='" & SP_FEDtChkString(orsTemp, "FTShdTransType") & "' AND FDShdTransDate='" & Format(SP_FEDdChkDate(orsTemp, "FDShdTransDate"), "yyyy/mm/dd") & "'"
                If SP_FEDtChkString(orsTemp, "FTSdtWeight") = "" Then 'สินค้าปกติ
                    tSql = tSql & vbCrLf & " AND  FTSdtBarCode='" & SP_FEDtChkString(orsTemp, "FTSdtBarCode") & "' AND FNSdtSeqNo=" & SP_FEDvChkNumeric(orsTemp, "FNSdtSeqNo")
                Else  'สินค้าน้ำหนัก
                    tSql = tSql & vbCrLf & " AND  FTSdtInstoreBar='" & SP_FEDtChkString(orsTemp, "FTSdtBarCode") & "' AND FNSdtSeqNo=" & SP_FEDvChkNumeric(orsTemp, "FNSdtSeqNo")
                End If
                Call SP_SQLvExecute(tSql)
            Else
                tSql = "UPDATE " & tVB_CNTblDT
                tSql = tSql & vbCrLf & "SET FCSdtBSubReg=FCSdtSaleAmt"
                tSql = tSql & vbCrLf & ",FCSdtBSubPro=0"
                 tSql = tSql & vbCrLf & ",FCSdtBluePrice=FCSdtSaleAmt"
                tSql = tSql & vbCrLf & ",FCSdtYellowPrice=0"
                tSql = tSql & vbCrLf & ",FCSdtCostEX=0" '*KT 51-10-21 V 3.0.140/V 3.0.200
                tSql = tSql & vbCrLf & " WHERE  FTTmnNum='" & SP_FEDtChkString(orsTemp, "FTTmnNum") & "' AND FTShdTransNo='" & tVB_Running & "'"
                tSql = tSql & vbCrLf & " AND  FTShdTransType='" & SP_FEDtChkString(orsTemp, "FTShdTransType") & "' AND FDShdTransDate='" & Format(SP_FEDdChkDate(orsTemp, "FDShdTransDate"), "yyyy/mm/dd") & "'"
                If SP_FEDtChkString(orsTemp, "FTSdtWeight") = "" Then 'สินค้าปกติ
                    tSql = tSql & vbCrLf & " AND  FTSdtBarCode='" & SP_FEDtChkString(orsTemp, "FTSdtBarCode") & "' AND FNSdtSeqNo=" & SP_FEDvChkNumeric(orsTemp, "FNSdtSeqNo")
                Else 'สินค้าน้ำหนัก
                    tSql = tSql & vbCrLf & " AND  FTSdtInstoreBar='" & SP_FEDtChkString(orsTemp, "FTSdtBarCode") & "' AND FNSdtSeqNo=" & SP_FEDvChkNumeric(orsTemp, "FNSdtSeqNo")
                End If
                Call SP_SQLvExecute(tSql)
            End If
            '----------------------------------------------------------------------------------------------------------------------
            orsTemp.MoveNext
    Loop
    
    tName = SP_STRtSiggleCode(tVB_CNUserName)
     
    tSql = "UPDATE " & tVB_CNTblDT
    tSql = tSql & vbCrLf & "SET FDDateUpd = '" & UT_tStringDate(Date) & "'"                                    '*Ao 2015/01/09
    tSql = tSql & vbCrLf & ",FTTimeUpd = '" & Format(Time, "HH:MM:SS") & "'"
    tSql = tSql & vbCrLf & ",FTWhoUpd = '" & tName & "'"
    tSql = tSql & vbCrLf & ",FDDateIns = '" & UT_tStringDate(Date) & "'"
    tSql = tSql & vbCrLf & ",FTTimeIns = '" & Format(Time, "HH:MM:SS") & "'"
    tSql = tSql & vbCrLf & ",FTWhoIns = '" & tName & "'"
    tSql = tSql & vbCrLf & "WHERE  FTTmnNum = '" & tVB_CNTerminalNum & "' AND FTShdTransNo = '" & tVB_Running & "'"
    tSql = tSql & vbCrLf & "AND  FTShdTransType = '" & tVB_Transaction & "' AND FDShdTransDate = '" & tUT_SaleDate & "'"
    Call SP_SQLvExecute(tSql)
    
    'CD
    tWhr = " WHERE  FTTmnNum=''" & tVB_CNTerminalNum & "'' AND FTShdTransNo=''" & tVB_Running & "''"
    tWhr = tWhr & " AND  FTShdTransType=''" & tVB_Transaction & "'' AND FDShdTransDate=''" & tUT_SaleDate & "''"
    tSql = "STP_CNxSQLFunctionDel @ptTable = '" & tVB_CNTblCD & "',@ptWhere = '" & tWhr & "'"
    Call SP_SQLvExecute(tSql)
    
    'find last insert CD
    tTable = tVB_CNTblCD
    Select Case Val(Format(tVB_TransactionSub))
        Case tEN_TrnRedeem:
             '*KT 51-11-12
            tSql = "STP_CNxSQLSelDepositCDByWhr @ptTmnNum ='" & tVB_TerminalNoAddDptLast & "',@ptRunning ='" & tVB_TransNoAddDptLast & "',@ptTransType ='" & tVB_TransTypeAddDptLast & "',@ptTransDate ='" & SP_DTEtStrSTD(dVB_TransDateAddDptLast) & "'"
        Case tEN_TrnDepositCancel:
             '*KT 51-11-12
            tSql = "STP_CNxSQLSelDepositCDByWhr @ptTmnNum ='" & tVB_TerminalNoAddDptLast & "',@ptRunning ='" & tVB_TransNoAddDptLast & "',@ptTransType ='" & tVB_TransTypeAddDptLast & "',@ptTransDate ='" & SP_DTEtStrSTD(dVB_TransDateAddDptLast) & "'"
        Case tEN_TrnAddDepositCancel:
            tSql = "STP_CNxSQLSelDepositCDByWhr @ptTmnNum ='" & tVB_TerminalNoAddDptLast & "',@ptRunning ='" & tVB_TransNoAddDptLast & "',@ptTransType ='" & tVB_TransTypeAddDptLast & "',@ptTransDate ='" & SP_DTEtStrSTD(dVB_TransDateAddDptLast) & "'"
        Case tEN_TrnAddDeposit:
            '*KT 51-11-12
            tSql = "STP_CNxSQLSelDepositCDByWhr @ptTmnNum ='" & tVB_TerminalNoAddDptLast & "',@ptRunning ='" & tVB_TransNoAddDptLast & "',@ptTransType ='" & tVB_TransTypeAddDptLast & "',@ptTransDate ='" & SP_DTEtStrSTD(dVB_TransDateAddDptLast) & "'"
    End Select
    
    If tSql <> "" Then
        If SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine) <> 0 Then Call SP_DATxRsNothing(orsTemp): Exit Function 'PH 2.0.0 อ่านข้อมูลจากฐานข้อมูล Server เท่านั้น
        Do While Not orsTemp.EOF
            Call wFunctionKB.W_DupTblBtwSleNHold(orsTemp, tTable, 1, False)
            orsTemp.MoveNext
        Loop
        
        If Not (orsTemp.BOF And orsTemp.EOF) Then               '*Ao 2015/01/09
            tSql = "UPDATE " & tVB_CNTblCD
            tSql = tSql & vbCrLf & "SET FDDateUpd = '" & UT_tStringDate(Date) & "'"                                    '*Ao 2015/01/09
            tSql = tSql & vbCrLf & ",FTTimeUpd = '" & Format(Time, "HH:MM:SS") & "'"
            tSql = tSql & vbCrLf & ",FTWhoUpd = '" & tName & "'"
            tSql = tSql & vbCrLf & ",FDDateIns = '" & UT_tStringDate(Date) & "'"
            tSql = tSql & vbCrLf & ",FTTimeIns = '" & Format(Time, "HH:MM:SS") & "'"
            tSql = tSql & vbCrLf & ",FTWhoIns = '" & tName & "'"
            tSql = tSql & vbCrLf & "WHERE  FTTmnNum = '" & tVB_CNTerminalNum & "' AND FTShdTransNo = '" & tVB_Running & "'"
            tSql = tSql & vbCrLf & "AND  FTShdTransType = '" & tVB_Transaction & "' AND FDShdTransDate = '" & tUT_SaleDate & "'"
            Call SP_SQLvExecute(tSql)
        End If
        
    End If
    Call SP_DATxRsClose(orsTemp)
    
    'Point
    '*KT 55-07-09 V.4.0.21  เก็บข้อมลตาราง TPSTSalePoint เมื่อทำ DepositCancel, AddDepositCancel
'    Dim tWhe$
'    Dim orsTMPHD  As ADODB.Recordset
'    tTable = "TPSTSalePoint"
'    Select Case Val(Format(tVB_TransactionSub))                            '*Ao 2013/03/13   TK POINT      'ไม่ใช้แล้ว
'        Case tEN_TrnDepositCancel
'            tWhe = "  WHERE ((FTTmnNum='" & tVB_TerminalNoDptCancel & "'"
'            tWhe = tWhe & "  AND FTShdTransNo='" & tVB_TransNoDptCancel & "'"
'            tWhe = tWhe & "  AND FTShdTransType='" & tVB_TransTypeDptCancel & "'"
'            tWhe = tWhe & "  AND FDShdTransDate ='" & SP_DTEtStrSTD(dVB_TransDateDptCancel) & "')"
'            tWhe = tWhe & "  OR "
'            tWhe = tWhe & "   (FTShdDepRefTmnNum='" & tVB_TerminalNoDptCancel & "'"
'            tWhe = tWhe & "  AND FTShdDepReTransNo='" & tVB_TransNoDptCancel & "'"
'            tWhe = tWhe & "  AND FDShdDepReTransDate='" & SP_DTEtStrSTD(dVB_TransDateDptCancel) & "')"
'            tWhe = tWhe & "  AND FTShdTransType IN ('" & Format(tEN_TrnAddDeposit, "00") & "'))"
'            tWhe = tWhe & "  AND ISNULL(FTShdStaCancel,'N') <>'Y'"
'             tWhe = tWhe & "  ORDER BY  FDShdSysDate DESC,FDShdTransDate DESC,FTShdTransNo DESC"
'        Case tEN_TrnAddDepositCancel
'            tWhe = " WHERE    ((FTShdDepReTransNo = '" & tW_TrnNoTmp & "') AND (FTShdDepRefTmnNum = '" & tW_TmnNumTmp & "')"
'            tWhe = tWhe & "  AND (FDShdDepReTransDate = '" & Format(dW_TrnDateTmp, "yyyy/mm/dd") & "'))"
'            tWhe = tWhe & "  AND ((FDShdSysDate > '" & Format(dW_DateTime, "yyyy/mm/dd") & "') OR (FDShdSysDate = '" & Format(dW_DateTime, "yyyy/mm/dd") & "' AND FTShdSysTime >='" & tW_Time & "'))"
'            tWhe = tWhe & "  AND (ISNULL(FTShdStaCancel,'N')<> 'Y')"
'            tWhe = tWhe & "  AND (FTShdTransType IN ('" & Format(tEN_TrnAddDeposit, "00") & "'))"
'            tWhe = tWhe & "  ORDER BY  FDShdSysDate DESC,FDShdTransDate DESC,FTShdTransNo DESC"
'    End Select
    
'    Select Case Val(Format(tVB_TransactionSub))
'        Case tEN_TrnDepositCancel, tEN_TrnAddDepositCancel                         '*Ao 2013/03/13   TK POINT      'ไม่ใช้แล้ว
'            tSql = "SELECT FTTmnNum,FTShdTransNo,FTShdTransType,FDShdTransDate,FDShdSysDate  FROM TPSTDepositHD " & tWhe
'            If SP_SQLvExecute(tSql, orsTMPHD, oDB.ocnOnLine) = 0 Then
'                 Do While Not orsTMPHD.EOF
'
'                         tWhe = "  WHERE (FTTmnNum='" & SP_FEDtChkString(orsTMPHD, "FTTmnNum") & "'"
'                         tWhe = tWhe & "  AND FTShdTransNo='" & SP_FEDtChkString(orsTMPHD, "FTShdTransNo") & "'"
'                         tWhe = tWhe & "  AND FTShdTransType='" & SP_FEDtChkString(orsTMPHD, "FTShdTransType") & "'"
'                         tWhe = tWhe & "  AND FDShdTransDate ='" & SP_DTEtStrSTD(SP_FEDdChkDate(orsTMPHD, "FDShdTransDate")) & "')"
'
'                         tSql = "SELECT * FROM TPSTSalePoint " & tWhe
'                         If SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine) = 0 Then
'                              Do While Not orsTemp.EOF
'                                      Call wFunctionKB.W_DupTblBtwSleNHoldPoint(orsTemp, tTable, 1, False)
'                                      orsTemp.MoveNext
'                              Loop
'                         End If
'                         orsTMPHD.MoveNext
'                Loop
'            End If
'    End Select
    
    'RC
    Select Case Val(Format(tVB_TransactionSub))
        Case tEN_TrnDepositCancel, tEN_TrnAddDepositCancel
           
    End Select
    
    
    
    Call SP_DATxRsNothing(orsHD)
    Call SP_DATxRsNothing(orsTemp)
    W_PRCxInsRedeemDTCD = True
    Exit Function
ErrHandle:
    Call SP_DATxRsNothing(orsHD)
    Call SP_DATxRsNothing(orsTemp)
End Function

'Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'    If UnloadMode = 0 Then Cancel = 1
'End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrHandle
    Call SP_DATxRsNothing(orsHD)
'    Set wCNRedeemAmt = Nothing
    Call SP_OBJxSetNothing(Me)
ErrHandle:
    
End Sub

Private Sub ocmCancel_Click()
    Select Case Val(Format(tVB_TransactionSub))
        Case tEN_TrnRedeem:
            bVB_TrnRedeem = False
        Case tEN_TrnDepositCancel, tEN_TrnAddDepositCancel: '*KT 51-10-08 V 3.0.131/V 3.0.200
            bVB_TrnDepositCancel = False
        Case tEN_TrnAddDeposit
            bVB_TrnAddDeposit = False
    End Select
    
    bW_CANNOT_do = False '*Bump 65-12-05 CommSheet TTK-ISS3000-814 TK ISSUE - Warning message was displayed and cannot proceed with the Layaway Pickup or Layawa.xlsx
    olaForm(3).Visible = False
    '========================
    Unload Me
End Sub

Private Sub ocmOK_Click()
    Dim tTable$

    If bW_CANNOT_do Then '*Bump 65-10-13 STK-POS ComSheet CR803 - 006.xlsx
        Dim tMsg As String
        tMsg = olaForm(3).Caption
        MsgBox tMsg, vbExclamation, "Layaway"
        'Call ocmCancel_Click
        Exit Sub
    End If
    
    tTable = tVB_CNTblHD
    Call wFunctionKB.W_DupTblBtwSleUpdHoldHD(orsHD, tTable, 1)
    If Not W_PRCxInsRedeemDTCD Then
        MsgBox "System not found data", vbCritical
        Call ocmCancel_Click
        Exit Sub
    End If
    Select Case Val(Format(tVB_TransactionSub))
        Case tEN_TrnRedeem:
            bVB_TrnRedeem = True
            Call W_UPDxPhoneNo(W_PRCtPhone) '*Eaw 56-10-28 CommSheet TK-ISS3000-401.xlsx Update Phone No. ให้เป็นเบอร์ล่าสุด
        Case tEN_TrnDepositCancel, tEN_TrnAddDepositCancel: '*KT 51-10-08 V 3.0.131/V 3.0.200
            bVB_TrnDepositCancel = True
            Call W_UPDxPhoneNo(W_PRCtPhone) '*Eaw 56-10-28 CommSheet TK-ISS3000-401.xlsx Update Phone No. ให้เป็นเบอร์ล่าสุด
        Case tEN_TrnAddDeposit
            bVB_TrnAddDeposit = True
            Call W_UPDxPhoneNo(W_PRCtPhone) '*Eaw 56-10-28 CommSheet TK-ISS3000-401.xlsx Update Phone No. ให้เป็นเบอร์ล่าสุด
    End Select
    Unload Me
End Sub

Private Sub W_GETxLastAddDeposit(ByVal ptTmnNo$, ByVal ptTrnNo$, ByVal pdTrnDate As Date, ByVal ptTrnType$, ByRef ptTmnNoLast$, ByRef ptTrnNoLast$, ByRef pdTrnDateLast As Date, ByRef ptTrnTypeLast$)
'__________________________
'   Call:
'   Cmt:   หาหมาเลขเครื่องและหมายเลข Ticket ล่าสุดของ Add Deposit
''*KT 51-11-13
'__________________________
    Dim tSql As String
    Dim orsTemp As ADODB.Recordset
    Dim tTbl$
    
On Error GoTo ErrHandle
    tTbl = "TPSTDepositHD"
    tSql = "SELECT  FTTmnNum,FTShdTransNo,FDShdTransDate,FTShdTransType"
    tSql = tSql & vbCrLf & "FROM " & tTbl
    tSql = tSql & vbCrLf & "WHERE      (((" & tTbl & ".FTShdDepReTransNo = '" & ptTrnNo & "') AND (" & tTbl & ".FTShdDepRefTmnNum = '" & ptTmnNo & "') AND"
    tSql = tSql & vbCrLf & "(" & tTbl & ".FDShdDepReTransDate = '" & Format(pdTrnDate, "yyyy/mm/dd") & "'))"
    tSql = tSql & vbCrLf & "OR "
    tSql = tSql & vbCrLf & "((" & tTbl & ".FTShdTransNo = '" & ptTrnNo & "') AND (" & tTbl & ".FTTmnNum = '" & ptTmnNo & "') AND"
    tSql = tSql & vbCrLf & "(" & tTbl & ".FDShdTransDate = '" & Format(pdTrnDate, "yyyy/mm/dd") & "')))"
    tSql = tSql & vbCrLf & "AND ISNULL(FTShdStaCancel,'N')<> 'Y'"
    tSql = tSql & vbCrLf & "AND FTShdTransType IN ('" & Format(tEN_TrnDeposit, "00") & "','" & Format(tEN_TrnAddDeposit, "00") & "')"
    tSql = tSql & vbCrLf & "ORDER BY FDShdSysDate DESC ,FDShdTransDate DESC,FTShdTransNo DESC"
    If SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine) <> 0 Then 'PH 2.0.0 อ่านข้อมูลจากฐานข้อมูล Server เท่านั้น
        ptTmnNoLast = ptTmnNo
        ptTrnNoLast = ptTrnNo
        pdTrnDateLast = pdTrnDate
        ptTrnTypeLast = ptTrnType
        Call SP_DATxRsNothing(orsTemp)
        Exit Sub
    End If
    If Not (orsTemp.BOF Or orsTemp.EOF) Then
        ptTmnNoLast = SP_FEDtChkString(orsTemp, "FTTmnNum")
        ptTrnNoLast = SP_FEDtChkString(orsTemp, "FTShdTransNo")
        pdTrnDateLast = SP_FEDdChkDate(orsTemp, "FDShdTransDate")
        ptTrnTypeLast = SP_FEDtChkString(orsTemp, "FTShdTransType")
    Else
        ptTmnNoLast = ptTmnNo
        ptTrnNoLast = ptTrnNo
        pdTrnDateLast = pdTrnDate
        ptTrnTypeLast = ptTrnType
    End If
    Call SP_DATxRsNothing(orsTemp)
    Exit Sub
ErrHandle:
    Call SP_DATxRsNothing(orsTemp)
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Sub

Private Sub W_GETxLastAddDepositEx(ByVal ptTmnNo$, ByVal ptTrnNo$, ByVal pdTrnDate As Date, ByVal ptTrnType$, ByRef ptTmnNoLast$, ByRef ptTrnNoLast$, ByRef pdTrnDateLast As Date, ByRef ptTrnTypeLast$, ByVal pdDateTime As Date)
'__________________________
'   Call:
'   Cmt:   หาหมาเลขเครื่องและหมายเลข Ticket ล่าสุดของ Add Deposit ที่หลังจากเอกสาร Add Deposit ที่อ้างอิงการ cancel
'*KT 51-11-13
'__________________________
    Dim tSql As String
    Dim orsTemp As ADODB.Recordset
    Dim tTbl$
    
On Error GoTo ErrHandle
    tTbl = "TPSTDepositHD"
    tSql = "SELECT  FTTmnNum,FTShdTransNo,FDShdTransDate,FTShdTransType,FDShdSysDate"
    tSql = tSql & vbCrLf & "FROM " & tTbl
    tSql = tSql & vbCrLf & "WHERE    ((" & tTbl & ".FTShdDepReTransNo = '" & ptTrnNo & "') AND (" & tTbl & ".FTShdDepRefTmnNum = '" & ptTmnNo & "')"
    tSql = tSql & vbCrLf & "AND (" & tTbl & ".FDShdDepReTransDate = '" & Format(pdTrnDate, "yyyy/mm/dd") & "')"
    tSql = tSql & vbCrLf & "AND (" & tTbl & ".FDShdSysDate >= '" & Format(pdDateTime, "yyyy/mm/dd") & "')"
    tSql = tSql & vbCrLf & "AND (ISNULL(FTShdStaCancel,'N')<> 'Y')"
    tSql = tSql & vbCrLf & "AND (FTShdTransType IN ('" & Format(tEN_TrnDeposit, "00") & "','" & Format(tEN_TrnAddDeposit, "00") & "')))"   '*Ao 2013/02/19
'    tSql = tSql & vbCrLf & "ORDER BY  FDShdSysDate DESC,FDShdTransDate DESC,FTShdTransNo DESC"             '*Ao 2013/02/19
    tSql = tSql & vbCrLf & "ORDER BY  FDShdSysDate DESC,FDShdTransDate DESC,FTShdTransType DESC,FCShdDepositBal"             '*Ao 2013/02/19
    If SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine) <> 0 Then 'PH 2.0.0 อ่านข้อมูลจากฐานข้อมูล Server เท่านั้น
        ptTmnNoLast = ptTmnNo
        ptTrnNoLast = ptTrnNo
        pdTrnDateLast = pdTrnDate
        ptTrnTypeLast = ptTrnType
        Call SP_DATxRsNothing(orsTemp)
        Exit Sub
    End If
    If Not (orsTemp.BOF Or orsTemp.EOF) Then
        ptTmnNoLast = SP_FEDtChkString(orsTemp, "FTTmnNum")
        ptTrnNoLast = SP_FEDtChkString(orsTemp, "FTShdTransNo")
        pdTrnDateLast = SP_FEDdChkDate(orsTemp, "FDShdTransDate")
        ptTrnTypeLast = SP_FEDtChkString(orsTemp, "FTShdTransType")
    Else
        ptTmnNoLast = ptTmnNo
        ptTrnNoLast = ptTrnNo
        pdTrnDateLast = pdTrnDate
        ptTrnTypeLast = ptTrnType
    End If
    Call SP_DATxRsNothing(orsTemp)
    Exit Sub
ErrHandle:
    Call SP_DATxRsNothing(orsTemp)
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Sub

Private Function W_GETcAddDepositCancelAmt(ByVal ptTmnNo$, ByVal ptTrnNo$, ByVal pdTrnDate As Date, ByVal ptTrnType$, ByRef ptTmnNoLast$, ByRef ptTrnNoLast$, ByRef pdTrnDateLast As Date, ByRef ptTrnTypeLast$, ByVal pdDateTime As Date, ByVal ptTime As String) As Double
'__________________________
'   Call:
'   Cmt:   หาหมาเลขเครื่องและหมายเลข Ticket ล่าสุดของ Add Deposit ที่หลังจากเอกสาร Add Deposit ที่อ้างอิงการ cancel
'*KT 51-11-13
'PH 2.0.0
'______________________
    Dim tSql As String
    Dim orsTemp As ADODB.Recordset
    Dim tTbl$
    
On Error GoTo ErrHandle
    tTbl = "TPSTDepositHD"
'    tSql = "SELECT  SUM(ISNULL(FCShdDepositAmt,0) + ISNULL(FCShdRnd,0)) AS FCSum"
    tSql = "SELECT  SUM(ISNULL(FCShdDepositAmt,0) ) AS FCSum" '*Eaw 56-07-24
    tSql = tSql & vbCrLf & "FROM " & tTbl
    tSql = tSql & vbCrLf & "WHERE    ((" & tTbl & ".FTShdDepReTransNo = '" & ptTrnNo & "') AND (" & tTbl & ".FTShdDepRefTmnNum = '" & ptTmnNo & "')"
    tSql = tSql & vbCrLf & "AND (" & tTbl & ".FDShdDepReTransDate = '" & Format(pdTrnDate, "yyyy/mm/dd") & "')"
    tSql = tSql & vbCrLf & "AND ((" & tTbl & ".FDShdSysDate > '" & Format(pdDateTime, "yyyy/mm/dd") & "') OR (" & tTbl & ".FDShdSysDate = '" & Format(pdDateTime, "yyyy/mm/dd") & "' AND " & tTbl & ".FTShdSysTime >='" & ptTime & "'))"
    tSql = tSql & vbCrLf & "AND (" & tTbl & ".FTShdTransType = '" & Format(tEN_TrnAddDeposit, "00") & "')"
    tSql = tSql & vbCrLf & "AND (ISNULL(FTShdStaCancel,'N')<> 'Y'))"
    If SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine) <> 0 Then 'PH 2.0.0 อ่านข้อมูลจากฐานข้อมูล Server เท่านั้น
        W_GETcAddDepositCancelAmt = 0
        Call SP_DATxRsNothing(orsTemp)
        Exit Function
    End If
    If Not (orsTemp.BOF Or orsTemp.EOF) Then
        W_GETcAddDepositCancelAmt = SP_FMTcCurPoint(SP_FEDvChkNumeric(orsTemp, "FCSum"), True, False)
    Else
        W_GETcAddDepositCancelAmt = 0
    End If
    Call SP_DATxRsNothing(orsTemp)
    Exit Function
ErrHandle:
    Call SP_DATxRsNothing(orsTemp)
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Function
Private Sub W_CHKbDepositVatRate(ByVal ptFTCdpDepositNo As String, ptLabel As String)
'*Bump 65-10-13 STK-POS ComSheet CR803 - 006.xlsx
On Error GoTo ErrHandle
    Dim tSql As String
    Dim orsTemp As ADODB.Recordset
    Dim dDepositVatRate As Double
    Dim dMasterVatRate As Double
    
    If Not bVB_Disallow_AddDeposit_OldVatRate Then Exit Sub
    
    bW_CANNOT_do = False
                    
    tSql = "SELECT DISTINCT TOP 1 DT.FCSdtTax AS FCDepositTaxRate" ',FCTaxRate"
    tSql = tSql & vbCrLf & "FROM TPSTDepositDT DT WITH(NOLOCK)INNER JOIN TPSTDepositHD HD WITH(NOLOCK)  ON (DT.FTShdTransNo=HD.FTShdTransNo"
    tSql = tSql & vbCrLf & " AND (DT.FTTmnNum = HD.FTTmnNum ) AND ( DT.FDShdTransDate = HD.FDShdTransDate)" '* Bump 66-02-08 [6.2302.1] [66/02/09]
    tSql = tSql & vbCrLf & "AND HD.FTCdpDepositNo= '" & ptFTCdpDepositNo & "')"
    tSql = tSql & vbCrLf & "WHERE (DT.FCSdtTax <> 0) " '* Bump 66-02-08 [6.2302.1] [66/02/09]
   ' tSql = tSql & vbCrLf & ",TCNMTaxMtn WITH(NOLOCK) " * Bump 66-02-08 [6.2302.1] [66/02/09]
   ' tSql = tSql & vbCrLf & "WHERE TCNMTaxMtn.FTTaxCode='01'" * Bump 66-02-08 [6.2302.1] [66/02/09]
    If SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine) = 0 Then 'PH 2.0.0 อ่านข้อมูลจากฐานข้อมูล Server เท่านั้น
        If Not (orsTemp.BOF Or orsTemp.EOF) Then
            dDepositVatRate = SP_FEDvChkNumeric(orsTemp, "FCDepositTaxRate")
            'dMasterVatRate = SP_FEDvChkNumeric(orsTemp, "FCTaxRate")
            Call SP_DATxRsNothing(orsTemp)
            
            If dDepositVatRate <> cVB_CNGSTRate Then 'dMasterVatRate Then
'                    olaForm(3).Caption = "CANNOT do Layaway " & ptLabel & vbNewLine & "because is tagged with old GST rate " & dDepositVatRate & "%" & vbNewLine & "NOT matching standard GST rate " & dMasterVatRate & "%."
                    olaForm(3).Caption = "CANNOT do Layaway " & ptLabel & vbNewLine & "because is tagged with old GST rate " & dDepositVatRate & "%" & vbNewLine & "NOT matching standard GST rate " & cVB_CNGSTRate & "%." '*Amorn 2023-02-15 Comsheet 819
                    ''- CANNOT do Layaway Add-Deposit and Add-Deposit Cancel to #~456 (because #~456 is tagged with old GST rate 8%; ie. NOT matching standard GST rate 9%)
                    olaForm(3).Visible = True
                    bW_CANNOT_do = True
                    Exit Sub
            End If
        End If
    End If
    Exit Sub
    '==================================
ErrHandle:
    bW_CANNOT_do = False
    olaForm(3).Visible = False
End Sub

