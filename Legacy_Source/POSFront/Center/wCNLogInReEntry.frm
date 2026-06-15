VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.1#0"; "Codejock.SuiteCtrls.Unicode.v11.1.3.ocx"
Begin VB.Form wCNLogInReEntry 
   Caption         =   "Re Entry;klg_Title"
   ClientHeight    =   5400
   ClientLeft      =   4110
   ClientTop       =   1590
   ClientWidth     =   6645
   ControlBox      =   0   'False
   Icon            =   "wCNLogInReEntry.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   6645
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   1200
      Index           =   2
      Left            =   0
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   -240
      Width           =   6645
      _Version        =   720897
      _ExtentX        =   11721
      _ExtentY        =   2117
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Re Entry;klg_Header"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   12
         Tag             =   "2;"
         Top             =   480
         Width           =   6255
      End
   End
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   3960
      Index           =   0
      Left            =   0
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   600
      Width           =   6645
      _Version        =   720897
      _ExtentX        =   11721
      _ExtentY        =   6985
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
      Begin VB.TextBox otbTmnNo 
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
         Left            =   3835
         MaxLength       =   5
         TabIndex        =   1
         Top             =   1320
         Width           =   2510
      End
      Begin VB.TextBox otbAccType 
         Alignment       =   2  'Center
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
         Left            =   5400
         MaxLength       =   1
         TabIndex        =   0
         Text            =   "1"
         Top             =   600
         Width           =   945
      End
      Begin MSComCtl2.DTPicker odtDate 
         Height          =   495
         Left            =   3840
         TabIndex        =   2
         Top             =   2040
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
         Format          =   172228611
         CurrentDate     =   39528
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Terminal Number;klg_2"
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
         Left            =   600
         TabIndex        =   9
         Tag             =   "2;"
         Top             =   1440
         Width           =   2970
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sale Date;klg_3"
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
         Left            =   600
         TabIndex        =   8
         Tag             =   "2;"
         Top             =   2160
         Width           =   1965
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
         Left            =   4200
         TabIndex        =   7
         Top             =   2640
         Width           =   1860
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Account(0), Normal(1);klg_1"
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
         Left            =   600
         TabIndex        =   6
         Tag             =   "2;"
         Top             =   720
         Width           =   3570
      End
   End
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   1200
      Index           =   1
      Left            =   0
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   4260
      Width           =   6645
      _Version        =   720897
      _ExtentX        =   11721
      _ExtentY        =   2117
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
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
         TabIndex        =   3
         Tag             =   "2;"
         Top             =   480
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
         Height          =   495
         Left            =   4920
         TabIndex        =   4
         Tag             =   "2;"
         Top             =   480
         Width           =   1455
      End
   End
End
Attribute VB_Name = "wCNLogInReEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private tW_AccType As String

Private Sub Form_Load()
    Call SP_FRMxChildCenter(Me)
    Call SP_FRMxGetCaptionFrmDB2Tag(Me)
    Call SP_FRMxSetCapForm(Me)
    Call W_FRMxInitail
    Call SP_STAxMouse(Me, False)
End Sub

Private Sub W_FRMxInitail()
    tW_AccType = "1"
    If bVB_ConfigDate Then
        If tVB_SelConfigDate = "0" Then
            olaForm(4).Caption = "(" & UCase(tCS_FormatDateM) & ")"
        Else
            olaForm(4).Caption = "(" & UCase(tCS_FormatDateSec) & ")"
        End If
    Else
        olaForm(4).Caption = "(" & UCase(tCS_FormatDateSec) & ")"
    End If
    Select Case tVB_CNMnuCurrent
        Case "1720"
            otbTmnNo.Text = tVB_CNTerminalNum
            otbTmnNo.Locked = True
            otbTmnNo.BackColor = nVB_CNClrBLock
            otbTmnNo.TabStop = False
            If bVB_ConfigDate Then
                If tVB_SelConfigDate = "0" Then
                    odtDate.CustomFormat = tCS_FormatDateM
                    odtDate.Value = Format(CDate(tUT_SaleDate), tCS_FormatDateM)
                Else
                    odtDate.CustomFormat = tCS_FormatDateSec
                    odtDate.Value = Format(CDate(tUT_SaleDate), tCS_FormatDateSec)
                End If
            Else
                odtDate.CustomFormat = tCS_FormatDateSec
                odtDate.Value = Format(CDate(tUT_SaleDate), tCS_FormatDateSec)
            End If
        Case "1721"
            otbTmnNo.Text = "" 'tVB_CNTerminalNum
            otbTmnNo.Locked = False
            otbTmnNo.BackColor = nVB_CNClrBUnLock
            otbTmnNo.TabStop = True
            If bVB_ConfigDate Then
                If tVB_SelConfigDate = "0" Then
                    odtDate.CustomFormat = tCS_FormatDateM
                    odtDate.Value = Format(CDate(tUT_SaleDate), tCS_FormatDateM)
                Else
                    odtDate.CustomFormat = tCS_FormatDateSec
                    odtDate.Value = Format(CDate(tUT_SaleDate), tCS_FormatDateSec)
                End If
            Else
                odtDate.CustomFormat = tCS_FormatDateSec
                odtDate.Value = Format(CDate(tUT_SaleDate), tCS_FormatDateSec)
            End If
'        Case "1722"
'            otbTmnNo.Text = tVB_CNTerminalNum
'            otbTmnNo.Locked = True
'            otbTmnNo.BackColor = nVB_CNClrBLock
'            otbTmnNo.TabStop = False
'            If bVB_ConfigDate Then
'                If tVB_SelConfigDate = "0" Then
'                    odtDate.CustomFormat = tCS_FormatDateM
'                    odtDate.Value = Format(Date - 1, tCS_FormatDateM)
'                Else
'                    odtDate.CustomFormat = tCS_FormatDateSec
'                    odtDate.Value = Format(Date - 1, tCS_FormatDateSec)
'                End If
'            Else
'                odtDate.CustomFormat = tCS_FormatDateSec
'                odtDate.Value = Format(Date - 1, tCS_FormatDateSec)
'            End If
    End Select
    

End Sub

Private Sub ocmCancel_Click()
    Unload Me
End Sub

Private Sub ocmOK_Click()
    Dim nMaxCode As Long
    Dim nMaxLogIn As Long
    Dim tMaxCode As String
    
    If W_PRCbCheckTerminalReEntry(tVB_CNMnuCurrent, Trim(otbTmnNo.Text)) Then
        Select Case tVB_CNMnuCurrent
                Case "1720"
                    If CDate(Format(CDate(odtDate.Value), "yyyy/MM/dd")) <= CDate(Format(CDate(tUT_SaleDate), "yyyy/MM/dd")) Then
                        '*Eaw 56-08-20 CommSheet TK-ISS3000-316-UAT_Req.xls พิมพ์ใบ Open Reentry
'                        tVB_CNTerminalNum = otbTmnNo.Text
'                        tUT_SaleDate = Format(CDate(odtDate.Value), "yyyy/MM/dd")
'                        tVB_CNReEntryDate = Format(CDate(odtDate.Value), "yyyy/MM/dd") '*Eaw 56-07-02 CommSheet TK-ISS3000-245-TR.xls
'                        tVB_CNAccTypeReEntryMode = otbAccType.Text
'                        Call W_CLRxClearDataHD_DT
'                        Call SaveSetting("AdaPosMall", "POSFront", "TmnNum", tVB_CNTerminalNum)
'                        Call SaveSetting("AdaPosMall", "POSFront", "Mode", "2")
'                        Call SaveSetting("AdaPosMall", "POSFront", "SaleDate", tUT_SaleDate)
'                        Call SaveSetting("AdaPosMall", "POSFront", "AccType", tVB_CNAccTypeReEntryMode)
'
'                        Call SP_INSxTransHD(tVB_CNTblHD)
'                        Call wFunctionKB.W_SetToComplete
'                        Call SP_UPDxSetCompleteTerminalMtn
                        '--------------------------------------------------------------------
                        tVB_CNReEntryIDMain = "1720"
                        tVB_CNMnuCurrent = "172001"
                        tVB_Transaction = aVB_TransMode(37).ID
                        tVB_TransactionSub = tVB_Transaction
                        tVB_TransName = aVB_TransMode(37).Name
                        tVB_TransNameABB = aVB_NameAbb(tEN_AbbSaleTaxInc - 1).UsrLangPri & ";" & aVB_NameAbb(tEN_AbbSaleTaxInc - 1).UsrLangSec
                        tVB_CNTypePrinter = aVB_TransMode(37).PrintType
                        tVB_CNTerminalNum = otbTmnNo.Text
                        tVB_CNVirTerminalNum = otbTmnNo.Text        '*Ao 2013/09/24
                        tVB_EJDate = tUT_SaleDate            '*Ao 2013/10/01       ReEntry  ให้เก็บ EJ ตาม Operation Date
                        tUT_SaleDate = Format(CDate(odtDate.Value), "yyyy/MM/dd")
                        tVB_CNReEntryDate = Format(CDate(odtDate.Value), "yyyy/MM/dd") '*Eaw 56-07-02 CommSheet TK-ISS3000-245-TR.xls
                        tVB_CNAccTypeReEntryMode = otbAccType.Text
                        Call W_CLRxClearDataHD_DT
                        Call SaveSetting("AdaPosMall", "POSFront", "TmnNum", tVB_CNTerminalNum)
                        Call SaveSetting("AdaPosMall", "POSFront", "Mode", "2")
                        Call SaveSetting("AdaPosMall", "POSFront", "SaleDate", tUT_SaleDate)
                        Call SaveSetting("AdaPosMall", "POSFront", "AccType", tVB_CNAccTypeReEntryMode)
                        Call SaveSetting("AdaPosMall", "POSFront", "VirtualPosNo", tVB_CNVirTerminalNum) '*Thirapong 64-06-09
                        nMaxCode = CLng(Val(Format(SP_SQLtFunction("MAX", "RIGHT(FTLogCode,5)", "TPSTSalHD", " WHERE Left(FTLogCode,9)='" & SP_String2YYMM(Date) & Format(tVB_CNTerminalNum, "00000") & "'", oDB.ocnOnLine, True))))
                        'เปรียบเทียบระหว่าง Master กับ transaction
                        If nMaxCode > nVB_CNLastLogin Then
                             nMaxLogIn = nMaxCode
                        Else
                            nMaxLogIn = nVB_CNLastLogin
                        End If
                        tMaxCode = Format(nMaxLogIn + 1, "00000")
                        nVB_CNLastLogin = nMaxLogIn + 1 '*KT 52-01-21 เก็บค่าไว้ในตัวแปร
                        tUT_LoginCode = SP_String2YYMM(Date) & Format(tVB_CNTerminalNum, "00000") & tMaxCode
                        bUT_SignOut = False
                    Else
                        Call SP_MSGnShowing("tms_CN005243", nCS_Error)
                        Call SP_OBJxSetFocus(odtDate)
                        Exit Sub
                    End If
                Case "1721"
                    If CDate(Format(CDate(odtDate.Value), "yyyy/MM/dd")) <= CDate(Format(CDate(tUT_SaleDate), "yyyy/MM/dd")) Then
                        '*Eaw 56-08-20 CommSheet TK-ISS3000-316-UAT_Req.xls พิมพ์ใบ Open Reentry
'                        tVB_CNTerminalNum = otbTmnNo.Text
'                        tVB_CNVirTerminalNum = otbTmnNo.Text '*Eaw 56-06-28
'                        tUT_SaleDate = Format(CDate(odtDate.Value), "yyyy/MM/dd")
'                        tVB_CNReEntryDate = Format(CDate(odtDate.Value), "yyyy/MM/dd") '*Eaw 56-07-02 CommSheet TK-ISS3000-245-TR.xls
'                        tVB_CNAccTypeReEntryMode = otbAccType.Text
'                        tVB_CNTblHD = "TPSHD" & tVB_CNTerminalNum
'                        tVB_CNTblDT = "TPSDT" & tVB_CNTerminalNum
'                        tVB_CNTblRC = "TPSRC" & tVB_CNTerminalNum
'                        tVB_CNTblCD = "TPSCD" & tVB_CNTerminalNum
'                        tVB_CNTblDTTmp = "TPSDT" & tVB_CNTerminalNum & "Tmp"
'                        tVB_CNTblCDTmp = "TPSCD" & tVB_CNTerminalNum & "Tmp"
'                        Call W_CLRxClearDataHD_DT
'                        Call SP_PRCxTerminalRepOnToOff(True)
'                        Call SaveSetting("AdaPosMall", "POSFront", "TmnNum", tVB_CNTerminalNum)
'                        Call SaveSetting("AdaPosMall", "POSFront", "Mode", "2")
'                        Call SaveSetting("AdaPosMall", "POSFront", "SaleDate", tUT_SaleDate)
'                        Call SaveSetting("AdaPosMall", "POSFront", "AccType", tVB_CNAccTypeReEntryMode)
'
'                        Call SP_INSxTransHD(tVB_CNTblHD)
'                        Call wFunctionKB.W_SetToComplete
'                        Call SP_UPDxSetCompleteTerminalMtn
                        '--------------------------------------------------------------------
                        tVB_CNReEntryIDMain = "1721"
                        tVB_CNMnuCurrent = "172101"
                        tVB_Transaction = aVB_TransMode(42).ID
                        tVB_TransactionSub = tVB_Transaction
                        tVB_TransName = aVB_TransMode(42).Name
                        tVB_TransNameABB = aVB_NameAbb(tEN_AbbSaleManualTax - 1).UsrLangPri & ";" & aVB_NameAbb(tEN_AbbSaleManualTax - 1).UsrLangSec
                        tVB_CNTypePrinter = aVB_TransMode(42).PrintType
                        '*Eaw 56-09-30 Copy VirtualPos อีกครั้ง
                        tVB_CNTerminalNum = otbTmnNo.Text
                        Call SP_PRCxTerminalRepOnToOff(True)
                        '----------------------------------------------------
                        tVB_CNVirTerminalNum = otbTmnNo.Text '*Eaw 56-06-28
                        tVB_CNTerminalNum = tVB_CNPhysicalTerminalNum '*BG 56-09-23  Comsheet-356 ให้ virtual ทำงานแบบ Offline ได้ และพิมพ์ EJ
                        tVB_EJDate = tUT_SaleDate            '*Ao 2013/10/01       ReEntry  ให้เก็บ EJ ตาม Operation Date
                        tUT_SaleDate = Format(CDate(odtDate.Value), "yyyy/MM/dd")
                        tVB_CNReEntryDate = Format(CDate(odtDate.Value), "yyyy/MM/dd") '*Eaw 56-07-02 CommSheet TK-ISS3000-245-TR.xls
                        tVB_CNAccTypeReEntryMode = otbAccType.Text
                        tVB_CNTblHD = "TPSHD" & tVB_CNTerminalNum
                        tVB_CNTblDT = "TPSDT" & tVB_CNTerminalNum
                        tVB_CNTblRC = "TPSRC" & tVB_CNTerminalNum
                        tVB_CNTblCD = "TPSCD" & tVB_CNTerminalNum
                        tVB_CNTblDTTmp = "TPSDT" & tVB_CNTerminalNum & "Tmp"
                        tVB_CNTblCDTmp = "TPSCD" & tVB_CNTerminalNum & "Tmp"
                        Call W_CLRxClearDataHD_DT
                        Call UT_CreateSaleTableTMP '*Eaw 56-09-17
'                        Call SP_PRCxTerminalRepOnToOff(True) '*Eaw 56-09-30
                        Call SaveSetting("AdaPosMall", "POSFront", "TmnNum", tVB_CNTerminalNum)
                        Call SaveSetting("AdaPosMall", "POSFront", "Mode", "2")
                        Call SaveSetting("AdaPosMall", "POSFront", "SaleDate", tUT_SaleDate)
                        Call SaveSetting("AdaPosMall", "POSFront", "AccType", tVB_CNAccTypeReEntryMode)
                        Call SaveSetting("AdaPosMall", "POSFront", "VirtualPosNo", tVB_CNVirTerminalNum) '*Thirapong 64-06-09
                        nMaxCode = CLng(Val(Format(SP_SQLtFunction("MAX", "RIGHT(FTLogCode,5)", "TPSTSalHD", " WHERE Left(FTLogCode,9)='" & SP_String2YYMM(Date) & Format(tVB_CNTerminalNum, "00000") & "'", oDB.ocnOnLine, True))))
                        'เปรียบเทียบระหว่าง Master กับ transaction
                        If nMaxCode > nVB_CNLastLogin Then
                             nMaxLogIn = nMaxCode
                        Else
                            nMaxLogIn = nVB_CNLastLogin
                        End If
                        tMaxCode = Format(nMaxLogIn + 1, "00000")
                        nVB_CNLastLogin = nMaxLogIn + 1 '*KT 52-01-21 เก็บค่าไว้ในตัวแปร
                        tUT_LoginCode = SP_String2YYMM(Date) & Format(tVB_CNTerminalNum, "00000") & tMaxCode
                        bUT_SignOut = False
                    Else
                        Call SP_MSGnShowing("tms_CN005243", nCS_Error)
                        Call SP_OBJxSetFocus(odtDate)
                        Exit Sub
                    End If
        End Select
        
        bVB_ReEntryMode = True
        Unload Me
        Call wBackGrand.W_PRVxFormActivate '*Ao 2011/03/28 CommSheet ML-PRV-046- POS Front Terminate Reboot and Not responding.xls
    Else
        Call SP_OTBxSelText(otbTmnNo)
    End If
End Sub


Private Sub odtDate_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 13
            Call SP_OBJxSetFocus(ocmOk) '*Eaw 56-07-25 CommSheet TK-ISS3000-282-TR.xls
            Call ocmOK_Click
    End Select
End Sub

Private Sub otbAccType_Change()
    If tW_AccType = "1" Then
        If otbAccType.Text = "0" Then
            Select Case tVB_CNMnuCurrent
                    Case "1720"
                        wCNAllow.tW_ID = "1720"
                    Case "1721"
                        wCNAllow.tW_ID = "1721"
            End Select
            wCNAllow.Show vbModal
            If tVB_CNUserAlwC = "" Then
                otbAccType.Text = "1"
                tW_AccType = "1"
            End If
        End If
    End If
End Sub

Private Sub otbAccType_GotFocus()
   otbAccType.SelStart = 0
   otbAccType.SelLength = Len(otbAccType.Text)
End Sub

Private Sub otbAccType_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDelete         'delete
            otbAccType.Text = ""
        Case 13
            If Trim(otbAccType.Text) <> "" Then
                SendKeys "{TAB}"
            End If
    End Select
End Sub

Private Sub otbAccType_KeyPress(KeyAscii As Integer)
    Call UT_ChkKeyAscii(KeyAscii)
    Select Case KeyAscii
        Case Asc("0"), Asc("1")
            'none
        Case vbKeyDelete
            'none
        Case vbKeyBack
            'none
        Case Else
            KeyAscii = 0
    End Select
    If otbAccType.Text <> "" Then
        tW_AccType = otbAccType.Text
    End If
End Sub

Private Sub otbTmnNo_GotFocus()
   otbTmnNo.SelStart = 0
   otbTmnNo.SelLength = Len(otbTmnNo.Text)
End Sub

Private Sub otbTmnNo_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDelete         'delete
            otbTmnNo.Text = ""
        Case 13
            If Trim(otbTmnNo.Text) <> "" Then
                SendKeys "{TAB}"
            End If
    End Select
End Sub

Private Sub otbTmnNo_KeyPress(KeyAscii As Integer)
    Call UT_ChkKeyAscii(KeyAscii)
End Sub

Private Function W_PRCbCheckTerminalReEntry(ByVal ptType As String, ByVal ptTmnNo As String) As Boolean
    Dim tSql As String
    Dim orsTemp As ADODB.Recordset
On Error GoTo ErrHandle
    tSql = "SELECT FTTmnNum,FTTmnStaVM,FTTmnIPAddr,FTTmnStatus FROM TCNMTerminalMtn WHERE FTTmnNum ='" & ptTmnNo & "'"
    Select Case ptType
        Case "1720"
            tSql = tSql & vbCrLf & " AND ISNULL(FTTmnStaVM,'N')<>'Y'"  'Y'" '*Eaw 56-10-14 เปลี่ยนจาก'1'-->'Y'
            '*Eaw 56-09-09 CommSheet TK-ISS3000-327 Sale Tax include ใช้ใน Offline Mode ได้
            Call SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine, , , , , True, True)
            If Not (orsTemp.BOF And orsTemp.EOF) Then
                W_PRCbCheckTerminalReEntry = True
            Else
                W_PRCbCheckTerminalReEntry = False
                Call SP_MSGnShowing("tms_CN005250", nCS_Error)
            End If
        Case "1721"
'            tSql = tSql & vbCrLf & " AND ISNULL(FTTmnStatus,'N')<>'A'" '*Eaw 56-09-30 เพิ่ม check status ด้วยถ้า Active อยู่ไม่ให้ใช้ '*Eaw 56-10-26 ไม่ใช้เงื่อนไขนี้ตอน Query
'            tSql = tSql & vbCrLf & " OR ISNULL(FTTmnStaClossing,'Y'<>'N' )" '*Eaw 56-08-17 ถ้า status closing='Y' ให้เอามาใช้ใน ReEntry Mode ได้เลย
            tSql = tSql & vbCrLf & " AND ISNULL(FTTmnStaVM,'N')='Y'" '*Eaw 56-09-09 CommSheet TK-ISS3000-327 Tenant sale ต้องใช้ Virtual POS เท่านั้น  'Y'" '*Eaw 56-10-14 เปลี่ยนจาก'1'-->'Y'
            '*Eaw 56-09-30 Ternant Sale ใช้ใน Offline Mode ได้
            Call SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine, , , , , True, True)
            If Not (orsTemp.BOF And orsTemp.EOF) Then
            '*Eaw 56-10-26 ถ้า Status = A แสดงว่าใช้งานอยู่ให้แสดง Msg ให้ต่างจากกรณีที่หาไม่เจอ CommSheet TK-ISS3000-400-FAPL-145.xls
'                W_PRCbCheckTerminalReEntry = True
                If UCase(SP_FEDtChkString(orsTemp, "FTTmnStatus")) = "A" Then
                    W_PRCbCheckTerminalReEntry = False
                    Call SP_MSGnShowing("tms_CN005254", nCS_Error)
                Else
                    W_PRCbCheckTerminalReEntry = True
                End If
            Else
                W_PRCbCheckTerminalReEntry = False
                Call SP_MSGnShowing("tms_CN005250", nCS_Error)
            End If

        End Select
    Call SP_DATxRsNothing(orsTemp)
    Exit Function
ErrHandle:
    W_PRCbCheckTerminalReEntry = False
     Call SP_DATxRsNothing(orsTemp)
End Function
