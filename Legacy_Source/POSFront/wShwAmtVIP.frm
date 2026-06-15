VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.1#0"; "Codejock.SuiteCtrls.Unicode.v11.1.3.ocx"
Begin VB.Form wShwAmtVIP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Amount of VIP Service Sale"
   ClientHeight    =   2475
   ClientLeft      =   2880
   ClientTop       =   2655
   ClientWidth     =   4785
   ControlBox      =   0   'False
   Icon            =   "wShwAmtVIP.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2475
   ScaleWidth      =   4785
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   1815
      Index           =   2
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   -240
      Width           =   4785
      _Version        =   720897
      _ExtentX        =   8440
      _ExtentY        =   3201
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
      Begin VB.Label olaAmount 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1800
         TabIndex        =   4
         Top             =   720
         Width           =   2775
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   840
         Width           =   1530
      End
   End
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   1200
      Index           =   1
      Left            =   0
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1260
      Width           =   4785
      _Version        =   720897
      _ExtentX        =   8440
      _ExtentY        =   2117
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
      Begin VB.CommandButton ocmOk 
         Cancel          =   -1  'True
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
         TabIndex        =   0
         Tag             =   "2;"
         Top             =   480
         Width           =   1455
      End
   End
End
Attribute VB_Name = "wShwAmtVIP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Call SP_FRMxChildCenter(Me)
    Call SP_FRMxGetCaptionFrmDB2Tag(Me)
    Call SP_FRMxSetCapForm(Me)
    Call W_SHWxSumAmtVIPSale
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SP_OBJxSetNothing(Me)
End Sub

Private Sub ocmOK_Click()
'    Select Case SP_VALnText2Double(tVB_TransactionSub)
'        Case tEN_TrnOpen:
'        Case tEN_TrnStartBank:
'        Case tEN_TrnSale: 'Call W_PRCxSale
'        Case tEN_TrnReturn:
'        Case tEN_TrnIPV:
'        Case tEN_TrnDeposit:
'        Case tEN_TrnRedeem:
'        Case tEN_TrnClose:
'        Case tEN_TrnPrintClose:
'        Case tEN_TrnVIPSale: 'Call W_PRCxVIPSale
'        Case tEN_TrnTaxExceptSale:
'        Case tEN_TrnPickUp:
'        Case tEN_TrnDutyFreeSale:
'        Case tEN_TrnIPVReturn:
'        Case tEN_TrnVoucher:
'        Case tEN_TrnDepositCancel:
'        Case tEN_TrnIPVRedeem:
'        Case tEN_TrnNoSale:
'        Case tEN_TrnAddDeposit:
'        Case tEN_TrnAddDepositCancel:
'        Case tEN_TrnCancelVoucher:
'        Case tEN_TrnOtherReceive:
'        Case tEN_TrnOtherDisbursement:
'        Case tEN_TrnDepositMoney:
'        Case tEN_TrnTraining:
'        Case tEN_TrnIPVVIPService:
'        Case tEN_TrnIPVTaxExceptSale:
'        Case tEN_TrnIPVDutyFreeSale:
'        Case tEN_TrnCreateVIP: 'Call W_PRCxCreateVIPSale
'    End Select
    Unload Me
End Sub

Private Sub W_SHWxSumAmtVIPSale()
'---------------------------------------------------
'   Call:
'   Cmt: ทำการหายอดรวมของลูกค้า VIP ที่ขายแบบ VIP Service Sale โดยใช้ Function การพักบิลมาใช้ในการขาย VIP Service Sale ในแต่ละครั้ง
'---------------------------------------------------
    Dim tWhe$
    Dim tSql$
    Dim orsTemp As ADODB.Recordset
    Dim tCardNo As String '*Bmp 61-02-15 Token
    Dim tTokenCardNo As String
'    tWhe = " WHERE  FTShdTransType='" & Format(tEN_TrnCreateVIP, "00") & "' AND FDShdTransDate='" & tUT_SaleDate & "'"
'    tWhe = tWhe & " AND  FTCstVIPCode='" & tVB_CstID & "'"
'    olaAmount.Caption = SP_FMTtNum2Shw(SP_SQLtFunction("SUM", "FCShdGrand + FCShdRnd", "TPSTHoldHD", tWhe, oDB.ocnOnLine), False, True) 'PH 2.0.0 connect database server

    tSql = "SELECT SUM(FCShdGrand + FCShdRnd) AS FCGrand,MAX(FTMemberPriceID) AS FTMemberPriceID,MAX(FTPrvCardID) AS FTPrvCardID,MAX(FTPrvCardType) AS FTPrvCardType"
    tSql = tSql & vbCrLf & "FROM TPSTHoldHD"
    tSql = tSql & vbCrLf & "WHERE  FTShdTransType='" & Format(tEN_TrnCreateVIP, "00") & "' AND FDShdTransDate='" & tUT_SaleDate & "'"
    tSql = tSql & vbCrLf & "AND  FTCstVIPCode='" & tVB_CstID & "'"
    tSql = tSql & vbCrLf & "AND  ISNULL(FTShdStaPaid,'')<>'3'"   '*Eaw 56-03-14
    If SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine) = 0 Then
        If Not (orsTemp.BOF And orsTemp.EOF) Then
            olaAmount.Caption = SP_FMTtNum2Shw(SP_FEDvChkNumeric(orsTemp, "FCGrand"), False, True) 'PH 2.0.0 connect database server
            xVB_CNMemberP.ID = SP_FEDtChkString(orsTemp, "FTMemberPriceID")
        Else
            olaAmount.Caption = "0.00" 'PH 2.0.0 connect database server
            xVB_CNMemberP.ID = ""
        End If
        '*KT 53-11-04 Preview Day
        '-------------------------INITIAL-----------------------------------------------------------------------
        Call SP_INIxSwitchModePreviewDay(False)
        Call SP_INIxClearVarPrv
        tVB_CNPreDayIDReg = SP_FEDtChkString(orsTemp, "FTPrvCardID")
        If tVB_CNPreDayIDReg <> "" Then
            'อยู่ในช่วง Preview Day
            If SP_CHKbPeriodPreDay Then
                nVB_RegPreType = SP_FEDtChkString(orsTemp, "FTPrvCardType")
                bVB_CNTrStaRegPreviewDay = True
                Call SP_INIxSwitchModePreviewDay(True)
                wMain.olaTransName.Caption = "Preview/" & tVB_TransName
                Select Case nVB_RegPreType
                    Case 1 'Credit Card
                        Call SP_GETxRecDataCrdPrv
                    Case 2 'Member Card
                        Call SP_GETxRecDataMallPrv
                    Case 3 'None Card
                        Call SP_GETxRecDataMallPrv
                End Select
                If Not SP_UPDbPreviewDayHD Then 'บันทึกข้อมูล Preview Day ลงตาราง HD
                    Call SP_MSGnShowing("tms_CN005221", nCS_Warning)
                    Call SP_INIxSwitchModePreviewDay(False)
                    Call SP_INIxClearVarPrv
                    wMain.olaTransName.Caption = tVB_TransName
                End If
            End If
        End If
        '-------------------------END INITIAL---------------------------------------------------------------------
    Else
        olaAmount.Caption = "0.00" 'PH 2.0.0 connect database server
        xVB_CNMemberP.ID = ""
    End If
    Call SP_DATxRsClose(orsTemp)
    'PH2-ML-SF-126 VIP Sale เมื่อทำการกด function Member Price แล้ว เมื่อทำการขายที่เครื่องอื่นๆ ไม่ต้องทำการกด Function Member Price อีก
    '*KT 52-08-14 member price
    '-------------------------------------------------------------------------------------------------------
    If xVB_CNMemberP.ID <> "" Then
        '*Bmp 61-02-15 Token
        tCardNo = xVB_CNMemberP.ID
        tTokenCardNo = ""
        If SP_PRCbToken("SP_GETnMemLPoint", tVB_MemPointID, tTokenCardNo) Then
            tCardNo = tTokenCardNo
        End If
        '===================================
    
        bVB_CNMemberPrice = True
        tSql = "SELECT Top 1 FTMcdName,FTMcdSurName,FTMcdCode,FTMcdTel,FDMcdExpDate,FTMcdSta  "
        tSql = tSql & vbCrLf & " FROM TCNMMallCard"
        'tSql = tSql & vbCrLf & " WHERE FTMcdCode='" & xVB_CNMemberP.ID & "'"
        tSql = tSql & vbCrLf & " WHERE FTMcdCode='" & tCardNo & "'" '*Bmp 61-02-15 Token
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
    Call SP_DATxRsNothing(orsTemp)
    Exit Sub
ErrHandle:
    olaAmount.Caption = "0.00"
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
    Call SP_DATxRsNothing(orsTemp)
End Sub


