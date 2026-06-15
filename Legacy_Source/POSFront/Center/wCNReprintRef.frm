VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.1#0"; "Codejock.SuiteCtrls.Unicode.v11.1.3.ocx"
Begin VB.Form wCNReprintRef 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reprint Receipt;klg_Title"
   ClientHeight    =   2565
   ClientLeft      =   6930
   ClientTop       =   4440
   ClientWidth     =   4755
   ControlBox      =   0   'False
   Icon            =   "wCNReprintRef.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2565
   ScaleWidth      =   4755
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   1935
      Index           =   2
      Left            =   0
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   -255
      Width           =   4740
      _Version        =   720897
      _ExtentX        =   8361
      _ExtentY        =   3413
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
      Begin VB.TextBox otbTransNo 
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
         Left            =   2160
         MaxLength       =   5
         TabIndex        =   0
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction No:;klg_TransNo"
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
         TabIndex        =   4
         Tag             =   "2;"
         Top             =   600
         Width           =   2220
      End
   End
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   1200
      Index           =   1
      Left            =   0
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1350
      Width           =   4740
      _Version        =   720897
      _ExtentX        =   8361
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
         Left            =   3120
         TabIndex        =   2
         TabStop         =   0   'False
         Tag             =   "2;"
         Top             =   495
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
         Left            =   1560
         TabIndex        =   1
         TabStop         =   0   'False
         Tag             =   "2;"
         Top             =   495
         Width           =   1455
      End
   End
End
Attribute VB_Name = "wCNReprintRef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function W_CHKtPrintType(ByVal ptTransType$) As String
    Dim tSql$
    Dim orsTemp As ADODB.Recordset
On Error GoTo ErrHandle
    W_CHKtPrintType = "1"
    tSql = "SELECT  FTSttPrintType FROM TSysTransType WHERE FTSttTranCode = '" & ptTransType & "'"
    If SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine, , , , , True) = 0 Then
        If Not (orsTemp.BOF Or orsTemp.EOF) Then
            W_CHKtPrintType = SP_FEDtChkString(orsTemp, "FTSttPrintType")
            If W_CHKtPrintType = "" Then W_CHKtPrintType = "1"
        End If
    End If
    Call SP_DATxRsNothing(orsTemp)
    Exit Function
ErrHandle:
    'non
    W_CHKtPrintType = "1"
    Call SP_DATxRsNothing(orsTemp)
End Function


Private Function W_CHKtTransNo(ByVal ptTransNo$, Optional ByVal pbUseTmpTbl As Boolean = True) As String
'-----------------------------------------------
'Ret: กรณีที่มีหมายเลข Ticket นี้ในฐานข้อมูลวันที่ขายปัจจุบัน จะเป็น True ถ้าไม่เจอ  False
'Cmt: ค้นว่ามีหมายเลขเอกสารนี้อยู่ในฐานข้อมูลหรือไม่
'------------------------------------------------
    Dim tSql$
    Dim orsTemp As ADODB.Recordset
On Error GoTo ErrHandle
    W_CHKtTransNo = ""
    tSql = "SELECT FTTmnNum, FTShdTransNo,FDShdTransDate,FTShdTransType"
    If pbUseTmpTbl Then
         tSql = tSql & " FROM " & tVB_CNTblHD
    Else
         tSql = tSql & " FROM TPSTSalHD"
    End If
    tSql = tSql & " WHERE  FTTmnNum='" & tVB_CNTerminalNum & "' AND FTShdTransNo='" & ptTransNo & "'"
    tSql = tSql & " AND FDShdTransDate='" & tUT_SaleDate & "'"
    Select Case tVB_CNMnuCurrent
        Case "172004"
            tSql = tSql & " AND FTShdTransType IN ('" & Format(tEN_TrnSaleTaxInc, "00") & "','" & Format(tEN_TrnReturnTaxInc, "00") & "','" & Format(tEN_TrnTotalReportTaxInc, "00") & "')"
        Case "172104"
            tSql = tSql & " AND FTShdTransType IN ('" & Format(tEN_TrnSaleManualTax, "00") & "','" & Format(tEN_TrnReturnManualTax, "00") & "','" & Format(tEN_TrnTotalReportManualTax, "00") & "')"
        Case Else
            tSql = tSql & " AND  FTShdTransType NOT IN ('" & Format(tEN_TrnSaleTaxInc, "00") & "','" & Format(tEN_TrnReturnTaxInc, "00") & "','" & Format(tEN_TrnTotalReportTaxInc, "00") & "'"
            tSql = tSql & ",'" & Format(tEN_TrnSaleManualTax, "00") & "','" & Format(tEN_TrnReturnManualTax, "00") & "','" & Format(tEN_TrnTotalReportManualTax, "00") & "'"
            tSql = tSql & ")"
    End Select
    If SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine, , , , , True) = 0 Then
        If Not (orsTemp.BOF Or orsTemp.EOF) Then
            W_CHKtTransNo = SP_FEDtChkString(orsTemp, "FTShdTransType")
        End If
    End If
    Call SP_DATxRsNothing(orsTemp)
    Exit Function
ErrHandle:
    'non
    Call SP_DATxRsNothing(orsTemp)
End Function

Private Function W_CHKbCV(ByVal ptTransNo$, Optional ByVal pbUseTmpTbl As Boolean = True) As Boolean
'-----------------------------------------------
'Ret:
'Cmt:
'------------------------------------------------
    Dim tSql$
    Dim orsTemp As ADODB.Recordset
On Error GoTo ErrHandle
    W_CHKbCV = False
    tSql = "SELECT FTTmnNum, FTShdTransNo,FDShdTransDate,FTShdTransType"
    If pbUseTmpTbl Then
         tSql = tSql & " FROM " & tVB_CNTblRC
    Else
         tSql = tSql & " FROM TPSTSalRC"
    End If
    tSql = tSql & " WHERE  FTTmnNum='" & tVB_CNTerminalNum & "' AND FTShdTransNo='" & ptTransNo & "'"
    tSql = tSql & " AND FDShdTransDate='" & tUT_SaleDate & "'"
    Select Case tVB_CNMnuCurrent
        Case "172004"
            tSql = tSql & " AND FTShdTransType IN ('" & Format(tEN_TrnSaleTaxInc, "00") & "','" & Format(tEN_TrnReturnTaxInc, "00") & "','" & Format(tEN_TrnTotalReportTaxInc, "00") & "')"
        Case "172104"
            tSql = tSql & " AND FTShdTransType IN ('" & Format(tEN_TrnSaleManualTax, "00") & "','" & Format(tEN_TrnReturnManualTax, "00") & "','" & Format(tEN_TrnTotalReportManualTax, "00") & "')"
        Case Else
            tSql = tSql & " AND  FTShdTransType NOT IN ('" & Format(tEN_TrnSaleTaxInc, "00") & "','" & Format(tEN_TrnReturnTaxInc, "00") & "','" & Format(tEN_TrnTotalReportTaxInc, "00") & "'"
            tSql = tSql & ",'" & Format(tEN_TrnSaleManualTax, "00") & "','" & Format(tEN_TrnReturnManualTax, "00") & "','" & Format(tEN_TrnTotalReportManualTax, "00") & "'"
            tSql = tSql & ")"
    End Select
    tSql = tSql & " AND FTTdmCode = 'T018'"
    If SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine, , , , , True) = 0 Then
        If Not (orsTemp.BOF Or orsTemp.EOF) Then
            W_CHKbCV = True
        End If
    End If
    Call SP_DATxRsNothing(orsTemp)
    Exit Function
ErrHandle:
    'non
    W_CHKbCV = False
End Function


Private Sub Form_Load()
    Call SP_FRMxChildCenter(Me)
    Call SP_FRMxGetCaptionFrmDB2Tag(Me)
    Call SP_FRMxSetCapForm(Me)
End Sub

'Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'    If UnloadMode = 0 Then Cancel = 1
'End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SP_OBJxSetNothing(Me)
End Sub

Private Sub ocmCancel_Click()
    bVB_TrnReprint = False
    Unload Me
End Sub

Private Sub ocmCancel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Unload Me
End Sub

Private Sub ocmOK_Click()
'*KT 53-05-29 แก้ไขการเกิด Dot send โดยการแก้ไขจากตัวตั้งต้น 3.0.19
    Dim tTransType$
    Dim tPrintType$
    Dim tTypeSub As String
    
On Error GoTo ErrHandle

    Call SP_STAxMouse(Me, True)
    If bVB_TrainingMode Then
        tTransType = Format(tEN_TrnTraining, "00")
    Else
        tTransType = W_CHKtTransNo(Trim(otbTransNo.Text), False)
    End If
'    tPrintType = W_CHKtPrintType(tTransType)
    If bVB_TrainingMode Then
        tTypeSub = SP_CHKtTransNoSub(Trim(otbTransNo.Text))             '*Ao 2013/09/05
    Else
        tTypeSub = tTransType
    End If
    If SP_CHKbCheckVOID(Trim(otbTransNo.Text)) Then
        tPrintType = "1"
    Else
        tPrintType = SP_PRCtCheckPRT(tTypeSub)
    End If
    '-----------------------------------------------------------------------------------------------------
    'CV
    If Not W_CHKbCV(Trim(otbTransNo.Text), False) Then
        If tTransType <> "" Then
            oTransPrn.TypePriner = tPrintType
            oTransPrn.PhysicalTmnNum = tVB_CNPhysicalTerminalNum
            Call oTransPrn.OpenTransaction(Val(tTransType), CDate(tUT_SaleDate), tVB_CNTerminalNum, Trim(otbTransNo.Text), False, , True) '*KT 52-03-26 ทำการแก้ไขให้การ Rerint สามารถทำการ Reprint ข้ามกะได้ แต่ข้ามเครื่องไม่ได้
            bVB_RePrint_DeToken = True '*Thirapong 2020-10-14 CommSheet TTK-ISS3000-769
            oTransPrn.RePrint = True
            oTransPrn.RePrintNormal = True
            oTransPrn.RePrintCV1 = False
            oTransPrn.RePrintCV2 = False
            oTransPrn.PrintEJ = False 'Comsheet ML- SF- 399
            oTransPrn.PrintOut
            oTransPrn.RePrint = False
            oTransPrn.RePrintNormal = False
            oTransPrn.PrintEJ = True 'Comsheet ML- SF- 399
            bVB_RePrint_DeToken = False '*Thirapong 2020-10-14 CommSheet TTK-ISS3000-769
        Else
            otbTransNo.SetFocus
        End If
        Call SP_STAxMouse(Me, False)
    Else
        Call SP_STAxMouse(Me, False)
        wCNReprintSel.tW_TransNo = Trim(otbTransNo.Text)
        wCNReprintSel.tW_TransType = tTransType
        wCNReprintSel.tW_PrintType = tPrintType
        wCNReprintSel.Show vbModal
    End If
    Exit Sub
ErrHandle:
    oTransPrn.PrintEJ = True 'Comsheet ML- SF- 399
End Sub

Private Sub otbTransNo_GotFocus()
    Call SP_OBJxSetFocus(otbTransNo)
End Sub

Private Sub otbTransNo_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDelete         'delete
            otbTransNo.Text = ""
    End Select
End Sub

Private Sub otbTransNo_KeyPress(KeyAscii As Integer)
    Call UT_ChkKeyAscii(KeyAscii)
    Select Case KeyAscii
        Case 8, 6           'backspace, delete
            'do nothing
        Case 46             'ทศนิยม
            'do nothing
        Case 13
           Call ocmOK_Click
        Case Else
            If Not IsNumeric(Chr(KeyAscii)) Then
                KeyAscii = 0
            End If
    End Select
End Sub



