VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.1#0"; "Codejock.SuiteCtrls.v11.1.3.ocx"
Begin VB.Form wCNReprint 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reprint Receipt;klg_Title"
   ClientHeight    =   2565
   ClientLeft      =   3630
   ClientTop       =   8535
   ClientWidth     =   4770
   ControlBox      =   0   'False
   Icon            =   "wCNReprint.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2565
   ScaleWidth      =   4770
   Tag             =   "2;"
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   1200
      Index           =   0
      Left            =   0
      TabIndex        =   5
      Top             =   1635
      Width           =   4785
      _Version        =   720897
      _ExtentX        =   8440
      _ExtentY        =   2117
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
      PaintManager.Position=   2
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
         Left            =   3150
         TabIndex        =   2
         TabStop         =   0   'False
         Tag             =   "2;"
         Top             =   240
         Width           =   1455
      End
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
         Left            =   1560
         TabIndex        =   1
         TabStop         =   0   'False
         Tag             =   "2;"
         Top             =   240
         Width           =   1455
      End
   End
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   1935
      Index           =   2
      Left            =   0
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Width           =   4785
      _Version        =   720897
      _ExtentX        =   8440
      _ExtentY        =   3413
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
      PaintManager.Position=   2
      Begin VSFlex7Ctl.VSFlexGrid ogdForm 
         Height          =   1320
         Left            =   120
         TabIndex        =   0
         Top             =   180
         Width           =   4575
         _cx             =   8070
         _cy             =   2328
         _ConvInfo       =   1
         Appearance      =   2
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   11889258
         ForeColorSel    =   16777215
         BackColorBkg    =   16768725
         BackColorAlternate=   16768725
         GridColor       =   16768725
         GridColorFixed  =   16768725
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16768725
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   3
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   500
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"wCNReprint.frx":058A
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   -1  'True
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
      End
   End
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   1935
      Index           =   3
      Left            =   0
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   0
      Width           =   4785
      _Version        =   720897
      _ExtentX        =   8440
      _ExtentY        =   3413
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
      PaintManager.Position=   2
      Begin VB.TextBox otbForm 
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
         Index           =   1
         Left            =   2520
         TabIndex        =   4
         Top             =   840
         Width           =   2055
      End
      Begin VB.TextBox otbForm 
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
         Index           =   0
         Left            =   2520
         TabIndex        =   3
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label olaForm 
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction No :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   360
         TabIndex        =   9
         Top             =   900
         Width           =   2175
      End
      Begin VB.Label olaForm 
         BackStyle       =   0  'Transparent
         Caption         =   "Terminal No :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   8
         Top             =   300
         Width           =   2055
      End
   End
End
Attribute VB_Name = "wCNReprint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private tW_HotKey As String
Private nMaxRow As Long

Private Sub W_PRCxCallFunc(ByVal pnFuc As Long)
    Dim bFound As Boolean
    Dim nI As Long
On Error GoTo ErrHandle
    With ogdForm
        bFound = False
         For nI = .FixedRows To .rows - 1
                If .TextMatrix(nI, 0) <> "" Then
                    If CLng(Val(Format(.TextMatrix(nI, 0)))) = pnFuc Then
                        .Row = nI
                        bFound = True
                        Exit For
                    End If
                End If
         Next nI
         If Not bFound Then Exit Sub
         .Select .Row, 0, .Row, .Cols - 1
         .SetFocus
        .TopRow = .Row
    End With
    Exit Sub
ErrHandle:
    Exit Sub
End Sub

Private Function W_CHKtTransNo(ByVal ptTransNo$) As String
'-----------------------------------------------
'Ret: กรณีที่มีหมายเลข Ticket นี้ในฐานข้อมูลวันที่ขายปัจจุบัน จะเป็น True ถ้าไม่เจอ  False
'Cmt: ค้นว่ามีหมายเลขเอกสารนี้อยู่ในฐานข้อมูลหรือไม่
'------------------------------------------------
    Dim tSql$
    Dim orsTemp As ADODB.Recordset
On Error GoTo ErrHandle
    W_CHKtTransNo = ""
    tSql = "SELECT FTTmnNum,FTShdTransNo,FDShdTransDate,FTShdTransType,FTShdTransTypeSub"
    tSql = tSql & " FROM " & tVB_CNTblHD
    tSql = tSql & " WHERE  FTTmnNum='" & tVB_CNTerminalNum & "' AND FTShdTransNo='" & ptTransNo & "'"
    tSql = tSql & " AND FDShdTransDate='" & tUT_SaleDate & "'"
    Call SP_SQLvExecute(tSql, orsTemp)
    If Not (orsTemp.BOF Or orsTemp.EOF) Then
        W_CHKtTransNo = SP_FEDtChkString(orsTemp, "FTShdTransType")
    End If
    Call SP_DATxRsNothing(orsTemp)
    Exit Function
ErrHandle:
    'non
End Function

Private Sub Form_Activate()
    Call SP_OBJxSetFocus(ogdForm)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        Call ocmOK_Click
    Else
        Select Case KeyAscii
            Case 48 To 58
                With ogdForm
                    tW_HotKey = tW_HotKey & Chr(KeyAscii)
                    If tW_HotKey <> "" Then
                         W_PRCxCallFunc (Val(tW_HotKey))
                    End If
                    If IsNumeric(tW_HotKey) Then
                        If CLng(tW_HotKey) >= nMaxRow Then tW_HotKey = "" ': otmForm.Enabled = False
                    End If
                End With
        End Select
    End If
End Sub

Private Sub Form_Load()
    tW_HotKey = ""
    Call SP_FRMxChildCenter(Me)
    Call SP_FRMxGetCaptionFrmDB2Tag(Me)
    Call SP_FRMxSetCapForm(Me)
    Call W_FNCxAddFunc2Grid
    nMaxRow = ogdForm.Aggregate(flexSTMax, 0, 0, ogdForm.rows - 1, 0)
    bVB_TrnReprint = False
End Sub

'Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'    If UnloadMode = 0 Then Cancel = True
'End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SP_OBJxSetNothing(Me)
End Sub

Private Sub ocmCancel_Click()
    bVB_TrnReprint = False
    Unload Me
End Sub

Private Sub ocmOK_Click()
'*KT 53-05-29 แก้ไขการเกิด Dot send โดยการแก้ไขจากตัวตั้งต้น 3.0.19
Dim tTypeSub  As String

On Error GoTo ErrHandle
    With ogdForm
        Select Case .Row
            Case 0
                Call SP_STAxMouse(Me, True)
                tVB_CNMnuCurrent = ogdForm.TextMatrix(ogdForm.Row, 2)
                If Not SP_USRbMnuAccept(nEN_Usr2AthRead, tVB_CNMnuCurrent) Then Call SP_STAxMouse(Me, False): Exit Sub
                If tVB_OldRunning <> "" Then
                    If Not W_CHKbCV(Trim(tVB_OldRunning), False) Then
                        If bVB_TrainingMode Then
                            tVB_OldTransaction = Format(tEN_TrnTraining, "00")
                            '------------------------------------------------------------------------------------------
                            If SP_CHKbCheckVOID(tVB_OldRunning) Then
                                tVB_OldTypePrinter = "1"
                            Else
                                tTypeSub = SP_CHKtTransNoSub(tVB_OldRunning)
                                tVB_OldTypePrinter = SP_PRCtCheckPRT(tTypeSub)            '*Ao 2013/09/05
                            End If
                             '------------------------------------------------------------------------------------------
                            oTransPrn.TypePriner = tVB_OldTypePrinter
                            oTransPrn.PhysicalTmnNum = tVB_CNPhysicalTerminalNum
                            Call oTransPrn.OpenTransaction(Val(tEN_TrnTraining), CDate(tUT_SaleDate), tVB_OldTerminalNum, tVB_OldRunning, True)
                        Else
                             tVB_OldTransaction = W_CHKtTransNo(tVB_OldRunning)
                             '------------------------------------------------------------------------------------------
                             If SP_CHKbCheckVOID(tVB_OldRunning) Then
                                tVB_OldTypePrinter = "1"
                             Else
                                tVB_OldTypePrinter = SP_PRCtCheckPRT(tVB_OldTransaction)            '*Ao 2013/09/05
                             End If
                             '------------------------------------------------------------------------------------------
                            oTransPrn.TypePriner = tVB_OldTypePrinter
                            oTransPrn.PhysicalTmnNum = tVB_CNPhysicalTerminalNum
                            Call oTransPrn.OpenTransaction(Val(tVB_OldTransaction), CDate(tUT_SaleDate), tVB_OldTerminalNum, tVB_OldRunning, True)
                        End If
                        oTransPrn.RePrint = True
                        oTransPrn.RePrintNormal = True
                        oTransPrn.RePrintCV1 = False
                        oTransPrn.RePrintCV2 = False
                        oTransPrn.PrintEJ = False 'Comsheet ML- SF- 399
                        oTransPrn.PrintOut
                        oTransPrn.RePrint = False
                        oTransPrn.RePrintNormal = False
                        oTransPrn.PrintEJ = True 'Comsheet ML- SF- 399
                        Call SP_STAxMouse(Me, False)
                    Else
                        tVB_OldTransaction = W_CHKtTransNo(tVB_OldRunning)
                        Call SP_STAxMouse(Me, False)
                        wCNReprintSel.tW_TransNo = Trim(tVB_OldRunning)
                        wCNReprintSel.tW_TransType = tVB_OldTransaction
'                        wCNReprintSel.tW_PrintType = tVB_OldTypePrinter
                        '------------------------------------------------------------------------------------------
                        If SP_CHKbCheckVOID(tVB_OldRunning) Then
                            tVB_OldTypePrinter = "1"
                        Else
                            tVB_OldTypePrinter = SP_PRCtCheckPRT(tVB_OldTransaction)            '*Ao 2013/09/05
                        End If
                        '------------------------------------------------------------------------------------------
                        wCNReprintSel.tW_PrintType = tVB_OldTypePrinter
                        wCNReprintSel.Show vbModal
                    End If
                Else
                    tVB_OldRunning = tVB_LRunning
                    tVB_OldTerminalNum = tVB_CNTerminalNum
                    If bVB_TrainingMode Then
                         tVB_OldTransaction = Format(tEN_TrnTraining, "00")
                    Else
                         tVB_OldTransaction = W_CHKtTransNo(tVB_OldRunning)
                    End If
'                    tVB_OldTypePrinter = W_CHKtPrintType(tVB_OldTransaction)
                    '------------------------------------------------------------------------------------------
                    If bVB_TrainingMode Then
                        tTypeSub = SP_CHKtTransNoSub(tVB_OldRunning)
                    Else
                        tTypeSub = tVB_OldTransaction
                    End If
                    If SP_CHKbCheckVOID(tVB_OldRunning) Then
                        tVB_OldTypePrinter = "1"
                    Else
                        tVB_OldTypePrinter = SP_PRCtCheckPRT(tTypeSub)            '*Ao 2013/09/05
                    End If
                    '------------------------------------------------------------------------------------------
                    If tVB_OldTransaction <> "" Then
                        If Not W_CHKbCV(Trim(tVB_OldRunning), False) Then
                            oTransPrn.TypePriner = tVB_OldTypePrinter
                            oTransPrn.PhysicalTmnNum = tVB_CNPhysicalTerminalNum
                            Call oTransPrn.OpenTransaction(Val(tVB_OldTransaction), CDate(tUT_SaleDate), tVB_OldTerminalNum, tVB_OldRunning, True)
                            oTransPrn.RePrint = True
                            oTransPrn.RePrintNormal = True
                            oTransPrn.RePrintCV1 = False
                            oTransPrn.RePrintCV2 = False
                            oTransPrn.PrintEJ = False 'Comsheet ML- SF- 399
                            oTransPrn.PrintOut
                            oTransPrn.RePrint = False
                            oTransPrn.RePrintNormal = False
                            oTransPrn.PrintEJ = True 'Comsheet ML- SF- 399
                            Call SP_STAxMouse(Me, False)
                        Else
                            Call SP_STAxMouse(Me, False)
                            wCNReprintSel.tW_TransNo = Trim(tVB_OldRunning)
                            wCNReprintSel.tW_TransType = tVB_OldTransaction
                            wCNReprintSel.tW_PrintType = tVB_OldTypePrinter
                            wCNReprintSel.Show vbModal
                        End If
                    End If
                End If
            Case 1
                tVB_CNMnuCurrent = ogdForm.TextMatrix(ogdForm.Row, 2)
                If Not SP_USRbMnuAccept(nEN_Usr2AthRead, tVB_CNMnuCurrent) Then Call SP_STAxMouse(Me, False): Exit Sub
                Call SP_STAxMouse(Me, False)
                wCNReprintRef.Show vbModal
                bVB_TrnReprint = True
        End Select
    End With
    Exit Sub
ErrHandle:
    oTransPrn.PrintEJ = True 'Comsheet ML- SF- 399
End Sub


Private Sub ogdForm_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    tW_HotKey = ""
End Sub

Private Sub ogdForm_DblClick()
    Call ocmOK_Click
End Sub

Private Sub W_FNCxAddFunc2Grid()
'--------------------------------------------------------
'Call:
'Cmt: ทำการแสดงรายการของ Last Transaction or Previous Transaction
'-------------------------------------------------------
Dim nRow%
Dim tSmnID$
With ogdForm
    .rows = 2
    If bVB_ReEntryMode Then
        Select Case tVB_CNMnuCurrent
            Case "172004"
                .TextMatrix(0, 1) = Trim(SP_GETtVocabularyTicketFromArray("171201", aVB_TransMnuAll, tSmnID)) '"Last Transaction")
                .TextMatrix(0, 2) = "172004"
                .TextMatrix(1, 1) = Trim(SP_GETtVocabularyTicketFromArray("171202", aVB_TransMnuAll, tSmnID)) '"Previous Transaction"
                .TextMatrix(1, 2) = "172004"
            Case "172104"
                .TextMatrix(0, 1) = Trim(SP_GETtVocabularyTicketFromArray("171201", aVB_TransMnuAll, tSmnID)) '"Last Transaction")
                .TextMatrix(0, 2) = "172104"
                .TextMatrix(1, 1) = Trim(SP_GETtVocabularyTicketFromArray("171202", aVB_TransMnuAll, tSmnID)) '"Previous Transaction"
                .TextMatrix(1, 2) = "172104"
        End Select
    Else
        .TextMatrix(0, 1) = Trim(SP_GETtVocabularyTicketFromArray("171201", aVB_TransMnuAll, tSmnID)) '"Last Transaction")
        .TextMatrix(0, 2) = tSmnID
        .TextMatrix(1, 1) = Trim(SP_GETtVocabularyTicketFromArray("171202", aVB_TransMnuAll, tSmnID)) '"Previous Transaction"
        .TextMatrix(1, 2) = tSmnID
    End If
    For nRow = 0 To 1
            .TextMatrix(nRow, 1) = " " & nRow + 1 & "." & .TextMatrix(nRow, 1)
            .TextMatrix(nRow, 0) = nRow + 1
    Next nRow
End With
End Sub


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
