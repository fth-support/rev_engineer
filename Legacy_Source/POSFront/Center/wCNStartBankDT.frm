VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.1#0"; "Codejock.SuiteCtrls.v11.1.3.ocx"
Begin VB.Form wCNStartBankDT 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "COIN/CURRENCY;klg_Title"
   ClientHeight    =   8490
   ClientLeft      =   3135
   ClientTop       =   2520
   ClientWidth     =   7815
   Icon            =   "wCNStartBankDT.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   7815
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
      Left            =   6120
      TabIndex        =   2
      TabStop         =   0   'False
      Tag             =   "2;"
      Top             =   7800
      Width           =   1455
   End
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   1080
      Index           =   0
      Left            =   0
      TabIndex        =   3
      Top             =   7680
      Width           =   7815
      _Version        =   720897
      _ExtentX        =   13785
      _ExtentY        =   1905
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
      PaintManager.Position=   2
      Begin VB.TextBox otbCurrency 
         Height          =   495
         Left            =   2640
         TabIndex        =   9
         Text            =   "0"
         Top             =   120
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox otbCoin 
         Height          =   495
         Left            =   1680
         TabIndex        =   8
         Text            =   "0"
         Top             =   120
         Visible         =   0   'False
         Width           =   735
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
         Left            =   4560
         TabIndex        =   1
         TabStop         =   0   'False
         Tag             =   "2;"
         Top             =   120
         Width           =   1455
      End
   End
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   7920
      Index           =   1
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   7815
      _Version        =   720897
      _ExtentX        =   13785
      _ExtentY        =   13970
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
      PaintManager.Position=   2
      Begin VSFlex7Ctl.VSFlexGrid ogdFunc 
         Height          =   5970
         Left            =   120
         TabIndex        =   0
         Top             =   915
         Width           =   7575
         _cx             =   13361
         _cy             =   10530
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
         BackColor       =   16769505
         ForeColor       =   0
         BackColorFixed  =   11889258
         ForeColorFixed  =   0
         BackColorSel    =   11889258
         ForeColorSel    =   16777215
         BackColorBkg    =   16768725
         BackColorAlternate=   16768725
         GridColor       =   16777152
         GridColorFixed  =   16768725
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16768725
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   8
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   15
         Cols            =   8
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   500
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"wCNStartBankDT.frx":058A
         ScrollTrack     =   0   'False
         ScrollBars      =   2
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
         Editable        =   2
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
         WallPaperAlignment=   4
      End
      Begin VSFlex7Ctl.VSFlexGrid ogdSum 
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   7035
         Width           =   7575
         _cx             =   13361
         _cy             =   873
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483647
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   415
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"wCNStartBankDT.frx":065C
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
      Begin VSFlex7Ctl.VSFlexGrid VSFlexGrid1 
         Height          =   495
         Left            =   120
         TabIndex        =   7
         Top             =   450
         Width           =   7575
         _cx             =   13361
         _cy             =   873
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483647
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   415
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"wCNStartBankDT.frx":06C0
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
      Begin VB.Label olbForm 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "COIN/CURRENCY;klg_Header"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   14.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Index           =   0
         Left            =   240
         TabIndex        =   5
         Tag             =   "2;"
         Top             =   0
         Width           =   7335
      End
   End
End
Attribute VB_Name = "wCNStartBankDT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private tW_KeyPress As String

Private Sub Form_Load()
On Error GoTo ErrHandle
    Call SP_FRMxChildCenter(Me)
    Call SP_FRMxGetCaptionFrmDB2Tag(Me)
    Call SP_FRMxSetCapForm(Me)
    'Select Case SP_VALnText2Double(tVB_TransactionSub)
    'Call W_FRMxSetInitial
    If Not oOpos Is Nothing Then
        If bVB_OPOSCashDrawerEnable Then Call oOpos.KickDrawer
    Else
         Call cUT_Print.UT_KickDrawer(Me)
    End If
    Call W_FNCxAddFunc2Grid
    
    'ogdFunc.SetFocus
    ogdFunc.Select 0, 2
    ogdFunc.EditCell
   
    
'   Me.ZOrder 0 '*KTs 53-02-03 ป้องกันการ Terminated หรือ  Form wBackGrand ขึ้นมาบังหน้าจอที่ต้องทำงาน
    Exit Sub
ErrHandle:
    Call SP_CRTxLogPrg("ErrorNumber:" & Err.Number & " ErrorDesc:" & Err.Description & " Form:" & Me.Name & " Control:Form" & " Events:Form_Load")
End Sub

Private Sub ocmCancel_Click()
    bVB_CNStartBankDT = False
    Unload Me
End Sub

Private Sub ocmOK_Click()
    bVB_CNStartBankDT = True
    On Error GoTo ErrHandle

    Select Case SP_VALnText2Double(tVB_TransactionSub)
        Case tEN_TrnOpen:
                Unload Me
                Call wBackGrand.W_PRVxFormActivate '*Ao 2011/03/28 CommSheet ML-PRV-046- POS Front Terminate Reboot and Not responding.xls
        Case tEN_TrnStartBank
            If W_PRCbStartBank Then
                Unload Me
                'Call wBackGrand.W_PRVxFormActivate '*Ao 2011/03/28 CommSheet ML-PRV-046- POS Front Terminate Reboot and Not responding.xls
            Else
                Exit Sub
            End If
    End Select
    Exit Sub
ErrHandle:
    
End Sub

Private Sub W_FRMxSetInitial()
    ogdSum.TextMatrix(0, 2) = "0.00"
End Sub


'NEUNG XXX
Private Sub W_FNCxAddFunc2Grid(Optional ByVal ptChkBnkStaAlw As String = "Y")
'--------------------------------------------------------
'Call:
'Cmt: แสดง Banknote
'-------------------------------------------------------
Dim nI As Integer, nR As Integer
Dim orsTemp As ADODB.Recordset
Dim tSql As String
On Error GoTo ErrHandle
    With ogdFunc

        tSql = "SELECT * FROM TCNMBanknote"
        If ptChkBnkStaAlw = "Y" Then
            tSql = tSql & vbCrLf & " WHERE FTBnkStaAlw='Y'"
        End If
        tSql = tSql & vbCrLf & " ORDER BY FCBnkAmt,FTBnkStaTpye,FTBnkCode"
        
        Call SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine, , , , , True, True) 'read TCNMBanknote
        If Not (orsTemp.BOF Or orsTemp.EOF) Then
            nR = 0
            .rows = orsTemp.RecordCount
            Do While Not orsTemp.EOF
                .TextMatrix(nR, 0) = nR + 1
                .TextMatrix(nR, 1) = SP_FMTtNum2Shw(SP_FEDvChkNumeric(orsTemp, "FCBnkAmt")) & " " & SP_FEDtChkString(orsTemp, "FTBnkNameEn")
                .TextMatrix(nR, 2) = 0
                .TextMatrix(nR, 3) = SP_FMTcCurPoint(0, False, True)
                .TextMatrix(nR, 4) = SP_FEDvChkNumeric(orsTemp, "FCBnkAmt")
                .TextMatrix(nR, 5) = SP_FEDtChkString(orsTemp, "FTBnkCode")
                .TextMatrix(nR, 6) = SP_FEDtChkString(orsTemp, "FTBnkNameTh")
                .TextMatrix(nR, 7) = SP_FEDtChkString(orsTemp, "FTBnkStaTpye")
                .Cell(flexcpBackColor, nR, 2, nR, 2) = &H80000005
                
                nR = nR + 1
                orsTemp.MoveNext
            Loop
            '.ColAlignment(2) = flexAlignRightCenter
           ' .Cell(flexcpAlignment, .FixedRows, 1, .rows - 1, 1) = flexAlignRightCenter
        End If
    End With
          
    Exit Sub
ErrHandle:
    Set orsTemp = Nothing
    Call SP_MSGShowError
End Sub

Private Sub ogdFunc_AfterEdit(ByVal Row As Long, ByVal Col As Long)
   With ogdFunc
        Select Case Col
            Case 2
             
                If .TextMatrix(Row, Col) = "" Then .TextMatrix(Row, Col) = 0
                
                If .TextMatrix(Row, 4) = "T017" Then
                    Select Case SP_VALnText2Double(tVB_TransactionSub)
                        Case tEN_TrnPickUp, tEN_TrnOverShortReport
                               .TextMatrix(Row, 3) = SP_VALcText2Double(.TextMatrix(Row, 2))
                    End Select
                End If
                If SP_VALcText2Double(.TextMatrix(Row, 2)) = 0 Then
                    .TextMatrix(Row, 3) = 0#
                Else
                    .TextMatrix(Row, 3) = SP_VALcText2Double(.TextMatrix(Row, 2)) * SP_VALcText2Double(.TextMatrix(Row, 4))
                End If
                Call W_DEFxSumStartBank
        End Select
    End With
End Sub

Private Sub ogdFunc_AfterMoveRow(ByVal Row As Long, Position As Long)
tW_KeyPress = ""
End Sub

Private Sub ogdFunc_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
   With ogdFunc
        Select Case Col
            Case 2
                '.FocusRect = flexFocusHeavy
                '.HighLight = flexHighlightAlways
                .EditMaxLength = 5  '*Eaw 56-04-03
            Case Else
                Cancel = True
        End Select
    End With
End Sub

Private Sub ogdFunc_ChangeEdit()
    With ogdFunc
        Select Case .Col
            Case 1
                'MsgBox tW_KeyPress
                'MsgBox SP_DEFtChangeNum(tW_KeyPress)
                '.EditText = SP_DEFtChangeNum(tW_KeyPress)
        End Select
    End With
End Sub


Private Sub ogdFunc_GotFocus()
    With ogdFunc
        .HighLight = flexHighlightAlways
        .SelectionMode = flexSelectionByRow
        .Editable = flexEDKbdMouse
         tW_KeyPress = "" 'Replace(.TextMatrix(.Row, .Col), ".", "")
    End With
End Sub

Private Sub ogdFunc_KeyDown(KeyCode As Integer, Shift As Integer)
    With ogdFunc
        Select Case KeyCode
            Case 13
                tW_KeyPress = ""
                If .Row = .rows - 1 Then
                    If .EditText <> "" Then
                        ogdFunc.Select 0, 2
                    End If
                Else
                    If .EditText <> "" Then
                        ogdFunc.Select .Row + 1, 2
                    End If
                End If
            Case vbKeyDown
                If .Row = .rows - 1 Then
                    ogdFunc.Select .FixedRows, 2
                End If
            Case vbKeyUp
                If .Row = 0 Then
                    ogdFunc.Select .rows - 1, 2
                End If
        End Select
    End With
End Sub

Private Sub ogdFunc_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
 With ogdFunc
        Select Case KeyAscii
            Case 13
                tW_KeyPress = ""
                If .Row = .rows - 1 Then
                    If .EditText <> "" Then
                        ogdFunc.Select 0, 2
                    End If
                Else
                    If .EditText <> "" Then
                        ogdFunc.Select Row + 1, 2
                    End If
                End If
            Case Else
                Select Case KeyAscii
                    Case vbKeyDelete
                        tW_KeyPress = ""
                    Case vbKeyBack
                        '*Eaw 56-04-22  CommSheet TK-ISS3000-154-IT2-Eaw.xlsx
'                        If Len(Trim(tW_KeyPress)) > 0 Then
'                            tW_KeyPress = Left(tW_KeyPress, Len(tW_KeyPress) - 1)
'                        End If

                        If Len(Trim(tW_KeyPress)) > 0 Then
                            If Len(Trim(tW_KeyPress)) > 5 Then
                                tW_KeyPress = Left(tW_KeyPress, 4)
                            Else
                                tW_KeyPress = Left(tW_KeyPress, Len(tW_KeyPress) - 1)
                            End If
                        End If
                        With ogdFunc
                            Select Case .Col
                                Case 2
                                    .Text = tW_KeyPress
                            End Select
                        End With
                    '------------------------------------------------------------------------------------------
                    Case Else
                        Call cUT.UT_KBDKeyCurrency(KeyAscii, tW_KeyPress)    'check key, currency
                        If KeyAscii > 0 Then tW_KeyPress = IIf(KeyAscii = 13, "", tW_KeyPress & Chr(KeyAscii)) 'ถ้า  KeyAscii = Enter หมายถึงจบการแก้ไขราคาต่อหน่วย
                End Select
        End Select
    End With
End Sub

Private Sub ogdFunc_LostFocus()
    With ogdFunc
        .SelectionMode = flexSelectionFree
        .Editable = flexEDNone
    End With
End Sub

  
Private Sub W_DEFxSumStartBank()
Dim nLoop As Integer
Dim cSumCash As Double
Dim cSumCoin As Double
Dim cSumCur As Double

    cSumCash = 0
    cSumCoin = 0
    cSumCur = 0
    
    With ogdFunc
        Select Case SP_VALnText2Double(tVB_TransactionSub)
            Case tEN_TrnPickUp, tEN_TrnOverShortReport
                For nLoop = .FixedRows To .rows - 2
                    If .TextMatrix(nLoop, 2) <> "" Then
                         cSumCash = cSumCash + SP_FMTcCurPoint(.ValueMatrix(nLoop, 3), True, False)
                        If Trim(.TextMatrix(nLoop, 7)) = "Y" Then
                                cSumCur = cSumCur + SP_FMTcCurPoint(.ValueMatrix(nLoop, 3), True, False)
                        Else
                                cSumCoin = cSumCoin + SP_FMTcCurPoint(.ValueMatrix(nLoop, 3), True, False)
                        End If
                    End If
                Next nLoop
            Case Else
                For nLoop = .FixedRows To .rows - 1
                    If .TextMatrix(nLoop, 2) <> "" Then
                         cSumCash = cSumCash + SP_FMTcCurPoint(.ValueMatrix(nLoop, 3), True, False)
                        If Trim(.TextMatrix(nLoop, 7)) = "Y" Then
                                cSumCur = cSumCur + SP_FMTcCurPoint(.ValueMatrix(nLoop, 3), True, False)
                        Else
                                cSumCoin = cSumCoin + SP_FMTcCurPoint(.ValueMatrix(nLoop, 3), True, False)
                        End If
                    End If
                Next nLoop
        End Select
    End With
    otbCoin.Text = SP_FMTcCurPoint(cSumCoin, True, True)
    otbCurrency.Text = SP_FMTcCurPoint(cSumCur, True, True)
    ogdSum.TextMatrix(0, 2) = SP_FMTcCurPoint(cSumCash, True, True)
   
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

Private Function W_INSxTranDT26112012() As Boolean
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

        W_INSxTranDT26112012 = False
        
        With ogdFunc
            nR = 1
             For nI = .FixedRows To nRow
                    If .TextMatrix(nI, 0) <> "" Then
                    
                            tUT_SKUCode = .TextMatrix(nI, 5)
                            tSKUName = Trim(.TextMatrix(nI, 1))     'ชื่อสินค้า(Default ENG)1
                            tSKUName2 = Trim(.TextMatrix(nI, 6))    'ชื่อสินค้า(Local)2
                            tSKUNameSrvDoc = Trim(.TextMatrix(nI, 1))    'ชื่อสินค้า Service Doc(Default ENG) 1
                            tSKUNameSrvDoc2 = Trim(.TextMatrix(nI, 6))     'ชื่อสินค้า Service Doc (Local)2
                           
                            cQty = SP_FMTcCurPoint(.ValueMatrix(nI, 2), True, False)
                            cQtyAll = cQty
                            cRegPrice = SP_FMTcCurPoint(.ValueMatrix(nI, 4), True, False)
                            cSalePrice = SP_FMTcCurPoint(.ValueMatrix(nI, 4), True, False)
                            cSaleUnit = SP_FMTcCurPoint(.ValueMatrix(nI, 4), True, False)
                            cB4DisChg = SP_FMTcCurPoint(.ValueMatrix(nI, 4), True, False)
                            cSdtStkFac = 1
                            cSaleAmt = SP_FMTcCurPoint(.ValueMatrix(nI, 3), True, False)
                            If Trim(.TextMatrix(nI, 7)) = "Y" Then
                                tSdtStaPdt = "1"
                                otbCurrency.Text = SP_VALcText2Double(otbCurrency.Text) + cSaleAmt
                            Else
                                tSdtStaPdt = "0"
                                otbCoin.Text = SP_VALcText2Double(otbCoin.Text) + cSaleAmt
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
                                   W_INSxTranDT26112012 = True
                               End If
                           End If
                    End If
                    nR = nR + 1
             Next nI
        End With
        Exit Function
ErrHandle:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Function

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
                       tSdtStaPdt = "0"
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
        Exit Function
ErrHandle:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Function

                
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
On Error GoTo ErrHandle

    cVB_Change = 0
    cFAmt = CDbl(Val(Format(ogdSum.TextMatrix(0, 2))))
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
    tVB_TenderCode = aVB_TenderMnu(0).ID
    tVB_TenderType = aVB_TenderMnu(0).Name '*KT 53-10-15
    tVB_TenderName = aVB_TenderMnu(0).DescSale
    
    tSql = "DELETE " & tVB_CNTblRC
    tSql = tSql & vbCrLf & "WHERE FTTmnNum = '" & tVB_CNTerminalNum & "'"
    tSql = tSql & vbCrLf & "AND FTShdTransNo = '" & tVB_Running & "'"
    tSql = tSql & vbCrLf & "AND FDShdTransDate = '" & tUT_SaleDate & "'"
    tSql = tSql & vbCrLf & "AND FTShdTransType = '" & tVB_Transaction & "'"
    Call SP_SQLvExecute(tSql)
    
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
    tSql = tSql & vbCrLf & "'" & tVB_TenderName & "',"
    If tVB_CrdExpM = "" And tVB_CrdExpY = "" Then
        tSql = tSql & "'" & "" & "',"
    Else
        tSql = tSql & "'" & tVB_CrdExpM & "/" & tVB_CrdExpY & "',"
    End If
    tSql = tSql & 0 & ","
    tSql = tSql & 0 & ","
    tSql = tSql & UT_tLastInsSql & ")"
    Call SP_SQLvExecute(tSql)
    
    Exit Sub
ErrHandle:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Sub


Private Sub W_UPDxTransHD()
'------------------------------------------------------
'Call:
'Cmt:   *KT 50-12-24 ปรับปรุงข้อมูล Transaction HD
'-----------------------------------------------------
    Dim tSql$
    tSql = "UPDATE " & tVB_CNTblHD
    tSql = tSql & vbCrLf & " SET FCShdTotal =" & SP_FMTcCurPoint(SP_VALcText2Double(ogdSum.TextMatrix(0, 2)), True, False)
    tSql = tSql & vbCrLf & ",FCShdGrand=" & SP_FMTcCurPoint(SP_VALcText2Double(ogdSum.TextMatrix(0, 2)), True, False)
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
    
    tSql = tSql & vbCrLf & " WHERE FTTmnNum='" & tVB_CNTerminalNum & "'"
    tSql = tSql & vbCrLf & " AND FTShdTransNo='" & tVB_Running & "'"
    tSql = tSql & vbCrLf & " AND FTShdTransType='" & tVB_Transaction & "'"
    tSql = tSql & vbCrLf & " AND FDShdTransDate ='" & tUT_SaleDate & "'"
    tSql = tSql & vbCrLf & " AND ISNULL(FTShdStaDoc,'0')  <> '1'"
    'Call DB_ExecuteSQL(tSql)
    Call SP_SQLvExecute(tSql)
    cVB_CoinsStartBank = SP_VALcText2Double(otbCoin.Text)
    cVB_CurrencyStartBank = SP_VALcText2Double(otbCurrency.Text)
    cVB_TotalStartBank = SP_VALcText2Double(ogdSum.TextMatrix(0, 2))
End Sub

