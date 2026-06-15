VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.1#0"; "Codejock.SuiteCtrls.v11.1.3.ocx"
Begin VB.Form wTrainMode 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Training Mode;klg_Title"
   ClientHeight    =   2445
   ClientLeft      =   7410
   ClientTop       =   3150
   ClientWidth     =   4905
   ControlBox      =   0   'False
   Icon            =   "wTrainMode.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   4905
   Tag             =   "2;"
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   1200
      Index           =   0
      Left            =   0
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1515
      Width           =   4905
      _Version        =   720897
      _ExtentX        =   8652
      _ExtentY        =   2117
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
      PaintManager.Position=   2
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
         Left            =   1680
         TabIndex        =   1
         TabStop         =   0   'False
         Tag             =   "2;"
         Top             =   240
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
         Left            =   3240
         TabIndex        =   2
         TabStop         =   0   'False
         Tag             =   "2;"
         Top             =   240
         Width           =   1455
      End
   End
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   1815
      Index           =   2
      Left            =   0
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   0
      Width           =   4905
      _Version        =   720897
      _ExtentX        =   8652
      _ExtentY        =   3201
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
      PaintManager.Position=   2
      Begin VSFlex7Ctl.VSFlexGrid ogdForm 
         Height          =   1080
         Left            =   165
         TabIndex        =   0
         Top             =   180
         Width           =   4575
         _cx             =   8070
         _cy             =   1905
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
         FormatString    =   $"wTrainMode.frx":058A
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
End
Attribute VB_Name = "wTrainMode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private tW_HotKey As String
Private nMaxRow As Long

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
    bVB_TrnTrainningMode = False
    Call SP_FRMxChildCenter(Me)
    Call SP_FRMxGetCaptionFrmDB2Tag(Me)
    Call SP_FRMxSetCapForm(Me)
    Call W_FNCxAddFunc2Grid
    nMaxRow = ogdForm.Aggregate(flexSTMax, 0, 0, ogdForm.rows - 1, 0)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SP_OBJxSetNothing(Me)
End Sub

Private Sub ocmCancel_Click()
    bVB_TrnTrainningMode = False
    Unload Me
End Sub

Private Sub ocmOK_Click()
    Dim tSql$
    Dim orsTemp As ADODB.Recordset
    
    bVB_TrnTrainningMode = True
    
    tVB_CNMnuCurrent = ogdForm.TextMatrix(ogdForm.Row, 2)
    If Not SP_USRbMnuAccept(nEN_Usr2AthRead, tVB_CNMnuCurrent) Then Exit Sub
    With ogdForm
        Select Case .ValueMatrix(.Row, 0)
            Case 1
                 tVB_Transaction = aVB_TransMode(24).ID
                 tVB_TransactionSub = tVB_Transaction
                tVB_TransName = aVB_TransMode(24).Name
                tVB_TransNameABB = aVB_NameAbb(tEN_AbbTrainingOn - 1).UsrLangPri & ";" & aVB_NameAbb(tEN_AbbTrainingOn - 1).UsrLangSec
                tVB_CNTypePrinter = aVB_TransMode(24).PrintType
                If bVB_TrainingMode Then
                    Call SP_MSGnShowing("tms_CN005058", tCS_CNMsgErr)
                    Exit Sub
                End If
                bVB_TrainingMode = True
                Call SP_INSxTransHD(tVB_CNTblHD)
                wFunctionKB.W_SetToComplete
'*KT 51-10-24 V 3.0.141/V 3.0.200
'------------------------------------------------------------------------------
                Call SP_UPDxSetCompleteTerminalMtn
'-----------------------------------------------------------------------------
            Case 2
                tVB_Transaction = aVB_TransMode(24).ID
                tVB_TransactionSub = tVB_Transaction
                tVB_TransName = aVB_TransMode(24).Name
                tVB_TransNameABB = aVB_NameAbb(tEN_AbbTrainingOff - 1).UsrLangPri & ";" & aVB_NameAbb(tEN_AbbTrainingOff - 1).UsrLangSec
                tVB_CNTypePrinter = aVB_TransMode(24).PrintType
                If Not bVB_TrainingMode Then
                    Call SP_MSGnShowing("tms_CN005059", tCS_CNMsgErr)
                    Exit Sub
                End If
                bVB_TrainingMode = False
                Call SP_INSxTransHD(tVB_CNTblHD)
                 wFunctionKB.W_SetToComplete
 '*KT 51-10-24 V 3.0.141/V 3.0.200
'------------------------------------------------------------------------------
                Call SP_UPDxSetCompleteTerminalMtn
'------------------------------------------------------------------------------
        End Select
    End With
        
    tSql = "SELECT FTCfgCode,FTTmnName "
    tSql = tSql & " FROM TConfigTerminal"
    tSql = tSql & " WHERE FTCfgCode='LModeTraining'"
    Set orsTemp = ocnCfg.Execute(tSql)
    If Not (orsTemp.BOF Or orsTemp.EOF) Then
        If orsTemp.RecordCount > 1 Then
            tSql = "DELETE FROM "
            tSql = tSql & "  TConfigTerminal"
            tSql = tSql & " WHERE FTCfgCode='LModeTraining'"
            Call ocnCfg.Execute(tSql) 'Call ocnAll.Execute(tSql) '*KT 52-12-15
            tSql = "INSERT INTO TConfigTerminal"
            tSql = tSql & "(FTTmnCode,FTTmnName,FTCfgCode,FTCfgOn,FTCfgOff,FTCfgValue,FTCfgDescTha,FTCfgDescEng)"
            tSql = tSql & " VALUES ("
            tSql = tSql & "'" & tVB_CNTerminalNum & "'"
            tSql = tSql & ",'" & tVB_CNComName & "'"
            tSql = tSql & ",'" & "LModeTraining" & "'"
            tSql = tSql & ",'" & "" & "'"
            tSql = tSql & ",'" & "" & "'"
            tSql = tSql & ",'" & IIf(bVB_TrainingMode, "1", "0") & "'"
            tSql = tSql & ",'" & "สถานนะการทดลองใช้งานบนเครื่องจุดขาย 0 : ปกติ 1 : Training" & "'"
            tSql = tSql & ",'" & "Status mode trainning of Terminal Machine 0 :Nomal 1 : Training" & "'"
            tSql = tSql & ")"
        Else
            tSql = "UPDATE"
            tSql = tSql & "  TConfigTerminal"
            tSql = tSql & " SET FTCfgValue = '" & IIf(bVB_TrainingMode, "1", "0") & "'"
            tSql = tSql & " ,FTTmnName = '" & tVB_CNComName & "'"
            tSql = tSql & " ,FTTmnCode = '" & tVB_CNTerminalNum & "'"
            tSql = tSql & " WHERE FTCfgCode='LModeTraining'"
        End If
    Else
        tSql = "INSERT INTO TConfigTerminal"
        tSql = tSql & "(FTTmnCode,FTTmnName,FTCfgCode,FTCfgOn,FTCfgOff,FTCfgValue,FTCfgDescTha,FTCfgDescEng)"
        tSql = tSql & " VALUES ("
        tSql = tSql & "'" & tVB_CNTerminalNum & "'"
        tSql = tSql & ",'" & tVB_CNComName & "'"
        tSql = tSql & ",'" & "LModeTraining" & "'"
        tSql = tSql & ",'" & "" & "'"
        tSql = tSql & ",'" & "" & "'"
        tSql = tSql & ",'" & IIf(bVB_TrainingMode, "1", "0") & "'"
        tSql = tSql & ",'" & "สถานนะการทดลองใช้งานบนเครื่องจุดขาย 0 : ปกติ 1 : Training" & "'"
        tSql = tSql & ",'" & "Status mode trainning of Terminal Machine 0 :Nomal 1 : Training" & "'"
        tSql = tSql & ")"
    End If
    Call ocnCfg.Execute(tSql)
    Call SP_DATxRsNothing(orsTemp)
    Unload Me
End Sub

Private Sub ogdForm_DblClick()
    Call ocmOK_Click
End Sub

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

Private Sub W_FNCxAddFunc2Grid()
'--------------------------------------------------------
'Call:
'Cmt: ทำการแสดงรายการของ Training Mode On or Training Mode Off
'-------------------------------------------------------
Dim nRow%
Dim tSmnID$
With ogdForm
    .rows = 2
    .TextMatrix(0, 1) = Trim(SP_GETtVocabularyTicketFromArray("171401", aVB_TransMnuAll, tSmnID)) '"Training Mode On"
    .TextMatrix(0, 2) = tSmnID
    .TextMatrix(1, 1) = Trim(SP_GETtVocabularyTicketFromArray("171402", aVB_TransMnuAll, tSmnID)) '"Training Mode Off"
    .TextMatrix(1, 2) = tSmnID
    For nRow = 0 To 1
            .TextMatrix(nRow, 1) = " " & nRow + 1 & "." & .TextMatrix(nRow, 1)
            .TextMatrix(nRow, 0) = nRow + 1
    Next nRow
End With
End Sub

Private Sub ogdForm_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    tW_HotKey = ""
End Sub



