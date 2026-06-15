VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.1#0"; "Codejock.SuiteCtrls.v11.1.3.ocx"
Begin VB.Form wCNAge 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customer Age;klg_Title"
   ClientHeight    =   5040
   ClientLeft      =   7500
   ClientTop       =   5415
   ClientWidth     =   4605
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   222
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "wCNAge.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   4605
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   4320
      Index           =   0
      Left            =   0
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   -240
      Width           =   4605
      _Version        =   720897
      _ExtentX        =   8123
      _ExtentY        =   7620
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
      Begin VSFlex7Ctl.VSFlexGrid ogdAge 
         Height          =   2865
         Left            =   480
         TabIndex        =   0
         Top             =   1080
         Width           =   3720
         _cx             =   6562
         _cy             =   5054
         _ConvInfo       =   1
         Appearance      =   2
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   16777215
         ForeColor       =   0
         BackColorFixed  =   16768725
         ForeColorFixed  =   0
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
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   8
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   7
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   400
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"wCNAge.frx":058A
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
         WallPaperAlignment=   4
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Age:;klg_Age"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   24
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   585
         Index           =   0
         Left            =   480
         TabIndex        =   11
         Tag             =   "2;"
         Top             =   360
         Width           =   2820
      End
   End
   Begin VB.OptionButton optAge 
      BackColor       =   &H00FFDED5&
      Caption         =   "<15"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   8
      Top             =   5760
      Width           =   1455
   End
   Begin VB.OptionButton optAge 
      BackColor       =   &H00FFDED5&
      Caption         =   "16-20"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   1
      Left            =   0
      TabIndex        =   7
      Top             =   6240
      Width           =   1455
   End
   Begin VB.OptionButton optAge 
      BackColor       =   &H00FFDED5&
      Caption         =   "21-30"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   2
      Left            =   0
      TabIndex        =   6
      Top             =   6720
      Width           =   1455
   End
   Begin VB.OptionButton optAge 
      BackColor       =   &H00FFDED5&
      Caption         =   "31-45"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   3
      Left            =   0
      TabIndex        =   5
      Top             =   7200
      Width           =   1455
   End
   Begin VB.OptionButton optAge 
      BackColor       =   &H00FFDED5&
      Caption         =   "45-60"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   4
      Left            =   0
      TabIndex        =   4
      Top             =   7680
      Width           =   1455
   End
   Begin VB.OptionButton optAge 
      BackColor       =   &H00FFDED5&
      Caption         =   ">60"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   5
      Left            =   0
      TabIndex        =   3
      Top             =   8160
      Value           =   -1  'True
      Width           =   1455
   End
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   1200
      Index           =   1
      Left            =   0
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   3840
      Width           =   4605
      _Version        =   720897
      _ExtentX        =   8123
      _ExtentY        =   2117
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
      Begin VB.Timer otmLock 
         Interval        =   1500
         Left            =   0
         Top             =   960
      End
      Begin VB.TextBox otbID 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   600
         MaxLength       =   20
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   960
         Visible         =   0   'False
         Width           =   1380
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
         Left            =   2760
         TabIndex        =   2
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
         Left            =   1200
         TabIndex        =   1
         TabStop         =   0   'False
         Tag             =   "2;"
         Top             =   480
         Width           =   1455
      End
   End
End
Attribute VB_Name = "wCNAge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private tW_HotKey$
Private bW_FirstLoad As Boolean
Private bW_KeyEnter As Boolean
Private tW_Cap As String  '*KT 52-08-17 Locked Keyboard

Private Sub W_PRCxCallFunc(ByVal pnFuc As Long)         '*AO51-03-12  Comm.Sheet SF-076
    Dim bFound As Boolean
    Dim nI As Long
    On Error GoTo ErrHandle
    With ogdAge
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

Private Sub W_FRMxShwGrdAge()
'---------------------------------------------------------------
'Call:
'Cmt: *KT 50-12-26 ดึงข้อมูลช่วงอายุมาแสดงใน Grid
'---------------------------------------------------------------
    Dim tSql$
    Dim orsTemp As ADODB.Recordset
    Dim nI As Long
On Error GoTo ErrHandle             '*KT 54-03-25 ป้องกันโปรแกรมหลุด
    tSql = "SELECT  FTCatCode,FTCatDescTh,FTCatDescEn"
    tSql = tSql & " FROM TCNMCstAtt"
    tSql = tSql & " WHERE FTCatStaType='N'"
    tSql = tSql & " ORDER BY FTCatCode"
    Call SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine, , , , , True, True)
    Set ogdAge.DataSource = orsTemp
    Call W_SETxGridCaption(ogdAge)
    With ogdAge
        .ColKey(0) = "FTCatCode"
        .ColKey(1) = "FTCatDescTh"
        .ColKey(2) = "FTCatDescEn"
        .ColHidden(2) = False
         For nI = .FixedRows To .rows - 1
            .TextMatrix(nI, .ColIndex("FTCatDescEn")) = .TextMatrix(nI, .ColIndex("FTCatCode")) & ". " & .TextMatrix(nI, .ColIndex("FTCatDescEn"))
        Next nI
    End With
    Set orsTemp = Nothing
    Exit Sub
ErrHandle:
    Set orsTemp = Nothing
End Sub

Public Sub W_SETxGridCaption(ByVal ptOgdName As VSFlexGrid)
'----------------------------------------------------------
'   Call: all default grid
'   Cmt : เอามารวมกันหมด จะได้วนแค่รอบเดียว
'----------------------------------------------------------
    
    Dim i&
    
    On Error GoTo Err_Handle
    
    With ogdAge
        For i = .FixedCols To .Cols - 1
            If bW_FirstLoad Then
                If .ColKey(i) = "" Then .ColKey(i) = Trim(.TextMatrix(0, i))
                .FixedAlignment(i) = flexAlignCenterCenter
                Select Case UCase(Mid(Trim(.ColKey(i)), 2, 1))
                    Case "T"
                        .ColDataType(i) = flexDTString
                        .ColAlignment(i) = flexAlignLeftCenter
                    Case "D"
                        .ColDataType(i) = flexDTDate
                        .ColAlignment(i) = flexAlignLeftCenter
                        .ColFormat(i) = "DD/MM/YYYY"
                    Case "N", "I"
                        .ColDataType(i) = flexDTLong
                        .ColAlignment(i) = flexAlignRightCenter
                        .ColFormat(i) = "###0"
                    Case "C"
                        .ColDataType(i) = flexDTDouble
                        .ColAlignment(i) = flexAlignRightCenter
                        .ColFormat(i) = "###,###,##0." & String(nVB_CNDecAmtForShw, "0")
                    Case "X", "B"
                        .ColDataType(i) = flexDTBoolean
                End Select
            End If
            .ColHidden(i) = True
            .ScrollBars = flexScrollBarVertical
            Select Case Trim(.ColKey(i))
                Case "FTCatCode"
                    .ColWidth(i) = 1155
                    .ColAlignment(i) = flexAlignCenterCenter
                     .TextMatrix(0, .ColIndex("FTCatCode")) = "Seq"
                Case "FTCatDescTh"
                    .ColWidth(i) = 1155
                    .ColAlignment(i) = flexAlignLeftCenter
                    .TextMatrix(0, .ColIndex("FTCatDescTh")) = "Span of Age"
                Case "FTCatDescEn"
                    .ColWidth(i) = 1155
                    .ColAlignment(i) = flexAlignLeftCenter
                    .TextMatrix(0, .ColIndex("FTCatDescEn")) = "Span of Age"
            End Select
        Next i
    End With
    Exit Sub
Err_Handle:
    Call SP_MSGShowError
End Sub

Private Sub W_PRCxSale()
'----------------------------------------------------------
'   Call:
'   Cmt : *KT 50-12-26
'----------------------------------------------------------
    Dim tSql$
    With ogdAge
        tSql = "UPDATE " & tVB_CNTblHD
        tSql = tSql & " SET FTShdAgeCode = '" & .TextMatrix(.Row, .ColIndex("FTCatCode")) & "'"
        tSql = tSql & "  WHERE FTTmnNum='" & tVB_CNTerminalNum & "'"
        tSql = tSql & "  AND FTShdTransNo='" & tVB_Running & "'"
        tSql = tSql & "  AND FTShdTransType='" & tVB_Transaction & "'"
        tSql = tSql & "  AND FDShdTransDate ='" & tUT_SaleDate & "'"
        tSql = tSql & "  AND FTShdStaDoc = '2'"
        Call SP_SQLvExecute(tSql)
    End With
End Sub

Private Sub W_PRCxVIPSale()
'----------------------------------------------------------
'   Call:
'   Cmt : *KT 50-12-26
'----------------------------------------------------------
    Dim tSql$
    With ogdAge
        tSql = "UPDATE " & tVB_CNTblHD
        tSql = tSql & " SET FTShdAgeCode = '" & .TextMatrix(.Row, .ColIndex("FTCatCode")) & "'"
        tSql = tSql & "  WHERE FTTmnNum='" & tVB_CNTerminalNum & "'"
        tSql = tSql & "  AND FTShdTransNo='" & tVB_Running & "'"
        tSql = tSql & "  AND FTShdTransType='" & tVB_Transaction & "'"
        tSql = tSql & "  AND FDShdTransDate ='" & tUT_SaleDate & "'"
        tSql = tSql & "  AND FTShdStaDoc = '2'"
        'Call DB_ExecuteSQL(tSql)
        Call SP_SQLvExecute(tSql)
    End With
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    '*KT 52-08-17 Locked Keyboard
    Select Case Shift
        Case 2
            Select Case KeyCode
                Case vbKeyF1
                        Call SP_OBJxSetFocus(ogdAge)
                         wFunctionKB.W_KBLockKB
                        If bUT_FuncLockKB Then
                            otbID.Text = "Locked"
                            Me.Caption = tW_Cap & "(Locked)"
                            Call SP_FRMxSetTimeCount(False)       'หยุดการจับเวลา Lock
                            otmLock.Enabled = False
                            ogdAge.Enabled = False
                        Else
                            If otbID.Text = "Locked" Then
                                otbID.Text = ""
                                ogdAge.Enabled = True
                                Call SP_OBJxSetFocus(ogdAge)
                            End If
                            Call SP_FRMxSetTimeCount(True, nW_TimerLock)       'จับเวลา Lock
                            otmLock.Enabled = True
                            Me.Caption = tW_Cap
                        End If
                        KeyCode = 0
            End Select
        Case Else
            If bUT_FuncLockKB Then
                KeyCode = 0
            End If
     End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Dim nMaxRow As Long
    '*KT 52-08-17 Locked Keyboard
    If bUT_FuncLockKB = False Then
        nMaxRow = ogdAge.Aggregate(flexSTMax, 0, 0, ogdAge.rows - 1, 0)
        If KeyAscii = 13 Then
            Call ocmOK_Click
        Else
            Select Case KeyAscii
                Case 48 To 58
                    With ogdAge
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
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
    tW_HotKey = ""
    bW_FirstLoad = True
    bW_KeyEnter = False
    Call SP_FRMxChildCenter(Me)
    Call SP_FRMxGetCaptionFrmDB2Tag(Me)
    Call SP_FRMxSetCapForm(Me)
    Call W_FRMxShwGrdAge
    Call SP_FRMxSetTimeCount(True, nW_TimerLock)       'จับเวลา Lock ใหม่  '*KT 52-08-17 Locked Keyboard
    tW_Cap = Me.Caption
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrHandle
    otmLock.Enabled = False
    Call SP_OBJxSetNothing(Me)
ErrHandle:

End Sub

Private Sub ocmCancel_Click()
    '*KT 52-08-17 Locked Keyboard
    otmLock.Enabled = False '*KT 54-04-12 V.3.0.57 ป้องกันโปรแกรมหลุดจากการที่ Timer ทำงานพร้อม Unload
    If bUT_FuncLockKB = False Then
        Unload Me
    Else
        If otmLock.Enabled = False Then otmLock.Enabled = True '*KT 54-04-12 V.3.0.57 ป้องกันโปรแกรมหลุดจากการที่ Timer ทำงานพร้อม Unload
    End If
End Sub

Private Sub ocmOK_Click()
    '*KT 52-08-17 Locked Keyboard
    If bUT_FuncLockKB = False Then
        
        '*KT 52-07-15 ป้องกันการ Enter หลายครั้งขณะที่ Terminal ทำงานหนักจะเกิด Error ขึ้น
        '------------------------------------------------------------------------------------
        If bW_KeyEnter Then
            Exit Sub
        Else
            bW_KeyEnter = True
        End If
        '------------------------------------------------------------------------------------
        otmLock.Enabled = False '*KT 54-04-12 V.3.0.57 ป้องกันโปรแกรมหลุดจากการที่ Timer ทำงานพร้อม Unload
        Select Case SP_VALnText2Double(tVB_TransactionSub)
            Case tEN_TrnOpen:
            Case tEN_TrnStartBank:
            Case tEN_TrnSale: Call W_PRCxSale
            Case tEN_TrnReturn:
            Case tEN_TrnIPV:
            Case tEN_TrnDeposit:
            Case tEN_TrnRedeem:
            Case tEN_TrnClose:
            Case tEN_TrnPrintClose:
            Case tEN_TrnVIPSale: Call W_PRCxVIPSale
            Case tEN_TrnTaxExceptSale:
            Case tEN_TrnPickUp:
            Case tEN_TrnDutyFreeSale:
            Case tEN_TrnIPVReturn:
            Case tEN_TrnVoucher:
            Case tEN_TrnDepositCancel:
            Case tEN_TrnIPVRedeem:
            Case tEN_TrnNoSale:
            Case tEN_TrnAddDeposit:
            Case tEN_TrnAddDepositCancel:
            Case tEN_TrnCancelVoucher:
            Case tEN_TrnOtherReceive:
            Case tEN_TrnOtherDisbursement:
            Case tEN_TrnDepositMoney:
            Case tEN_TrnTraining:
            Case tEN_TrnIPVVIPService:
            Case tEN_TrnIPVTaxExceptSale:
            Case tEN_TrnIPVDutyFreeSale:
        End Select
        Unload Me
    End If
End Sub

Private Sub ogdAge_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
   '*KT 52-08-17 Locked Keyboard
    If bUT_FuncLockKB = False Then
        tW_HotKey = ""
    End If
End Sub

Private Sub ogdAge_DblClick()
    '*KT 52-08-17 Locked Keyboard
    If bUT_FuncLockKB = False Then
        Call ocmOK_Click
    End If
End Sub


Private Sub otmLock_Timer()
     '*KT 52-08-17 Locked Keyboard
On Error GoTo ErrHandle
    If nW_TimerLock = 0 And Screen.ActiveForm.Name = "wCNAge" Then
         wFunctionKB.W_KBLockKB
        If bUT_FuncLockKB Then
            otbID.Text = "Locked"
            Me.Caption = tW_Cap & "(Locked)"
            ogdAge.Enabled = False
            Call SP_FRMxSetTimeCount(False)       'หยุดการจับเวลา Lock
            otmLock.Enabled = False
        Else
            If otbID.Text = "Locked" Then
                otbID.Text = ""
                ogdAge.Enabled = True
            End If
             Me.Caption = tW_Cap
            Call SP_OBJxSetFocus(ogdAge)
            Call SP_FRMxSetTimeCount(True, nW_TimerLock)       'จับเวลา Lock
            otmLock.Enabled = True
        End If
    Else
        If bUT_FuncLockKB Then
            otbID.Text = "Locked"
            Me.Caption = tW_Cap & "(Locked)"
            ogdAge.Enabled = False
        Else
            If otbID.Text = "Locked" Then
                otbID.Text = ""
                ogdAge.Enabled = True
                Call SP_OBJxSetFocus(ogdAge)
            End If
            Me.Caption = tW_Cap
        End If
    End If
    Exit Sub
ErrHandle:
End Sub


