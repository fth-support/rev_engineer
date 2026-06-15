VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.1#0"; "Codejock.SuiteCtrls.v11.1.3.ocx"
Begin VB.Form wCNDetailCstSex 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customer;klg_Title"
   ClientHeight    =   2760
   ClientLeft      =   2820
   ClientTop       =   3285
   ClientWidth     =   6105
   ControlBox      =   0   'False
   Icon            =   "wCNDetailCstSex.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   6105
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   2040
      Index           =   0
      Left            =   0
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   -240
      Width           =   6120
      _Version        =   720897
      _ExtentX        =   10795
      _ExtentY        =   3598
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
      Begin VSFlex7Ctl.VSFlexGrid ogdSex 
         Height          =   1040
         Left            =   360
         TabIndex        =   0
         Tag             =   "2;ลำดับ;No.;ฟังก์ชัน;Function;คำอธบาย;Description"
         Top             =   840
         Width           =   5415
         _cx             =   9551
         _cy             =   1834
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
         ForeColor       =   -2147483640
         BackColorFixed  =   11889258
         ForeColorFixed  =   -2147483634
         BackColorSel    =   11889258
         ForeColorSel    =   16777215
         BackColorBkg    =   -2147483636
         BackColorAlternate=   16768725
         GridColor       =   4227327
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   8
         GridLinesFixed  =   4
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   3
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   500
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"wCNDetailCstSex.frx":058A
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
         BackColorFrozen =   12632256
         ForeColorFrozen =   0
         WallPaperAlignment=   9
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sex;klg_Sex"
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
         Left            =   360
         TabIndex        =   4
         Tag             =   "2;"
         Top             =   240
         Width           =   2565
      End
   End
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   1200
      Index           =   1
      Left            =   0
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1560
      Width           =   6105
      _Version        =   720897
      _ExtentX        =   10769
      _ExtentY        =   2117
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
      Begin VB.Timer otmLock 
         Interval        =   1500
         Left            =   240
         Top             =   480
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
         Left            =   840
         MaxLength       =   20
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   480
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
         Left            =   4320
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
         Left            =   2760
         TabIndex        =   1
         TabStop         =   0   'False
         Tag             =   "2;"
         Top             =   480
         Width           =   1455
      End
   End
End
Attribute VB_Name = "wCNDetailCstSex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private tW_HotKey As String
Private nMaxRow As Long
Private bW_KeyEnter As Boolean
Private tW_Cap As String  '*KT 52-08-17 Locked Keyboard

Private Sub W_PRCxCallFunc(ByVal pnFuc As Long)         '*AO51-03-12  Comm.Sheet SF-076
    Dim bFound As Boolean
    Dim nI As Long
    On Error GoTo ErrHandle
    With ogdSex
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

Private Sub W_PRCxUpdateCstHD()
'----------------------------------------------------------
'   Call:
'   Cmt : *KT 50-12-26 ทำการ  Update ข้อมูลเพศและสัญชาติลูกค้า
'----------------------------------------------------------
    Dim tSql$
    tSql = "UPDATE " & tVB_CNTblHD
    tSql = tSql & vbCrLf & "SET FTShdSexCode = '" & ogdSex.TextMatrix(ogdSex.Row, 2) & "'"
    tSql = tSql & vbCrLf & "WHERE FTTmnNum='" & tVB_CNTerminalNum & "'"
    tSql = tSql & vbCrLf & "AND FTShdTransNo='" & tVB_Running & "'"
    tSql = tSql & vbCrLf & "AND FTShdTransType = '" & tVB_Transaction & "'"
    tSql = tSql & vbCrLf & "AND FDShdTransDate ='" & tUT_SaleDate & "'"
    tSql = tSql & vbCrLf & "AND FTShdStaDoc = '2'"
    Call SP_SQLvExecute(tSql)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    '*KT 52-08-17 Locked Keyboard
    Select Case Shift
        Case 2
            Select Case KeyCode
                Case vbKeyF1
                        Call SP_OBJxSetFocus(ogdSex)
                         wFunctionKB.W_KBLockKB
                        If bUT_FuncLockKB Then
                            otbID.Text = "Locked"
                            Me.Caption = tW_Cap & "(Locked)"
                            Call SP_FRMxSetTimeCount(False)       'หยุดการจับเวลา Lock
                            otmLock.Enabled = False
                            ogdSex.Enabled = False
                        Else
                            If otbID.Text = "Locked" Then
                                otbID.Text = ""
                                ogdSex.Enabled = True
                                Call SP_OBJxSetFocus(ogdSex)
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
    '*KT 52-08-17 Locked Keyboard
    If bUT_FuncLockKB = False Then
        If KeyAscii = 13 Then
            'Call ocmOK_Click
            If bW_KeyEnter Then
                KeyAscii = 0
            Else
                Call ocmOK_Click
            End If
        Else
            Select Case KeyAscii
                Case 48 To 58
                    
                    With ogdSex
                        tW_HotKey = tW_HotKey & Chr(KeyAscii)
                        If tW_HotKey <> "" Then
                             W_PRCxCallFunc (Val(tW_HotKey))
                        End If
                        If IsNumeric(tW_HotKey) Then
                            If CLng(tW_HotKey) >= nMaxRow Then tW_HotKey = ""
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
     bW_KeyEnter = False
    Call SP_FRMxChildCenter(Me)
    Call SP_FRMxGetCaptionFrmDB2Tag(Me)
    Call SP_FRMxSetCapForm(Me)
    nMaxRow = ogdSex.Aggregate(flexSTMax, 0, 0, ogdSex.rows - 1, 0)
    Call SP_FRMxSetTimeCount(True, nW_TimerLock)       'จับเวลา Lock ใหม่  '*KT 52-08-17 Locked Keyboard
    tW_Cap = Me.Caption
End Sub

'Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'    If UnloadMode = 0 Then Cancel = 1
'End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrHandle
    otmLock.Enabled = False
    Call SP_OBJxSetNothing(Me)
ErrHandle:
    
End Sub

Private Sub ocmCancel_Click()
    otmLock.Enabled = False '*KT 54-04-12 V.3.0.57 ป้องกันโปรแกรมหลุดจากการที่ Timer ทำงานพร้อม Unload
    '*KT 52-08-17 Locked Keyboard
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
        Call W_PRCxUpdateCstHD
        Unload Me
    End If
End Sub

Private Sub ogdSex_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    '*KT 52-08-17 Locked Keyboard
    If bUT_FuncLockKB = False Then
        tW_HotKey = ""
    End If
End Sub

Private Sub ogdSex_DblClick()
    '*KT 52-08-17 Locked Keyboard
    If bUT_FuncLockKB = False Then
        Call ocmOK_Click
    End If
End Sub

Private Sub otmLock_Timer()
     '*KT 52-08-17 Locked Keyboard
On Error GoTo ErrHandle
    If nW_TimerLock = 0 And Screen.ActiveForm.Name = "wCNDetailCstSex" Then
         wFunctionKB.W_KBLockKB
        If bUT_FuncLockKB Then
            otbID.Text = "Locked"
            Me.Caption = tW_Cap & "(Locked)"
            ogdSex.Enabled = False
            Call SP_FRMxSetTimeCount(False)       'หยุดการจับเวลา Lock
            otmLock.Enabled = False
        Else
            If otbID.Text = "Locked" Then
                otbID.Text = ""
                ogdSex.Enabled = True
            End If
             Me.Caption = tW_Cap
            Call SP_OBJxSetFocus(ogdSex)
            Call SP_FRMxSetTimeCount(True, nW_TimerLock)       'จับเวลา Lock
            otmLock.Enabled = True
        End If
    Else
        If bUT_FuncLockKB Then
            otbID.Text = "Locked"
            Me.Caption = tW_Cap & "(Locked)"
            ogdSex.Enabled = False
        Else
            If otbID.Text = "Locked" Then
                otbID.Text = ""
                ogdSex.Enabled = True
                Call SP_OBJxSetFocus(ogdSex)
            End If
            Me.Caption = tW_Cap
        End If
    End If
    Exit Sub
ErrHandle:
End Sub


