VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.1#0"; "Codejock.SuiteCtrls.v11.1.3.ocx"
Begin VB.Form wCNCurrency 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Foreign Currency;klg_Title"
   ClientHeight    =   3195
   ClientLeft      =   5100
   ClientTop       =   7260
   ClientWidth     =   5100
   ControlBox      =   0   'False
   Icon            =   "wCNCurrency.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   5100
   Tag             =   "2;"
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   2535
      Index           =   2
      Left            =   0
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   -240
      Width           =   5085
      _Version        =   720897
      _ExtentX        =   8969
      _ExtentY        =   4471
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
      Begin VSFlex7Ctl.VSFlexGrid ogdForm 
         Height          =   2000
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   4815
         _cx             =   8493
         _cy             =   3528
         _ConvInfo       =   1
         Appearance      =   2
         BorderStyle     =   0
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
         BackColorBkg    =   16763579
         BackColorAlternate=   16768725
         GridColor       =   16763579
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16763579
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   5
         Cols            =   2
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   500
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"wCNCurrency.frx":058A
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
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   1200
      Index           =   1
      Left            =   0
      TabIndex        =   4
      Top             =   1995
      Width           =   5085
      _Version        =   720897
      _ExtentX        =   8969
      _ExtentY        =   2117
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
      Begin VB.Timer otmForm 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   1080
         Top             =   480
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
         Left            =   3360
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
         Left            =   1800
         TabIndex        =   1
         TabStop         =   0   'False
         Tag             =   "2;"
         Top             =   480
         Width           =   1455
      End
   End
End
Attribute VB_Name = "wCNCurrency"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public tW_HotKey As String
Public nMaxRow As Long
Private Function W_GETcForeignPay() As Double
'----------------------------------------------------------
'Call:
'Cmt: ใช้คำนวนหาค่า Needed Foreign Currency '*Eaw 56-08-09 CommSheet TK-ISS3000-271-TR.xls
'----------------------------------------------------------
    Dim cLocalAmt As Double
    Dim tTotal As String
    Dim tMin As String
    Dim aNum() As String
    Dim aMinDeno() As String
    Dim nIndex As Integer
    Dim nDec As Long
    Dim nPoint As Long
    Dim nMinDec As Long
    Dim nMinPoint As Long
    Dim cNewAmt As Double
    Dim cMod As Double
On Error GoTo ErrHandle
    
    cLocalAmt = SP_FMTcCurPoint(cVB_Balance / cVB_FCSrcRteFac, False, True)
    tMin = SP_FMTtNum2Shw(cVB_MinDeno)
    tTotal = SP_FMTtNum2Shw(cLocalAmt)
    aNum = Split(tTotal, ".")
    aMinDeno = Split(tMin, ".")
    For nIndex = 0 To UBound(aNum())
        Select Case nIndex
            Case 0
                nDec = aNum(nIndex)
            Case 1
                nPoint = aNum(nIndex)
                Exit For
        End Select
    Next nIndex
    For nIndex = 0 To UBound(aMinDeno())
        Select Case nIndex
            Case 0
                nMinDec = SP_VALcText2Double(aMinDeno(nIndex))
            Case 1
                nMinPoint = SP_VALcText2Double(aMinDeno(nIndex))
                Exit For
        End Select
    Next nIndex
'    If cVB_MinDeno < 1 Then
'        nDec = nDec + 1
'        nPoint = 0
'        cNewAmt = SP_FMTcCurPoint(nDec & "." & nPoint, False, True)
'    Else
'        cMod = cLocalAmt \ cVB_MinDeno
'        If (cMod * cVB_MinDeno) >= cLocalAmt Then
'            cMod = cMod - 1
'        End If
'        cNewAmt = SP_FMTcCurPoint(cMod * cVB_MinDeno + cVB_MinDeno, False, True)
'    End If
'*Eaw 56-10-11 แก้ไขเนื่องจากมี effect จาก Min Deno เป็นเลขทศนิยม
    If nMinPoint <> 0 Then
        cNewAmt = SP_FMTcCurPoint(cLocalAmt, False, True)
    Else
        cMod = cLocalAmt \ cVB_MinDeno
        If (cMod * cVB_MinDeno) >= cLocalAmt Then
            cMod = cMod - 1
        End If
        cNewAmt = SP_FMTcCurPoint(cMod * cVB_MinDeno + cVB_MinDeno, False, True)
    End If
    cVB_NeedAmt = cNewAmt
    W_GETcForeignPay = cNewAmt
    Exit Function
ErrHandle:
    W_GETcForeignPay = cLocalAmt
End Function

Private Sub Form_Activate()
On Error GoTo ErrHandle
     Call SP_OBJxSetFocus(ogdForm)   'ogdForm.SetFocus
    Exit Sub
ErrHandle:
        
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
       If KeyCode = 13 Then ocmOK_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call ocmOK_Click
    Else
        Select Case KeyAscii
            Case 48 To 58
                otmForm.Enabled = True
                With ogdForm
                    tW_HotKey = tW_HotKey & Chr(KeyAscii)
                    If tW_HotKey <> "" Then
                         W_PRCxCallFunc (Val(tW_HotKey))
                    End If
                    If IsNumeric(tW_HotKey) Then
                        If CLng(tW_HotKey) >= nMaxRow Then tW_HotKey = "": otmForm.Enabled = False
                    End If
                End With
        End Select
    End If
End Sub

Private Sub Form_Load()
    tW_HotKey = ""
    Call SP_FRMxChildCenter(Me)
    If bVB_Res640X480 Then
'        Me.Left = 0
        Me.Top = 2100
    End If
    Call SP_FRMxGetCaptionFrmDB2Tag(Me)
    Call SP_FRMxSetCapForm(Me)
    Call W_FNCxAddCur2Grid
    nMaxRow = ogdForm.Aggregate(flexSTMax, 0, 0, ogdForm.rows - 1, 0)
End Sub

Private Sub W_FNCxAddCur2Grid()
'---------------------------------
'Call:
'Cmt: ทำการดึงข้อมุลสกุลเงินต่างประเทศมาแสดงที่ตาราง
'---------------------------------
    Dim tSql$
    Dim nI As Long
    Dim orsTemp As ADODB.Recordset
'    tSql = "SELECT 1 AS FNSeqNo,FTFcrType AS FTName,FTFcrType,FTFcrCode,FTFcrSign,FCFcrExchang,FTFcrStaLocal"
    tSql = "SELECT 1 AS FNSeqNo,FTFcrType AS FTName,FTFcrType,FTFcrCode,FTFcrSign,FCFcrExchang,FTFcrStaLocal,ISNULL(FCFcrMinDeno,0) AS FCFcrMinDeno,ISNULL(FCFcrMaxChange,0) AS FCFcrMaxChange" '*Eaw 56-08-08 เพิ่ม FCFcrMinDeno,FCFcrMaxChange '*Eaw 56-08-08 CommSheet TK-ISS3000-271-TR.xls
    tSql = tSql & vbCrLf & "FROM TCNMForCurrency"
    tSql = tSql & vbCrLf & "WHERE FTTdmCode ='" & tVB_TenderCode & "'"
    tSql = tSql & vbCrLf & "AND FTFcrStaLocal <>'Y' "
    tSql = tSql & vbCrLf & "ORDER BY FTFcrCode"
    Call SP_SQLbExecute(tSql, orsTemp, , , , oDB.ocnOnLine, True, True) 'read master
    Set ogdForm.DataSource = orsTemp
    ogdForm.Refresh
    For nI = 0 To orsTemp.Fields.Count - 1
        ogdForm.ColKey(nI) = orsTemp.Fields(nI).Name
        ogdForm.ColHidden(nI) = True
    Next nI
     ogdForm.ColHidden(ogdForm.ColIndex("FTName")) = False
    With ogdForm
        For nI = 0 To .rows - 1
            .TextMatrix(nI, .ColIndex("FNSeqNo")) = nI + 1
            .TextMatrix(nI, .ColIndex("FTName")) = nI + 1 & ". " & Trim(.TextMatrix(nI, .ColIndex("FTFcrType")))            '*Ao 2013/03/14
        Next nI
    End With
'    Set orsTemp = Nothing
    Call SP_DATxRsNothing(orsTemp)
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrHandle
    otmForm.Enabled = False
    Call SP_OBJxSetNothing(Me)
ErrHandle:
    
End Sub

'Private Sub Form_Unload(Cancel As Integer)
'    Set wCNCurrency = Nothing
'End Sub

Private Sub ocmCancel_Click()
    cVB_Pay = 0 '*KT 51-08-11 Communication Sheet ML-SF-435 Foreign Currency.xls ข้อที่2
    tVB_ForeignCurName = ""
    tVB_FTRteCode = ""
    cVB_FCSrcRteFac = 0
    tVB_ForeignCurSign = ""
    '*Eaw 56-08-08 CommSheet TK-ISS3000-271-TR.xls
    cVB_MinDeno = 0
    cVB_MaxChange = 0
    Unload Me
End Sub

Private Sub ocmOK_Click()
    With ogdForm
        If ogdForm.rows = 0 Then Unload Me: Exit Sub
        tVB_ForeignCurName = .TextMatrix(.Row, .ColIndex("FTFcrType"))
        tVB_FTRteCode = .TextMatrix(.Row, .ColIndex("FTFcrCode"))
        cVB_FCSrcRteFac = .TextMatrix(.Row, .ColIndex("FCFcrExchang"))
        tVB_ForeignCurSign = .TextMatrix(.Row, .ColIndex("FTFcrSign"))
        '*Eaw 56-08-08 CommSheet TK-ISS3000-271-TR.xls
        cVB_MinDeno = .TextMatrix(.Row, .ColIndex("FCFcrMinDeno"))
        cVB_MaxChange = .TextMatrix(.Row, .ColIndex("FCFcrMaxChange"))
        If cVB_FCSrcRteFac > 0 Then


'            wCNForeignPay.olaCurLocal.Caption = tVB_ForeignCurSign & " " & SP_FMTcCurPoint(cVB_Balance / cVB_FCSrcRteFac, False, True)
            If cVB_MinDeno <> 0 Then '*Eaw 56-08-08 CommSheet TK-ISS3000-271-TR.xls
                wCNForeignPay.olaCurLocal.Caption = tVB_ForeignCurSign & " " & SP_FMTcCurPoint(W_GETcForeignPay, False, True)
            Else
                wCNForeignPay.olaCurLocal.Caption = tVB_ForeignCurSign & " " & SP_FMTcCurPoint(cVB_Balance / cVB_FCSrcRteFac, False, True)
                cVB_NeedAmt = SP_FMTcCurPoint(cVB_Balance / cVB_FCSrcRteFac, False, True)
            End If
'            wCNForeignPay.olaCurLocal.Caption = tVB_ForeignCurSign & " " & SP_FMTcCurPoint(cVB_Balance / SP_FMTcCurPoint(cVB_FCSrcRteFac, True, True), False, True)   '*Eaw 56-01-29   แก้ไขการ Rounding สำหรับ TK Project
            wCNForeignPay.Show vbModal
        Else
            Call SP_MSGnShowing("tms_CN005187", nCS_Warning)
            cVB_Pay = 0 '*KT 51-08-11 Communication Sheet ML-SF-435 Foreign Currency.xls ข้อที่2
            tVB_ForeignCurName = ""
            tVB_FTRteCode = ""
            cVB_FCSrcRteFac = 0
            tVB_ForeignCurSign = ""
            '*Eaw 56-08-08
            cVB_MinDeno = 0
            cVB_MaxChange = 0
        End If
    End With
    Unload Me
End Sub

Private Sub ogdForm_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    tW_HotKey = ""
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

Private Sub otmForm_Timer()
    Dim nMaxRow As Long
    nMaxRow = ogdForm.Aggregate(flexSTMax, 0, 0, ogdForm.rows - 1, 0)
    If IsNumeric(tW_HotKey) Then tW_HotKey = "": otmForm.Enabled = False
End Sub

