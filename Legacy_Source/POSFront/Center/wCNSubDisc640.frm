VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.1#0"; "Codejock.SuiteCtrls.Unicode.v11.1.3.ocx"
Begin VB.Form wCNSubDisc640 
   BorderStyle     =   0  'None
   Caption         =   "wCNSubDisc640"
   ClientHeight    =   6000
   ClientLeft      =   2430
   ClientTop       =   3720
   ClientWidth     =   9600
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   9600
   ShowInTaskbar   =   0   'False
   Begin VSFlex7Ctl.VSFlexGrid ogdForm 
      Height          =   1430
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1275
      Width           =   9600
      _cx             =   16933
      _cy             =   2522
      _ConvInfo       =   1
      Appearance      =   2
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   16777215
      ForeColor       =   0
      BackColorFixed  =   11889258
      ForeColorFixed  =   16777215
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   11889258
      BackColorAlternate=   16768725
      GridColor       =   33023
      GridColorFixed  =   16761024
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   16777215
      FocusRect       =   0
      HighLight       =   0
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   8
      GridLinesFixed  =   4
      GridLineWidth   =   1
      Rows            =   4
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   350
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"wCNSubDisc640.frx":0000
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
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   855
      Index           =   2
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   5455
      Width           =   9600
      _Version        =   720897
      _ExtentX        =   16933
      _ExtentY        =   1508
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
      PaintManager.Position=   2
      Begin VB.Timer otmRec 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   360
         Top             =   0
      End
      Begin VB.CommandButton ocmHome 
         Caption         =   "TOTAL"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4560
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   90
         Width           =   1455
      End
      Begin VB.CommandButton ocmOk 
         Caption         =   "&Ok"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6240
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   90
         Width           =   1455
      End
      Begin VB.CommandButton ocmCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7920
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   90
         Width           =   1455
      End
      Begin VB.Timer otmForm 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   3360
         Top             =   0
      End
      Begin VB.Timer otmLock 
         Interval        =   2000
         Left            =   3840
         Top             =   0
      End
      Begin VB.Label olaMEM 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   240
         TabIndex        =   17
         Top             =   120
         Visible         =   0   'False
         Width           =   4225
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid ogdFunc 
      Height          =   2340
      Left            =   0
      TabIndex        =   5
      Top             =   3100
      Width           =   9600
      _cx             =   16933
      _cy             =   4128
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
      SelectionMode   =   1
      GridLines       =   8
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   9
      Cols            =   5
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   350
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"wCNSubDisc640.frx":00E7
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
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   675
      Index           =   0
      Left            =   0
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2700
      Width           =   9590
      _Version        =   720897
      _ExtentX        =   16916
      _ExtentY        =   1191
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
      PaintManager.Position=   2
      Begin VB.Label olaForm 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Discount Type"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   405
         Index           =   0
         Left            =   -240
         TabIndex        =   7
         Top             =   0
         Width           =   9255
      End
   End
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   1545
      Index           =   1
      Left            =   0
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   0
      Width           =   9600
      _Version        =   720897
      _ExtentX        =   16933
      _ExtentY        =   2725
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
      PaintManager.Position=   2
      Begin VB.Label olaBal 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H0000C000&
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   720
         Left            =   7920
         TabIndex        =   16
         Top             =   390
         Width           =   1410
      End
      Begin VB.Label olaBalCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Balance"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   435
         Left            =   4560
         TabIndex        =   15
         Top             =   0
         Width           =   1395
      End
      Begin VB.Label olaTSubTotal 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   3705
         TabIndex        =   14
         Top             =   750
         Width           =   795
      End
      Begin VB.Label olbSubTotal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sub Total."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   345
         Left            =   120
         TabIndex        =   13
         Top             =   750
         Width           =   1455
      End
      Begin VB.Label olaTSubPro 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   3705
         TabIndex        =   12
         Top             =   360
         Width           =   795
      End
      Begin VB.Label olbSubPro 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sub Pro."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   345
         Left            =   120
         TabIndex        =   11
         Top             =   405
         Width           =   1200
      End
      Begin VB.Label olbSubReg 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sub Reg."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   345
         Left            =   120
         TabIndex        =   10
         Top             =   45
         Width           =   1275
      End
      Begin VB.Label olaTSubReg 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   3705
         TabIndex        =   9
         Top             =   0
         Width           =   795
      End
   End
End
Attribute VB_Name = "wCNSubDisc640"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*Ao 52-02-25
' ถ้ามีการแก้ Code ใน From wCNSubDisc2 ให้แก้ใน From wCNSubDisc640 ด้วย
'wCNSubDisc2 ใช้ Resorution 800*600
'wCNSubDisc640 ใช้ Resorution 640*480
'________________________________________________________
Private tW_HotKey As String
Private nW_TransMnu As Long
Private tW_TransFunc As String
Private tW_SmnID As String
Private nMaxRow As Long

Private Sub W_GETxMemberBP()
Dim tSql As String
Dim orsTemp As ADODB.Recordset

On Error GoTo ErrHandle
    tSql = "SELECT FTCstCardCode,FTCstCardType,FTShdCshOrCrd FROM " & tVB_CNTblHD
    tSql = tSql & vbCrLf & "WHERE FTTmnNum = '" & tVB_CNTerminalNum & "'"
    tSql = tSql & vbCrLf & "AND FTShdTransNo = '" & tVB_Running & "'"
    tSql = tSql & vbCrLf & "AND FTShdTransType = '" & tVB_Transaction & "'"
    tSql = tSql & vbCrLf & "AND FDShdTransDate = '" & tUT_SaleDate & "'"
    Call SP_SQLvExecute(tSql, orsTemp)
    If Not (orsTemp.BOF And orsTemp.EOF) Then
        tVB_MemPointID = SP_FEDtChkString(orsTemp, "FTCstCardCode")
        tVB_MallCardIDRef = SP_FEDtChkString(orsTemp, "FTShdCshOrCrd") '*Tao 60-04-21 ถ้าได้ MallCardID จาก Manual ให้ใส่ *
        tVB_MemExp = SP_FEDtChkString(orsTemp, "FTCstCardType")
    End If
Exit Sub
ErrHandle:

End Sub

Private Sub W_CHKxDiscount91(ByVal pnKeyCode As Integer, ByVal pnShift As Integer)
'______________________________________
'   Call:
'   Cmt: '*Eaw 56-01-05 ใช้เรียก  Hot Key ของ Discount 91
'______________________________________
Dim tSql$
Dim tProc$
Dim orsFunc As ADODB.Recordset
Dim nR As Long
Dim nI As Long
 On Error GoTo ErrHandle
    bVB_Discount91 = False
    tSql = "SELECT FTSkbRef FROM TSysKeyboard   WHERE FNSkbKey='" & CStr(pnKeyCode) & "'"
    tSql = tSql & " AND FNSkbShift = '" & CStr(pnShift) & "'"
    tSql = tSql & " AND FTSkbwMain =  '1'"
    tSql = tSql & " AND FTSkbUsage ='Y'"
'    Debug.Print pnKeyCode
'    Debug.Print pnShift
    Call SP_SQLvExecute(tSql, orsFunc, oDB.ocnOnLine, , , , , True, True) 'read master
    If Not (orsFunc.EOF And orsFunc.BOF) Then
         tProc = SP_FEDtChkString(orsFunc, "FTSkbRef")
         If tProc = "W_KBDiscount91" Then
             bVB_Discount91 = True
'            With ogdFunc
'                 .Select 1, 0, 1, .Cols - 1
'                 .SetFocus
'                 .TopRow = .Row
'            End With
'            Call ocmOK_Click
            With ogdFunc
                nR = LBound(aVB_DiscMnu)
                For nI = LBound(aVB_DiscMnu) To UBound(aVB_DiscMnu) - 1
                    If aVB_DiscMnu(nI).SmnID = "1809" Then
                       .Select nR, 0, nR, .Cols - 1
                       .SetFocus
                       .TopRow = .Row
                       Call ocmOk_Click
                       Exit Sub
                    End If
                    nR = nR + 1
                Next nI
            End With
         End If
         '*BG 57-05-07  PRJ10001TK-PH2-External Spec-Hot Key 00.03.04 FSBT.docx
         If tProc = "W_KBSetSlipPrinter" Then
            Call SP_MSGxManualShow(tMS_CN522, vbExclamation + vbOKOnly, tVB_CNAppTitle)
        ElseIf tProc = "W_KBMainLayaway" Or tProc = "W_KBMainReturn" Then
            Call SP_MSGxManualShow(tMS_CN523, vbOKOnly, tVB_CNAppTitle)
         End If
         '----------------------------------------------------------------------------------------------------------
    End If
    Exit Sub
ErrHandle:
     bVB_Discount91 = False
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Sub

Private Sub W_CHKxReVendor()
'______________________________________
'   Call:
'   Cmt: '*Ao 2011/01/15
'______________________________________
Dim tSql As String
Dim orsTemp As ADODB.Recordset

    tSql = "UPDATE " & tVB_CNTblRC
    tSql = tSql & vbCrLf & "SET FTSrcStaNoPay='0'"
    tSql = tSql & vbCrLf & "WHERE"
    tSql = tSql & vbCrLf & "FTTmnNum = '" & tVB_CNTerminalNum & "'"
    tSql = tSql & vbCrLf & "AND FTShdTransNo = '" & tVB_Running & "'"
    tSql = tSql & vbCrLf & "AND FTShdTransType = '" & tVB_Transaction & "'"
    tSql = tSql & vbCrLf & "AND FDShdTransDate = '" & tUT_SaleDate & "'"
    Call SP_SQLvExecute(tSql)
    Exit Sub
Err_Handler:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Sub

Private Sub W_DEFxCheckVendor()
'__________________________________
'Call:
'Cmt: '*Ao 53-10-19
'__________________________________
Dim tWhr As String

    On Error GoTo ErrHandle
    
    tWhr = "WHERE FTTmnNum = '" & tVB_CNTerminalNum & "'"
    tWhr = tWhr & vbCrLf & "AND FTShdTransNo = '" & tVB_Running & "'"
    tWhr = tWhr & vbCrLf & "AND FTShdTransType = '" & tVB_Transaction & "'"
    tWhr = tWhr & vbCrLf & "AND FDShdTransDate = '" & tUT_SaleDate & "'"
    tWhr = tWhr & vbCrLf & "AND FTTdmCode = 'T006'"
    tWhr = tWhr & vbCrLf & "AND FTVendorCnpFlag = '1'"
    tWhr = tWhr & vbCrLf & "AND ISNULL(FTSrcStaNoPay,'') <> '1'"   '*Ao 2011/01/17   CommSheet ML-PRV-020-Change Tender is wrong concept.xls
    cVB_TotalVendor = SP_VALcText2Double(SP_SQLtFunction("SUM", "FCSrcNet", tVB_CNTblRC, tWhr))
    
Exit Sub
ErrHandle:
    cVB_TotalVendor = 0
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Sub

'Private Sub W_CHKxCheckVendor()
''_________________________________
''Call:
''Cmt'           '*Ao 53-10-13
''_________________________________
'Dim tSql As String, tWhr As String
'Dim orsTemp As ADODB.Recordset
'
'    On Error GoTo ErrHandle
'
'    tWhr = " WHERE  FTTmnNum=''" & tVB_CNTerminalNum & "'' AND FTShdTransNo=''" & tVB_Running & "''"
'    tWhr = tWhr & vbCrLf & " AND  FTShdTransType=''" & tVB_Transaction & "'' AND FDShdTransDate=''" & tUT_SaleDate & "''"
'    tWhr = tWhr & vbCrLf & " AND  FTTdmCode = ''T006''" & " AND FTVendorCnpFlag=''1''"
'    tWhr = tWhr & vbCrLf & " ORDER BY  FNSrcSeqNo ASC"
'    tSql = "STP_CNxSQLFunction @ptFunc = 'FNSrcSeqNo,FCSrcAmt,FCSrcNet,FCSrcChg' ,@ptTable = '" & tVB_CNTblRC & "',@ptWhere = '" & tWhr & "'"
'    Call SP_SQLvExecute(tSql, orsTemp)
'    If Not (orsTemp.BOF And orsTemp.EOF) Then
'        tSql = "UPDATE " & tVB_CNTblRC
'        tSql = tSql & vbCrLf & "SET FCSrcAmt=FCSrcFAmt"
'        tSql = tSql & vbCrLf & ",FCSrcNet=FCSrcFAmt"
'        tSql = tSql & vbCrLf & ",FCSrcChg=0"
''        tSql = tSql & vbCrLf & ",FTSrcRmk=''"
'
'        tSql = tSql & vbCrLf & "WHERE FTVendorCnpFlag = '1'"
'        tSql = tSql & vbCrLf & "AND FTTmnNum = '" & tVB_CNTerminalNum & "'"
'        tSql = tSql & vbCrLf & "AND FTShdTransNo = '" & tVB_Running & "'"
'        tSql = tSql & vbCrLf & "AND FTShdTransType = '" & tVB_Transaction & "'"
'        tSql = tSql & vbCrLf & "AND FDShdTransDate = '" & tUT_SaleDate & "'"
'        Call SP_SQLvExecute(tSql)
'
'    End If
'
'Exit Sub
'ErrHandle:
'    Call SP_MSGxShowError(Err.Number, Err.Description)
'End Sub
'
'Private Sub W_DEFxCheckVendor()
''__________________________________
''Call:
''Cmt: '*Ao 53-10-19
''__________________________________
'Dim tSql As String, tSkuIndex As String, tSkuIndexTmp As String, tWhr As String, tSort As String
'Dim orsRC As ADODB.Recordset
'Dim orsDT As ADODB.Recordset
'Dim orsTemp As ADODB.Recordset
'Dim cSrcNet As Double, cSaleAmt As Double, cChange As Double, cSrcNetTmp As Double
'Dim nPos As Integer, nSeqNo As Integer
'Dim bChkUPDATE As Boolean
'
'    On Error GoTo ErrHandle
'    tSkuIndexTmp = ""
'
'    tWhr = "WHERE FTTmnNum = ''" & tVB_CNTerminalNum & "''"
'    tWhr = tWhr & vbCrLf & "AND FTShdTransNo = ''" & tVB_Running & "''"
'    tWhr = tWhr & vbCrLf & "AND FTShdTransType = ''" & tVB_Transaction & "''"
'    tWhr = tWhr & vbCrLf & "AND FDShdTransDate = ''" & tUT_SaleDate & "''"
'    tWhr = tWhr & vbCrLf & "AND FTTdmCode = ''T006''"
'    tWhr = tWhr & vbCrLf & "AND FTVendorCnpFlag = ''1''"
'    tWhr = tWhr & vbCrLf & "ORDER BY FTSkuIndex,FNSrcSeqNo"
'    tSql = "STP_CNxSQLFunction @ptFunc = 'FNSrcSeqNo,FCSrcNet,FTSkuIndex' ,@ptTable = '" & tVB_CNTblRC & "',@ptWhere = '" & tWhr & "'"
'
'    Call SP_SQLvExecute(tSql, orsRC)
'    If Not (orsRC.BOF And orsRC.EOF) Then
'        orsRC.MoveFirst
'        tSkuIndexTmp = SP_FEDtChkString(orsRC, "FTSkuIndex")
'        cSrcNet = 0
'        tWhr = "WHERE FTTmnNum = ''" & tVB_CNTerminalNum & "''"
'        tWhr = tWhr & vbCrLf & "AND FTShdTransNo = ''" & tVB_Running & "''"
'        tWhr = tWhr & vbCrLf & "AND FTShdTransType = ''" & tVB_Transaction & "''"
'        tWhr = tWhr & vbCrLf & "AND FDShdTransDate = ''" & tUT_SaleDate & "''"
'        tWhr = tWhr & vbCrLf & "ORDER BY FNSdtSeqNo ASC"
'        tSql = "STP_CNxSQLFunction @ptFunc = 'FNSdtSeqNo,FCSdtSaleAmt,FTSkuIndex' ,@ptTable = '" & tVB_CNTblDT & "',@ptWhere = '" & tWhr & "'"
'        Call SP_SQLvExecute(tSql, orsDT)
'        If Not (orsDT.BOF And orsDT.EOF) Then
'            orsDT.MoveFirst
''            If aVB_TenderMnu(5).StaChange Then              'Allow
'            Do While Not orsRC.EOF
'
'                tSkuIndex = SP_FEDtChkString(orsRC, "FTSkuIndex")
'                If tSkuIndexTmp = tSkuIndex Then
'                    cSrcNet = cSrcNet + SP_FEDvChkNumeric(orsRC, "FCSrcNet")                'รวมค่า Vendor
'                Else
'
'                    orsDT.Filter = "FTSkuIndex ='" & tSkuIndexTmp & "'"                            'Index <> ไม่เท่ากัน
'                    If orsDT.RecordCount > 0 Then
'                        cSaleAmt = SP_FEDvChkNumeric(orsDT, "FCSdtSaleAmt")
'                        If cSrcNet > cSaleAmt Then
'                            cChange = cSrcNet - cSaleAmt
'                            cChange = SP_VALnText2Double(cChange)
'                            nPos = orsRC.AbsolutePosition
'                            orsRC.Filter = "FTSkuIndex ='" & tSkuIndexTmp & "'"
'                            tSort = "FCSrcNet DESC,FNSrcSeqNo"
'                            orsRC.Sort = tSort
'                            If orsRC.RecordCount > 0 Then
'                                If SP_RSTbClone2Offline(orsRC, orsTemp) Then
'                                    orsTemp.MoveFirst
'                                    bChkUPDATE = True
'                                    Do While Not (orsTemp.EOF) And bChkUPDATE
'                                        cSrcNetTmp = SP_FEDvChkNumeric(orsTemp, "FCSrcNet")                     'RC
'                                        nSeqNo = SP_FEDvChkNumeric(orsTemp, "FNSrcSeqNo")
'                                        If cChange > cSrcNetTmp Then
'                                            tSql = "UPDATE " & tVB_CNTblRC
'                                            If aVB_TenderMnu(5).StaChange Then              'Allow     Vendor Allow Change
'                                                tSql = tSql & vbCrLf & "SET FCSrcAmt=0"
'                                                tSql = tSql & vbCrLf & ",FCSrcNet=0"
'                                                tSql = tSql & vbCrLf & ",FCSrcChg=" & cSrcNetTmp
'                                            Else
''                                                    tSql = tSql & vbCrLf & "SET FCSrcAmt=FCSrcAmt"
''                                                    tSql = tSql & vbCrLf & ",FCSrcNet=0"
'                                                tSql = tSql & vbCrLf & ",SET FCSrcNet=0"
'                                                tSql = tSql & vbCrLf & ",FCSrcChg=0"
'                                            End If
'                                            cChange = cChange - cSrcNetTmp
'                                            bChkUPDATE = True
'                                        Else
'                                            tSql = "UPDATE " & tVB_CNTblRC
'                                            If aVB_TenderMnu(5).StaChange Then              'Allow     Vendor Allow Change
'                                                tSql = tSql & vbCrLf & "SET FCSrcAmt=FCSrcAmt"
'                                                tSql = tSql & vbCrLf & ",FCSrcNet=FCSrcNet - " & cChange
'                                                tSql = tSql & vbCrLf & ",FCSrcChg=" & cChange
'                                            Else
'                                                tSql = tSql & vbCrLf & "SET FCSrcAmt=FCSrcAmt"
'                                                tSql = tSql & vbCrLf & ",FCSrcNet=FCSrcNet - " & cChange
'                                                tSql = tSql & vbCrLf & ",FCSrcChg=0"
'                                            End If
'                                            bChkUPDATE = False
'                                        End If
'                                        tSql = tSql & vbCrLf & ",FTSrcRmk='Adjust'"
'                                        tSql = tSql & vbCrLf & "WHERE FNSrcSeqNo=" & nSeqNo
'                                        tSql = tSql & vbCrLf & "AND FTTmnNum = '" & tVB_CNTerminalNum & "'"
'                                        tSql = tSql & vbCrLf & "AND FTShdTransNo = '" & tVB_Running & "'"
'                                        tSql = tSql & vbCrLf & "AND FTShdTransType = '" & tVB_Transaction & "'"
'                                        tSql = tSql & vbCrLf & "AND FDShdTransDate = '" & tUT_SaleDate & "'"
'                                        Call SP_SQLvExecute(tSql)
'                                        orsTemp.MoveNext
'                                    Loop
'                                End If
'                            End If
'                            orsRC.Filter = adFilterNone
'                            tSort = "FTSkuIndex,FNSrcSeqNo"
'                            orsRC.Sort = tSort
'                            orsRC.AbsolutePosition = nPos
'                        End If
'
'                    End If          'If orsDT.RecordCount > 0 Then
'                    orsDT.Filter = adFilterNone
'
'                    cSrcNet = 0
'                    cSrcNet = cSrcNet + SP_FEDvChkNumeric(orsRC, "FCSrcNet")
'                    tSkuIndexTmp = SP_FEDtChkString(orsRC, "FTSkuIndex")
'                End If                          'If tSkuIndexTmp = tSkuIndex Then
'
'                orsRC.MoveNext
'            Loop
'            If orsRC.EOF Then               'Record สุดท้าย    เช็ค Record สุดท้าย
'                orsDT.Filter = "FTSkuIndex ='" & tSkuIndexTmp & "'"                            'Index <> ไม่เท่ากัน
'                If orsDT.RecordCount > 0 Then
'                    cSaleAmt = SP_FEDvChkNumeric(orsDT, "FCSdtSaleAmt")
'                    If cSrcNet > cSaleAmt Then
'                        cChange = cSrcNet - cSaleAmt
'                        cChange = SP_VALnText2Double(cChange)
'                        orsRC.Filter = "FTSkuIndex ='" & tSkuIndexTmp & "'"
'                        tSort = "FCSrcNet DESC,FNSrcSeqNo"
'                        orsRC.Sort = tSort
'                        If orsRC.RecordCount > 0 Then
'                            If SP_RSTbClone2Offline(orsRC, orsTemp) Then
'                                orsTemp.MoveFirst
'                                bChkUPDATE = True
'                                Do While Not (orsTemp.EOF) And bChkUPDATE
'                                    cSrcNetTmp = SP_FEDvChkNumeric(orsTemp, "FCSrcNet")                     'RC
'                                    nSeqNo = SP_FEDvChkNumeric(orsTemp, "FNSrcSeqNo")
'                                    If cChange > cSrcNetTmp Then
'                                        tSql = "UPDATE " & tVB_CNTblRC
'                                        If aVB_TenderMnu(5).StaChange Then              'Allow     Vendor Allow Change
'                                            tSql = tSql & vbCrLf & "SET FCSrcAmt=0"
'                                            tSql = tSql & vbCrLf & ",FCSrcNet=0"
'                                            tSql = tSql & vbCrLf & ",FCSrcChg=" & cSrcNetTmp
'                                        Else
''                                                    tSql = tSql & vbCrLf & "SET FCSrcAmt=FCSrcAmt"
''                                                    tSql = tSql & vbCrLf & ",FCSrcNet=0"
'                                            tSql = tSql & vbCrLf & ",SET FCSrcNet=0"
'                                            tSql = tSql & vbCrLf & ",FCSrcChg=0"
'                                        End If
'                                        cChange = cChange - cSrcNetTmp
'                                        bChkUPDATE = True
'                                    Else
'                                        tSql = "UPDATE " & tVB_CNTblRC
'                                        If aVB_TenderMnu(5).StaChange Then              'Allow     Vendor Allow Change
'                                            tSql = tSql & vbCrLf & "SET FCSrcAmt=FCSrcAmt"
'                                            tSql = tSql & vbCrLf & ",FCSrcNet=FCSrcNet - " & cChange
'                                            tSql = tSql & vbCrLf & ",FCSrcChg=" & cChange
'                                        Else
'                                            tSql = tSql & vbCrLf & "SET FCSrcAmt=FCSrcAmt"
'                                            tSql = tSql & vbCrLf & ",FCSrcNet=FCSrcNet - " & cChange
'                                            tSql = tSql & vbCrLf & ",FCSrcChg=0"
'                                        End If
'                                        bChkUPDATE = False
'                                    End If
'                                    tSql = tSql & vbCrLf & ",FTSrcRmk='Adjust'"
'                                    tSql = tSql & vbCrLf & "WHERE FNSrcSeqNo=" & nSeqNo
'                                    tSql = tSql & vbCrLf & "AND FTTmnNum = '" & tVB_CNTerminalNum & "'"
'                                    tSql = tSql & vbCrLf & "AND FTShdTransNo = '" & tVB_Running & "'"
'                                    tSql = tSql & vbCrLf & "AND FTShdTransType = '" & tVB_Transaction & "'"
'                                    tSql = tSql & vbCrLf & "AND FDShdTransDate = '" & tUT_SaleDate & "'"
'                                    Call SP_SQLvExecute(tSql)
'                                    orsTemp.MoveNext
'                                Loop
'                            End If
'                        End If
'                    End If
'                End If
'            End If
'        End If              'If Not (orsDT.BOF And orsDT.EOF) Then
'    End If                  'If Not (orsRC.BOF And orsRC.EOF) Then
'    cVB_TotalVendor = 0
'    tWhr = "WHERE FTTmnNum = '" & tVB_CNTerminalNum & "'"
'    tWhr = tWhr & vbCrLf & "AND FTShdTransNo = '" & tVB_Running & "'"
'    tWhr = tWhr & vbCrLf & "AND FTShdTransType = '" & tVB_Transaction & "'"
'    tWhr = tWhr & vbCrLf & "AND FDShdTransDate = '" & tUT_SaleDate & "'"
'    tWhr = tWhr & vbCrLf & "AND FTTdmCode = 'T006'"
'    tWhr = tWhr & vbCrLf & "AND FTVendorCnpFlag = '1'"
'    cVB_TotalVendor = SP_VALcText2Double(SP_SQLtFunction("SUM", "FCSrcNet", tVB_CNTblRC, tWhr))
'Exit Sub
'ErrHandle:
'    Call SP_MSGxShowError(Err.Number, Err.Description)
'End Sub

Private Sub Form_Activate()
    tVB_CNTblDT = tVB_CNTblDTTmp
    tVB_CNTblCD = tVB_CNTblCDTmp
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'*KT 51-11-04
    Dim tWhe$
On Error GoTo ErrHandle
    If bVB_CNSplash Then KeyCode = 0 '*KT 51-09-17 ป้องการ Key กรณีที่ Splash ขึ้น '*KT 51-10-21 V 3.0.140/V 3.0.200
'    Call W_CHKxDiscount91(KeyCode, Shift)    '*Eaw 56-01-05
'*KT 51-11-17
'Debug.Print "11"
    '----------------------------------------------------
    If bVB_CNRecovryAct Then
        KeyCode = 0
        Exit Sub
    End If
    '--------------------------------------------------
    Select Case KeyCode
        Case vbKeyHome
            nVB_OnTopCount = 0
'            Debug.Print "7"
            ReDim aVB_CreditCrd(0) As String
            aVB_CreditCrd(0) = ""
            tWhe = "  WHERE FTTmnNum=''" & tVB_CNTerminalNum & "''"
            tWhe = tWhe & "  AND FTShdTransNo=''" & tVB_Running & "''"
            tWhe = tWhe & "  AND FDShdTransDate =''" & tUT_SaleDate & "''"
            tWhe = tWhe & "  AND FTShdTransType=''" & tVB_Transaction & "''"
            tWhe = tWhe & "  AND ISNULL(FTSdtStkCode,'''') <> ''1''" '*KT 52-11-25 Deal&Promo
            cVB_TotalDisc = SP_FMTcCurPoint(SP_VALcText2Double(SP_SQLtFunctionSTP("SUM", "FCScdAmt", tVB_CNTblCD, tWhe)), True, False)
            
            bVB_CNTrnDisc = True
            Select Case Val(Format(tVB_TransactionSub))
                Case tEN_TrnDeposit 'Deposit Create Transaction
                    tWhe = " WHERE  FTTmnNum=''" & tVB_CNTerminalNum & "'' AND FTShdTransNo=''" & tVB_Running & "''"
                    tWhe = tWhe & " AND  FTShdTransType=''" & tVB_Transaction & "'' AND FDShdTransDate=''" & tUT_SaleDate & "''"
                    tWhe = tWhe & " AND  FTSdtCancelFlag <> ''1''"
                    cUT_SubTotal = SP_FMTcCurPoint(SP_VALcText2Double(SP_SQLtFunctionSTP("SUM", "FCSdtSaleAmt", tVB_CNTblDT, tWhe)), True, False)
                    Call W_DEFxCheckVendor     '*Ao 53-10-19
                     ''*Eaw 56-02-06 การ Recovery ข้อมูล
                    If Not bVB_CNPaid Then
                        wCNDepositAmt.Show vbModal
                        If Not bVB_TrnDeposit Then
                            Exit Sub
                        End If
                    End If
                    '------------------------------------------------------
'*KT 66-08-28 CommSheet TTK-ISS3000-829_SIT-002 Question Refer Layaway Payment Deposit Amount 7,8,9.xlsx
'เอาออกเพื่อไม่ให้เรียก Form ซ้ำ
'****************************************
'                    wCNDepositAmt.Show vbModal
'                    If Not bVB_TrnDeposit Then
'                        Exit Sub
'                    End If
'****************************************
            End Select
            cVB_TotalDiscAft = SP_VALcText2Double(ogdForm.Aggregate(flexSTSum, ogdForm.FixedRows, 4, ogdForm.rows - 1, 4))
            cVB_TotalDiscMallCrd = W_SUMcDiscByType(nEN_Disc11MallCrd)
            cVB_TotalDiscEmpCrd = W_SUMcDiscByType(nEN_Disc12EmpCrd)
'            Call W_DEFxCheckVendor     '*Ao 53-10-19
            Unload Me
    End Select
    Exit Sub
ErrHandle:
    Call SP_CRTxLogPrg("ErrorNumber:" & Err.Number & " ErrorDesc:" & Err.Description & " Form:" & Me.Name & " Control:Form" & " Events:Form_KeyDown")
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If bVB_CNPaid Then  '*Eaw 56-02-06 ถ้ามีการจ่ายเงินบางส่วนแล้วให้จ่ายให้จบบิลได้เท่านั้น
        If KeyAscii = vbKeyEscape Then
            Call SP_MSGnShowing("tms_CN005085", nCS_Error)
            Exit Sub
        End If
        If KeyAscii <> vbKeyHome Then
            Call SP_MSGnShowing("tms_CN005107", nCS_Information)
            Exit Sub
        End If
    End If
    If bVB_CNDiscount Then '*Eaw 56-02-06 ถ้ามีการลดท้ายบิลแล้ว หากกด  Esc จะเครียร์ค่าที่เคยลดท้ายบิลทิ้งทั้งหมด
        If KeyAscii = vbKeyEscape Then
                Call W_PRCxClearCDRTmp
                bVB_CNDiscount = False
        End If
    End If
    '-----------------------------------------------------------------------
    If KeyAscii = 13 Then
        Call ocmOk_Click
    Else
        Select Case KeyAscii
            Case 48 To 58
                otmForm.Enabled = True
                With ogdFunc
                    tW_HotKey = tW_HotKey & Chr(KeyAscii)
                    If tW_HotKey <> "" Then
                         Call W_PRCxCallFunc(Val(tW_HotKey))
                    End If
                    If IsNumeric(tW_HotKey) Then
                        If CLng(tW_HotKey) >= nMaxRow Then tW_HotKey = "": otmForm.Enabled = False
                    End If
                End With
        End Select
    End If
    
    
End Sub

Private Sub Form_Load()
    Dim tWhe As String
    Dim tPrint$
    Dim nI As Long
On Error GoTo ErrHandle
    '*KT 51-11-20
    '---------------------------------------------------------------------------------------------------------------------------------------------------------
    otmRec.Enabled = False              '*Ao 2014/05/07  PRJ10001TK-PH2-External Spec-Incomplete Transaction-00.02.04 FSBT
    tVB_CNTblDT = tVB_CNTblDTTmp
    tVB_CNTblCD = tVB_CNTblCDTmp
    '*Tao 61-08-15 ไม่ต้องเคลียร์
'    tVB_MemPointID = ""                 '*Ao 2013/02/27  TK POINT
'    tVB_MemExp = ""
'    tVB_ChkMemExp = ""
    olaMEM.Caption = ""
    bVB_ChkEmpPoint = True  '*Ao 2013/09/23     ต้อง Set ค่าให้ คิด POINT
    
    '---------------------------------------------------------------------------------------------------------------------------------------------------------
    tW_HotKey = ""
    Me.Top = IIf(bVB_Res640X480, cCS_Top680x480, cCS_Top800x600)
    Me.Left = IIf(bVB_Res640X480, cCS_Left680x480, cCS_Left800x600)
    If Not bVB_Res640X480 Then
        Call SP_FRMxGetCaptionFrmDB2Tag(Me)
        Call SP_FRMxSetCapForm(Me)
    End If
    Call W_FNCxAddFunc2Grid
    nMaxRow = ogdFunc.Aggregate(flexSTMax, 0, 0, ogdFunc.rows - 1, 0)
    Call W_SHWxDiscountByTrans(0)
    
    tWhe = "  WHERE FTTmnNum=''" & tVB_CNTerminalNum & "''"
    tWhe = tWhe & "  AND FTShdTransNo=''" & tVB_Running & "''"
    tWhe = tWhe & "  AND FDShdTransDate =''" & tUT_SaleDate & "''"
    tWhe = tWhe & "  AND FNDctNo IN (1,2,3,4,5,6,21,22)" '*KT 52-09-07 Member + Consignment
    tWhe = tWhe & "  AND ISNULL(FTSdtStkCode,'''') <> ''1''"  '*KT 52-11-25 Deal&Promo
    cVB_TotalDiscB = SP_FMTcCurPoint(SP_VALcText2Double(SP_SQLtFunctionSTP("SUM", "FCScdAmt", tVB_CNTblCD, tWhe)), True, False)
    
    'Comsheet ML- SF- 425  Balance In Discount Screen Is Wrong When Price Override
    'cVB_BalAftDisc = cUT_Total - cVB_TotalDiscB - Val(Format(ogdForm.Aggregate(flexSTSum, 0, 4, ogdForm.rows - 1, 4)))
    cVB_BalAftDisc = (cUT_SubTotal) - CDbl(Val(Format(ogdForm.Aggregate(flexSTSum, 0, 4, ogdForm.rows - 1, 4))))
    olaBal.Caption = SP_FMTcCurPoint(cVB_BalAftDisc, False, True)
    olaTSubReg.Caption = SP_FMTcCurPoint(cUT_SubReg, False, True)
    olaTSubPro.Caption = SP_FMTcCurPoint(cUT_SubPro, False, True)
    olaTSubTotal.Caption = SP_FMTcCurPoint(cUT_SubTotal, False, True)
    '---------------------
    'เพื่อ Clear ค่าทุกครั้ง
    tVB_CNUpdateEmpLimit = "" '*TW 51-08-18 เก็บ Script สำหรับ Update Empployee Card Limit
    cVB_CNRemainDiscAmt = 0 '*KT 53-01-28 Communication Sheet PH3-ML-SB-074 Remain Discount for Employee.xls
    tVB_CNRemainDiscID = "" '*KT 53-01-28 Communication Sheet PH3-ML-SB-074 Remain Discount for Employee.xls
    '--------------------
    If Not bVB_CNCalAutoTran Then
        Select Case Val(Format(tVB_TransactionSub))
            Case tEN_TrnReturn, tEN_TrnReturnTaxInc, tEN_TrnReturnManualTax, tEN_TrnSaleTaxInc, tEN_TrnSaleManualTax '*KT 51-06-07 Comsheet ML-SF-327
                'none
            Case Else
                Call W_PRCxAutoTransDisc  'Auto Transaction Discount
        End Select

    End If
    'Customer Display
    If bVB_OPOSLineDisplayEnable Then
'        tPrint = "TOTAL;" & olaBal.Caption
        tPrint = "TOTAL : " & olaBal.Caption '*Eaw 56-07-01
        Call SP_PrnPollDisplay(tPrint)
    End If
    bVB_CNRecovryAct = False
    
    Select Case Val(Format(tVB_TransactionSub))
        Case tEN_TrnSaleTaxInc, tEN_TrnSaleManualTax, tEN_TrnReturnTaxInc, tEN_TrnReturnManualTax
            For nI = ogdFunc.FixedRows To ogdFunc.rows - 1
                If ogdFunc.TextMatrix(nI, 4) = "1" Then
                    ogdFunc.Select nI, 0
                    Exit For
                End If
            Next nI
    End Select
    
    '*Eaw 56-02-06  กรณีเป็นบิลที่กู้คืนมา แล้วบิลมีการให้ส่วนลดท้ายบิลไปแล้ว
    If bVB_IsRecovery Then
        If bVB_CNDiscount Then
            Call W_PRCxRecDiscount  'กรณีมีส่วนลดท้ายบิลแล้ว
        Else
            Call W_INSxTblCD   'กรณีมีแต่ส่วนลดรายตัว
        End If
        Call W_GETxMemberBP             '*Ao 2014/05/07  PRJ10001TK-PH2-External Spec-Incomplete Transaction-00.02.04 FSBT
    End If
    '-----------------------------------------------------
'    Call SP_PRCxUpdateDiscTable    '*Ao 2013/03/06
    If Not bVB_IsRecovery Then  '*Eaw 56-07-03 ถ้าเป็น Recovery ค่าจะผิดจึงให้ Update เฉพาะกรณีที่ส่วนลดนี้ไม่ได้มาจากการ Recovery
        Call SP_PRCxUpdateDiscTable
    Else
        If bVB_IsRecovery Then              '*Ao 2014/05/07  PRJ10001TK-PH2-External Spec-Incomplete Transaction-00.02.04 FSBT
            otmRec.Enabled = True
        End If
        bVB_IsRecovery = False
    End If
    
    Exit Sub
ErrHandle:
    Call SP_CRTxLogPrg("ErrorNumber:" & Err.Number & " ErrorDesc:" & Err.Description & " Form:" & Me.Name & " Control:Form" & " Events:Form_Load")
End Sub

Private Sub W_FNCxAddFunc2Grid()
'--------------------------------------------------------
'Call:
'Cmt: แสดงรายการส่วนลดประเภทต่างๆ
'-------------------------------------------------------
    Dim nI As Long
    Dim nR As Long
On Error GoTo ErrHandle

    With ogdFunc
        nR = LBound(aVB_DiscMnu)
      .rows = UBound(aVB_DiscMnu)
        For nI = LBound(aVB_DiscMnu) To UBound(aVB_DiscMnu) - 1
             If aVB_DiscMnu(nI).Show Then
                If aVB_DiscMnu(nI).SmnID <> "1829" And aVB_DiscMnu(nI).SmnID <> "1830" Then 'Line Item Discount  และ Line Item Markdown ไม่นำมาแสดง
'                    .TextMatrix(nR, 0) = aVB_DiscMnu(nI).ID
'                    .TextMatrix(nR, 1) = aVB_DiscMnu(nI).ID & "." & IIf(aVB_DiscMnu(nI).NameUsr <> "", aVB_DiscMnu(nI).NameUsr, aVB_DiscMnu(nI).Name)
                    .TextMatrix(nR, 0) = nR + 1             '*Ao 2013/05/16   แก้ไขการเรียง เมณู
                    .TextMatrix(nR, 1) = nR + 1 & "." & IIf(aVB_DiscMnu(nI).NameUsr <> "", aVB_DiscMnu(nI).NameUsr, aVB_DiscMnu(nI).Name)
                    .TextMatrix(nR, 2) = aVB_DiscMnu(nI).FuncName
                    .TextMatrix(nR, 3) = aVB_DiscMnu(nI).SmnID
                    If bVB_ReEntryMode Then
                        Select Case Val(Format(tVB_TransactionSub))
                            Case tEN_TrnSaleTaxInc, tEN_TrnReturnTaxInc
'                                If aVB_DiscMnu(ni).SmnID = "1809" Or aVB_DiscMnu(ni).SmnID = "1810" Then
                                If aVB_DiscMnu(nI).SmnID = "1809" Or aVB_DiscMnu(nI).SmnID = "1810" Or aVB_DiscMnu(nI).SmnID = "1850" Or aVB_DiscMnu(nI).SmnID = "1887" Then  '*Eaw 56-06-27 เพิ่ม BP กับ TC CommSheet TK-ISS3000-244-TR.xlsx
                                    .TextMatrix(nR, 4) = "1"
                                    If aVB_DiscMnu(nI).SmnID = "1809" Then
                                        .Select nR, 0
                                    End If
                                Else
                                    .TextMatrix(nR, 4) = "0"
                                    Call W_SetBackColor(ogdFunc, nR, 13619151)
                                End If
                            Case tEN_TrnSaleManualTax, tEN_TrnReturnManualTax
                                .TextMatrix(nR, 4) = "0"
                                Call W_SetBackColor(ogdFunc, nR, 13619151)
                            Case Else
                                .TextMatrix(nR, 4) = "1"
                        End Select
                    End If
                    nR = nR + 1
                End If
            End If
        Next nI
        .rows = nR
    End With
    Exit Sub
ErrHandle:
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Set wCNSubDisc2 = Nothing
On Error GoTo ErrHandle
'Debug.Print "12"
    '----------------------------------------------------
    If bVB_CNRecovryAct Then
        Cancel = 1
        Exit Sub
    End If
    '--------------------------------------------------
    '*KT 51-11-04 ***KT 51-11-20
    '------------------------------------------------------------------------------------------------------------------------
    '*KT 51-12-09 กรณีที่จะทำ Zero Trans ต้องใช้ bVB_CNTrnDisc เป็นเงื่อนไขร่วมเพราะ Event Upload นั้นเมื่อกด Cancel ก็จะเข้า Event นี้ด้วย
    'เกิด Error เมื่อทำการลดส่วนลดทั้งหมดในหน้าจอ Discount แล้วกด Esc หรือ ปุ่ม cancel เพื่อกลับเข้าหน้า Item Entry เกิดการเข้า Zero Trans ทันที
    If SP_VALcText2Double(olaBal.Caption) <= 0 And bVB_CNTrnDisc Then
        '*KT 52-01-28 V 1.5.3 กรณีที่เป็น Deposit ให้กลับไปทำงานที่ wFunction
        Select Case Val(Format(tVB_TransactionSub))
            Case tEN_TrnDeposit
                bVB_CNZero = True
                Call SP_OBJxSetNothing(Me)
                Exit Sub
            Case Else
                Call SP_MSGnShowing("tms_CN005096", nCS_Warning)
                cVB_Pay = 0
                cVB_Balance = 0
                cVB_NetSale = 0
                cVB_RoundRcv = 0
                cVB_Change = 0
                tVB_CrdCardNo = ""
                tVB_CrdApprCode = ""
                Call SP_CRTxLogCashBalEvent(Me.Name, "Form_Unload[1]", "Cash Tender", cVB_Balance)
                Select Case Val(Format(tVB_TransactionSub))
                    Case tEN_TrnReturn, tEN_TrnCancelVoucher, tEN_TrnDepositCancel, tEN_TrnReturnTaxInc, tEN_TrnReturnManualTax
                        tVB_TenderCode = aVB_TenderMnu(0).ID
                        tVB_TenderType = aVB_TenderMnu(0).Name
                        tVB_TenderName = aVB_TenderMnu(0).DescReturn
                    Case Else
                        tVB_TenderCode = aVB_TenderMnu(0).ID
                        tVB_TenderType = aVB_TenderMnu(0).Name
                        tVB_TenderName = aVB_TenderMnu(0).DescSale
                 End Select
                Call wFunctionKB.W_KBRcvCash
                bVB_FormPayAct = False
                tVB_CNTblDT = "TPSDT" & tVB_CNTerminalNum
                tVB_CNTblCD = "TPSCD" & tVB_CNTerminalNum
                bVB_CNZero = True
        End Select
    Else
        bVB_CNZero = False
    End If
    '------------------------------------------------------------------------------------------------------------------
     otmForm.Enabled = False
     otmLock.Enabled = False
     otmRec.Enabled = False
'     Call SP_PRCxUpdateDiscountable                 '*Ao 2013/10/02     CommSheet TK-ISS3000-371-FAPL-132
     
     Call SP_OBJxSetNothing(Me)
     
     Exit Sub
'     Debug.Print "10"
ErrHandle:
     Call SP_CRTxLogPrg("ErrorNumber:" & Err.Number & " ErrorDesc:" & Err.Description & " Form:" & Me.Name & " Control:Form" & " Events:Form_Unload")
End Sub

Private Sub ocmCancel_Click()
Dim tMsg As String
On Error GoTo ErrHandle
    
    If SP_CHKbDiscEmp Then                   '*Ao 2014/04/29 PRJ10001TK-PH2-External Spec-Staff Discount-00.02.03 FSBT.docx
        tMsg = "No ESC allowed with Staff Discount given"
        MsgBox tMsg, vbOKOnly, tVB_CNAppTitle
        Exit Sub
    End If
    oDisc.dTranDate = tUT_SaleDate
    oDisc.tTerminalNo = tVB_CNTerminalNum
    oDisc.tTransNo = tVB_Running
    oDisc.DISxClearTrans
    bVB_CNTrnDisc = False
    bVB_CNTrnDiscEx = False
    bVB_FormPayAct = False
'*Eaw 56-07-02 CommSheet TK-ISS3000-249-TR.xls ไม่ต้อง Recovery RC
'    '*Eaw 56-02-06  ถ้ามีการชำระบางส่วนแล้ว ห้ามกด Esc ในหน้า Discount
'    If bVB_CNPaid Then
'        Call SP_MSGnShowing("tms_CN005085", nCS_Error)
'        Exit Sub
'    End If
    '------------------------------------------------------------------------------------------
    '*KT 51-12-15 Phase 2 ข้อ 27  Communication Sheet PH2-ML-SF-028 Sale amount incorrect when Recall Transaction.xls
    'มีบางครั้งที่กด Cancel แล้วไม่เข้า Form_Activate ของ wMianRet
    tVB_CNTblDT = "TPSDT" & tVB_CNTerminalNum
    tVB_CNTblCD = "TPSCD" & tVB_CNTerminalNum
    Call SP_PRCxClearMemberPoint
    '-----------------------------------------------------------
    nVB_DiscMallCount = 0
    nVB_RegMallCount = 0
    nVB_CurMemPointCount = 0
    nVB_MemPointMnu = 0
    tVB_CNUpdateEmpLimit = "" '*KT 53-10-26 หมายเลข Commsheet PH3-ML-SF-133
    cVB_CNRemainDiscAmt = 0 '*KT 53-10-26 หมายเลข Commsheet PH3-ML-SF-133
    tVB_CNRemainDiscID = "" '*KT 53-10-26 หมายเลข Commsheet PH3-ML-SF-133
    Call W_CHKxReVendor    '*Ao 2011/01/17   CommSheet ML-PRV-020-Change Tender is wrong concept.xls
    Call W_PRCxClearCDRTmp   '*Eaw 56-02-06 ถ้ามีการลดท้ายบิลแล้ว หากกด  Esc จะเครียร์ค่าที่เคยลดท้ายบิลทิ้งทั้งหมด
    bVB_CNDiscount = False   '*Eaw 56-02-06
    bVB_ChkEmpPoint = True '*Eaw 56-08-21
    bMainFunc = True                                   '*Ao 2014/09/05   V 4.2.143
    Unload Me
    Exit Sub
ErrHandle:
    If Not otmForm.Enabled Then otmForm.Enabled = True
    If Not otmLock.Enabled Then otmLock.Enabled = True
    Call SP_CRTxLogPrg("ErrorNumber:" & Err.Number & " ErrorDesc:" & Err.Description & " Form:" & Me.Name & " Control:ocmCancel" & " Events:ocmCancel_Click")
End Sub

Private Sub ocmOk_Click()
'    Dim cTempAmt As Double
'    Dim cTempPer As Double
    Dim tPrint$
On Error GoTo ErrHandle
    With ogdFunc
        nW_TransMnu = .TextMatrix(.Row, 0)
        tW_TransFunc = .TextMatrix(.Row, 2)
        tW_SmnID = .TextMatrix(.Row, 3)
        tVB_CNTrnDiscType = tW_SmnID
        
        tVB_CNMnuCurrent = tW_SmnID
        If Not SP_USRbMnuAccept(nEN_Usr2AthRead, tVB_CNMnuCurrent) Then
'             DoEvents '*KT 53-01-14 DoEvents
             If tVB_CNUserAlwC = "" Then
                 Exit Sub
             End If
        End If
        
        '*KT 54-09-13 V.4.0.1 Req3 Phase4.0(Member Point) Member Point on Other Receive หน้าจอ Discount สำหรับ Transaction Other Receive ให้ทำรายการ Register Member Point ได้อย่างเดียว
        Select Case Val(Format(tVB_TransactionSub))
            Case tEN_TrnOtherReceive
                If nW_TransMnu <> 8 Then
                    Call SP_MSGnShowing("tms_CN005107", nCS_Information)
                    Exit Sub
                End If
            Case tEN_TrnSaleTaxInc, tEN_TrnSaleManualTax
                If .TextMatrix(.Row, 4) = "0" Then
                    Exit Sub
                End If
            '_________________________________________________________________________________________
            Case tEN_TrnReturn, tEN_TrnReturnTaxInc, tEN_TrnReturnManualTax                    '*Ao 2012/11/26  Return & Return Refer Sale ให้เข้าหน้า Discount เหมือนขาย
                Select Case tW_SmnID            '1809:Menber 1887:TC Card 1810:Employee Card
                    Case "1809", "1887", "1810", "1850"
                    
                    Case Else
                        Call SP_MSGnShowing("tms_CN005107", nCS_Information)
                        otmForm.Enabled = True
                        otmLock.Enabled = True
                        Exit Sub
                End Select
            '_________________________________________________________________________________________
        End Select
        
       otmForm.Enabled = False
       otmLock.Enabled = False
       Select Case tW_SmnID 'nW_TransMnu
            Case "1809", "1887" '1 'Mall Card/Member Card              1809 : Member card   1887 :TC Card
                tVB_CNSmnID = tW_SmnID '*Eaw 56-07-03 CommSheet TK-ISS3000-263.xls
                If oDisc.DISbChkAlw(nEN_Disc11MallCrd, False) Then
                     If nVB_DiscMallCount >= 3 Then
                        Call SP_MSGnShowing("tms_CN005133", nCS_Warning)
                        otmForm.Enabled = True
                        otmLock.Enabled = True
                        Exit Sub
                    End If
                    '*KT 55-07-10 ระบบ Trackdata number of MSR
                    If tW_SmnID = "1887" Then '1887 = TC Card
                        wCNDiscMall.nW_TrackNo = 4
                    Else
                         wCNDiscMall.nW_TrackNo = 2
                    End If
                     wCNDiscMall.Show vbModal
                      If bVB_DiscAccept Then 'Comsheet ML-SF-184
                        If bVB_OPOSLineDisplayEnable Then
                            tPrint = "DISC;" & SP_FMTcCurPoint(oDisc.cLastDiscount, False)
                            Call SP_PrnPollDisplay(tPrint)
                        End If
                        Call W_SHWxDiscountByTrans(oDisc.cLastDiscount)
                     End If
                Else
                    Call SP_MSGnShowing("tms_CN005178", nCS_Warning)
                    otmForm.Enabled = True
                    otmLock.Enabled = True
                    Exit Sub
                End If
            Case "1810" '2 'Employee Card
                If oDisc.DISbChkAlw(nEN_Disc12EmpCrd, False) Then
'                    If Not W_CHKbMinDisc(cVB_BalAftDisc) Then '*Eaw 56-08-23 CommSheet TK-ISS3000-265-TR_Req.xls
                    If bVB_ChkMinEmpDisc Then
                        If cVB_BalAftDisc < cVB_MinValue Then
                            Call SP_MSGnShowing("tms_CN005249", nCS_Warning)
                            otmForm.Enabled = True
                            otmLock.Enabled = True
                            Exit Sub
                        End If
                    End If
                    If Not SP_CHKbEmpPolicy(tVB_CNTblCD, tVB_CNWheBill) Then               '*Ao 2013/07/26
                        Call SP_MSGnShowing("tms_CN005179", nCS_Warning)
                        otmForm.Enabled = True
                        otmLock.Enabled = True
                        Exit Sub
                    End If
                    wCNDiscEmp.Show vbModal
                    If oDisc.cLastDiscount >= 0 And bVB_DiscAccept Then
                        If bVB_OPOSLineDisplayEnable Then
                            tPrint = "DISC;" & SP_FMTcCurPoint(oDisc.cLastDiscount, False)
                            Call SP_PrnPollDisplay(tPrint)
                        End If
                        Call W_SHWxDiscountByTrans(oDisc.cLastDiscount)
                        cVB_CNDiscEmpAmt = oDisc.cLastDiscount
'                    Else
'                        Call W_CHKxDiscEmp     '*Ao 2013/10/29
                    End If
                Else
                    Call SP_MSGnShowing("tms_CN005179", nCS_Warning)
                    otmForm.Enabled = True
                    otmLock.Enabled = True
                    Exit Sub
                End If
            Case "1811" '3 'Cooperate Coupon
                 If oDisc.DISbChkAlw(nEN_Disc09CoopCpn, True) Then
                    If W_PRCbChkItemDiscCpCpnAlw(nEN_Disc09CoopCpn) Then
                        wCNDiscCoupon.Caption = "Cooperate Coupon"
                        wCNDiscCoupon.Show vbModal
                        If oDisc.cLastDiscount >= 0 And bVB_DiscAccept Then
                            If bVB_OPOSLineDisplayEnable Then
                                tPrint = "DISC;" & SP_FMTcCurPoint(oDisc.cLastDiscount, False)
                                Call SP_PrnPollDisplay(tPrint)
                            End If
                            Call W_SHWxDiscountByTrans(oDisc.cLastDiscount, nEN_Disc09CoopCpn)
                        End If
                    Else
                        Call SP_MSGnShowing("tms_CN005101", nCS_Error) '*KT 51-04-19 Comsheet ML-SF-225
                    End If
                Else
                    Call SP_MSGnShowing("tms_CN005180", nCS_Warning)
                    otmForm.Enabled = True
                    otmLock.Enabled = True
                    Exit Sub
                End If
            Case "1812" '4 'Cooperate Brand
                If oDisc.DISbChkAlw(nEN_Disc10CoopBrand, True) Then
                     If W_PRCbChkItemDiscCpCpnAlw(nEN_Disc10CoopBrand) Then
                        wCNDiscCoupon.Caption = "Cooperate Brand"
                        wCNDiscCoupon.Show vbModal
                        If oDisc.cLastDiscount >= 0 And bVB_DiscAccept Then
                            If bVB_OPOSLineDisplayEnable Then
                                tPrint = "DISC;" & SP_FMTcCurPoint(oDisc.cLastDiscount, False)
                                Call SP_PrnPollDisplay(tPrint)
                            End If
                            Call W_SHWxDiscountByTrans(oDisc.cLastDiscount, nEN_Disc09CoopCpn)
                        End If
                    Else
                        Call SP_MSGnShowing("tms_CN005102", nCS_Error) ''*KT 51-04-19 Comsheet ML-SF-225
                    End If
                Else
                    Call SP_MSGnShowing("tms_CN005181", nCS_Warning)
                    otmForm.Enabled = True
                    otmLock.Enabled = True
                    Exit Sub
                End If
            Case "1813" '5 'Mall Coupon or Discount Coupon
                If oDisc.DISbChkAlw(nEN_Disc08MallCpn, False) Then
'                    If oDisc.DISbChkMallCpnMinAlw(cVB_BalAftDisc) Then '*KT 54-07-28 V.4.0.1 Req.6 Phase4.0(Member Point) Promotion Coupon
                        wCNDiscCpn.Show vbModal
                        If oDisc.cLastDiscount >= 0 And bVB_DiscAccept Then
                            If bVB_OPOSLineDisplayEnable Then
                                tPrint = "DISC;" & SP_FMTcCurPoint(oDisc.cLastDiscount, False)
                                Call SP_PrnPollDisplay(tPrint)
                            End If
                            Call W_SHWxDiscountByTrans(oDisc.cLastDiscount)
                        End If
'                    Else
'                        Call SP_MSGnShowing("tms_CN005225", nCS_Warning)
'                        otmForm.Enabled = True
'                        otmLock.Enabled = True
'                        Exit Sub
'                    End If
                Else
                    Call SP_MSGnShowing("tms_CN005182", nCS_Warning)
                    otmForm.Enabled = True
                    otmLock.Enabled = True
                    Exit Sub
                End If
            Case "1814" '6 'Transaction Discount
                If oDisc.DISbChkAlw(nEN_Disc13TranDisc, False) Then
                    wCNTDisc2.Show vbModal
                    If oDisc.cLastDiscount >= 0 And bVB_DiscAccept Then
                        If bVB_OPOSLineDisplayEnable Then
                            tPrint = "DISC;" & SP_FMTcCurPoint(oDisc.cLastDiscount, False)
                            Call SP_PrnPollDisplay(tPrint)
                        End If
                        Call W_SHWxDiscountByTrans(oDisc.cLastDiscount)
                    End If
                Else
                    Call SP_MSGnShowing("tms_CN005183", nCS_Warning)
                    otmForm.Enabled = True
                    otmLock.Enabled = True
                    Exit Sub
                End If
            Case "1815" '7 'Transaction Markdown
                 If oDisc.DISbChkAlw(nEN_Disc14TranMDown, False) Then
                    wCNTDisc2.Show vbModal
                    If oDisc.cLastDiscount >= 0 And bVB_DiscAccept Then
                        tPrint = "DISC;" & SP_FMTcCurPoint(oDisc.cLastDiscount, False)
                        If bVB_OPOSLineDisplayEnable Then
                            Call SP_PrnPollDisplay(tPrint)
                        End If
                         Call W_SHWxDiscountByTrans(oDisc.cLastDiscount)
                    End If
                Else
                    Call SP_MSGnShowing("tms_CN005184", nCS_Warning)
                    otmForm.Enabled = True
                    otmLock.Enabled = True
                    Exit Sub
                End If
            Case "1850" '8 'Mall Card Point /Bonus Point Award
                If nVB_RegMallCount >= 3 Then
                    Call SP_MSGnShowing("tms_CN005133", nCS_Warning)
                    otmForm.Enabled = True
                    otmLock.Enabled = True
                    Exit Sub
                End If
                wCNDiscMall.nW_TrackNo = 2
                wCNDiscMall.Show vbModal
            Case "1893" '9 'Transaction markdown1
                If oDisc.DISbChkAlw(nEN_Disc23TranMDown1, False) Then
                    wCNTMkdNo.Show vbModal
                    If tVB_CNTranMDownNo <> "" Then
                        wCNTDisc2.Show vbModal
                        If oDisc.cLastDiscount >= 0 And bVB_DiscAccept Then
                            tPrint = "DISC;" & SP_FMTcCurPoint(oDisc.cLastDiscount, False)
                            If bVB_OPOSLineDisplayEnable Then
                                Call SP_PrnPollDisplay(tPrint)
                            End If
                             Call W_SHWxDiscountByTrans(oDisc.cLastDiscount)
                        End If
                    End If
                Else
                    Call SP_MSGnShowing("tms_CN005184", nCS_Warning)
                    otmForm.Enabled = True
                    otmLock.Enabled = True
                    Exit Sub
                End If
            Case "1892" '10 'Transaction markdown2
                If oDisc.DISbChkAlw(nEN_Disc24TranMDown2, False) Then
                    wCNTMkdNo.Show vbModal
                    If tVB_CNTranMDownNo <> "" Then
                        wCNTDisc2.Show vbModal
                        If oDisc.cLastDiscount >= 0 And bVB_DiscAccept Then
                            tPrint = "DISC;" & SP_FMTcCurPoint(oDisc.cLastDiscount, False)
                            If bVB_OPOSLineDisplayEnable Then
                                Call SP_PrnPollDisplay(tPrint)
                            End If
                             Call W_SHWxDiscountByTrans(oDisc.cLastDiscount)
                        End If
                    End If
                Else
                    Call SP_MSGnShowing("tms_CN005184", nCS_Warning)
                    otmForm.Enabled = True
                    otmLock.Enabled = True
                    Exit Sub
                End If
            Case "1891" ' 11 'Transaction markdown3
                If oDisc.DISbChkAlw(nEN_Disc25TranMDown3, False) Then
                    wCNTMkdNo.Show vbModal
                    If tVB_CNTranMDownNo <> "" Then
                         wCNTDisc2.Show vbModal
                        If oDisc.cLastDiscount >= 0 And bVB_DiscAccept Then
                            tPrint = "DISC;" & SP_FMTcCurPoint(oDisc.cLastDiscount, False)
                            If bVB_OPOSLineDisplayEnable Then
                                Call SP_PrnPollDisplay(tPrint)
                            End If
                             Call W_SHWxDiscountByTrans(oDisc.cLastDiscount)
                        End If
                    End If
                Else
                    Call SP_MSGnShowing("tms_CN005184", nCS_Warning)
                    otmForm.Enabled = True
                    otmLock.Enabled = True
                    Exit Sub
                End If
            Case "1890" '12 'Transaction markdown4
                If oDisc.DISbChkAlw(nEN_Disc26TranMDown4, False) Then
                    wCNTMkdNo.Show vbModal
                    If tVB_CNTranMDownNo <> "" Then
                        wCNTDisc2.Show vbModal
                        If oDisc.cLastDiscount >= 0 And bVB_DiscAccept Then
                            tPrint = "DISC;" & SP_FMTcCurPoint(oDisc.cLastDiscount, False)
                            If bVB_OPOSLineDisplayEnable Then
                                Call SP_PrnPollDisplay(tPrint)
                            End If
                             Call W_SHWxDiscountByTrans(oDisc.cLastDiscount)
                        End If
                    End If
                Else
                    Call SP_MSGnShowing("tms_CN005184", nCS_Warning)
                    otmForm.Enabled = True
                    otmLock.Enabled = True
                    Exit Sub
                End If
            Case "1889"  '13 'Transaction markdown5
                If oDisc.DISbChkAlw(nEN_Disc27TranMDown5, False) Then
                    wCNTMkdNo.Show vbModal
                    If tVB_CNTranMDownNo <> "" Then
                        wCNTDisc2.Show vbModal
                        If oDisc.cLastDiscount >= 0 And bVB_DiscAccept Then
                            tPrint = "DISC;" & SP_FMTcCurPoint(oDisc.cLastDiscount, False)
                            If bVB_OPOSLineDisplayEnable Then
                                Call SP_PrnPollDisplay(tPrint)
                            End If
                             Call W_SHWxDiscountByTrans(oDisc.cLastDiscount)
                        End If
                    End If
                Else
                    Call SP_MSGnShowing("tms_CN005184", nCS_Warning)
                    otmForm.Enabled = True
                    otmLock.Enabled = True
                    Exit Sub
                End If
        End Select
    End With
    If Not otmForm.Enabled Then otmForm.Enabled = True
    If Not otmLock.Enabled Then otmLock.Enabled = True
    Exit Sub
ErrHandle:
    If Not otmForm.Enabled Then otmForm.Enabled = True
    If Not otmLock.Enabled Then otmLock.Enabled = True
    Call SP_CRTxLogPrg("ErrorNumber:" & Err.Number & " ErrorDesc:" & Err.Description & " Form:" & Me.Name & " Control:ocmOK" & " Events:ocmOK_Click")
End Sub

Private Sub ogdForm_KeyDown(KeyCode As Integer, Shift As Integer)       '*Ao 51-06-26
    Select Case KeyCode
        Case vbKeyHome
            Call Form_KeyDown(vbKeyHome, 0)
        Case vbKeyDown
            If ogdFunc.Row < ogdFunc.rows - 1 Then
                ogdFunc.Select ogdFunc.Row + 1, 0, ogdFunc.Row + 1, 0
            End If
        Case vbKeyUp
            If ogdFunc.Row > 0 Then
                ogdFunc.Select ogdFunc.Row - 1, 0, ogdFunc.Row - 1, 0
            End If
    End Select
End Sub

Private Sub ogdFunc_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    'tW_HotKey = ""
End Sub

Private Sub ogdFunc_DblClick()
    Call ocmOk_Click
End Sub

Private Sub ogdFunc_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim tWhe$
'*KT 51-11-17
'   Debug.Print "9"
    Call W_CHKxDiscount91(KeyCode, Shift)    '*Eaw 56-01-05
    '----------------------------------------------------
    If bVB_CNRecovryAct Then
        KeyCode = 0
        Exit Sub
    End If
    '--------------------------------------------------
    Select Case KeyCode
        Case vbKeyHome
'            tWhe = "  WHERE FTTmnNum=''" & tVB_CNTerminalNum & "''"
'            tWhe = tWhe & "  AND FTShdTransNo=''" & tVB_Running & "''"
'            tWhe = tWhe & "  AND FDShdTransDate =''" & tUT_SaleDate & "''"
'            tWhe = tWhe & "  AND ISNULL(FTSdtStkCode,'') <> '1'" '*KT 52-11-25 Deal&Promo
'            cVB_TotalDisc = SP_FMTcCurPoint(SP_VALcText2Double(SP_SQLtFunctionSTP("SUM", "FCScdAmt", tVB_CNTblCD, tWhe)), True, False)
'            bVB_CNTrnDisc = True
'            Select Case Val(Format(tVB_TransactionSub))
'                Case tEN_TrnDeposit 'Deposit Create Transaction
'                    '***KT 51-11-20
'                    tWhe = " WHERE  FTTmnNum=''" & tVB_CNTerminalNum & "'' AND FTShdTransNo=''" & tVB_Running & "''"
'                    tWhe = tWhe & " AND  FTShdTransType=''" & tVB_Transaction & "'' AND FDShdTransDate=''" & tUT_SaleDate & "''"
'                    tWhe = tWhe & " AND  ISNULL(FTSdtCancelFlag,''0'') <> ''1''"
'                    'cUT_SubTotal = SP_FMTcCurPoint(SP_VALcText2Double(SP_SQLtFunction("SUM", "FCSdtSaleAmt", tVB_CNTblDT, tWhe)), True, False)
'                    cUT_SubTotal = SP_FMTcCurPoint(SP_VALcText2Double(SP_SQLtFunctionSTP("SUM", "FCSdtSaleAmt", tVB_CNTblDT, tWhe)), True, False)
'                    wCNDepositAmt.Show vbModal
'                    If Not bVB_TrnDeposit Then
'                        Exit Sub
'                    End If
'            End Select
'            cVB_TotalDiscAft = SP_VALcText2Double(ogdForm.Aggregate(flexSTSum, ogdForm.FixedRows, 4, ogdForm.rows - 1, 4))
'            cVB_TotalDiscMallCrd = W_SUMcDiscByType(nEN_Disc11MallCrd)
'            Unload Me

            nVB_OnTopCount = 0
'            Debug.Print "7"
            ReDim aVB_CreditCrd(0) As String
            aVB_CreditCrd(0) = ""
            tWhe = "  WHERE FTTmnNum=''" & tVB_CNTerminalNum & "''"
            tWhe = tWhe & "  AND FTShdTransNo=''" & tVB_Running & "''"
            tWhe = tWhe & "  AND FDShdTransDate =''" & tUT_SaleDate & "''"
            tWhe = tWhe & "  AND FTShdTransType=''" & tVB_Transaction & "''"
            tWhe = tWhe & "  AND ISNULL(FTSdtStkCode,'''') <> ''1''" '*KT 52-11-25 Deal&Promo
            cVB_TotalDisc = SP_FMTcCurPoint(SP_VALcText2Double(SP_SQLtFunctionSTP("SUM", "FCScdAmt", tVB_CNTblCD, tWhe)), True, False)
            bVB_CNTrnDisc = True
            Select Case Val(Format(tVB_TransactionSub))
                Case tEN_TrnDeposit 'Deposit Create Transaction
                    tWhe = " WHERE  FTTmnNum=''" & tVB_CNTerminalNum & "'' AND FTShdTransNo=''" & tVB_Running & "''"
                    tWhe = tWhe & " AND  FTShdTransType=''" & tVB_Transaction & "'' AND FDShdTransDate=''" & tUT_SaleDate & "''"
                    tWhe = tWhe & " AND  FTSdtCancelFlag <> ''1''"
                    cUT_SubTotal = SP_FMTcCurPoint(SP_VALcText2Double(SP_SQLtFunctionSTP("SUM", "FCSdtSaleAmt", tVB_CNTblDT, tWhe)), True, False)
                    Call W_DEFxCheckVendor                      '*Ao 53-10-19
                    ''*Eaw 56-02-06 การ Recovery ข้อมูล
                    If Not bVB_CNPaid Then
                        wCNDepositAmt.Show vbModal
                        If Not bVB_TrnDeposit Then
                            Exit Sub
                        End If
                    End If
'                    wCNDepositAmt.Show vbModal
'                    If Not bVB_TrnDeposit Then
'                        Exit Sub
'                    End If
'---------------------------------------------------------------------------------------
            End Select
            cVB_TotalDiscAft = SP_VALcText2Double(ogdForm.Aggregate(flexSTSum, ogdForm.FixedRows, 4, ogdForm.rows - 1, 4))
            cVB_TotalDiscMallCrd = W_SUMcDiscByType(nEN_Disc11MallCrd)
            cVB_TotalDiscEmpCrd = W_SUMcDiscByType(nEN_Disc12EmpCrd)
            Unload Me
    End Select
End Sub

Private Sub W_PRCxCallFunc(ByVal pnFuc As Long)
    Dim bFound As Boolean
    Dim nI As Long
On Error GoTo ErrHandle
    With ogdFunc
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
'    Dim nMaxRow As Long
'    nMaxRow = ogdFunc.Aggregate(flexSTMax, 0, 0, ogdFunc.rows - 1, 0)
    If IsNumeric(tW_HotKey) Then tW_HotKey = "": otmForm.Enabled = False
End Sub


Public Sub W_PRCxAutoTransDisc()
'-------------------------------------------------------
'Call:
'Cmt: ทำการคำนวนส่วนลดที่เกิดจาก Auto Transaction
'-------------------------------------------------------
    Dim tPrint$
On Error GoTo ErrHandle
    Call oDisc.DISxClearAutoTran '*TW 51-05-23
    '*KT 54-11-23 V.4.0.5 ให้ตรวจสอบว่ามีส่วนลด AutoTrans ที่ Active หรือไม่ก่อนเข้าสู่การคำนวน เพื่อลดเวลา
    If oDisc.DISbChkAutoTransAct Then
        oDisc.bRecalByItem = False
        oDisc.DISxClearAll
        oDisc.dTranDate = tUT_SaleDate
        oDisc.tTerminalNo = tVB_CNTerminalNum
        oDisc.tTransNo = tVB_Running
        Call oDisc.DIScDiscount(nEN_Disc07AutoTran)
        If oDisc.cLastDiscount > 0 Then
            If bVB_OPOSLineDisplayEnable Then
                tPrint = "DISC;" & SP_FMTcCurPoint(oDisc.cLastDiscount, False)
                Call SP_PrnPollDisplay(tPrint)
            End If
            Call W_SHWxDiscountByTrans(oDisc.cLastDiscount)
        End If
    End If
    bVB_CNCalAutoTran = True '*KT 51-05-29 ML-SF-275
    Exit Sub
ErrHandle:
    
End Sub

Private Sub W_SHWxDiscountByTrans(ByVal pcAmt As Double, Optional ByVal pnDiscType As Long = 99)
'-------------------------------------------------------
'Call:
'Cmt: แสดงรายการส่วนลดที่เกิดขึ้นในหน้าจอนี้
'-------------------------------------------------------
    Dim tSql$, tWhr$
    Dim orsTemp As ADODB.Recordset
    Dim nRow As Long
    Dim nSeqNo As Long
On Error GoTo ErrHandle
    With ogdForm
        'Clear grid
        .rows = .FixedRows
        .rows = 4
        .TextMatrix(1, 0) = "1"
        .TextMatrix(2, 0) = "2"
        .TextMatrix(3, 0) = "3"
        '----------------------------------------
        'Show data
        tWhr = "  WHERE FTTmnNum=''" & tVB_CNTerminalNum & "''"
        tWhr = tWhr & "  AND FTShdTransNo=''" & tVB_Running & "''"
        tWhr = tWhr & "  AND FTShdTransType=''" & tVB_Transaction & "''"
        tWhr = tWhr & "  AND FDShdTransDate =''" & tUT_SaleDate & "''"
        tWhr = tWhr & "  AND FNDctNo IN (7,8,9,10,11,12,13,14,23,24,25,26,27)"  '*KT 54-06-29 V.4.0.0 Req.18 Add 5 Types of TransMarkdown
        tWhr = tWhr & "  GROUP BY FNDctNo,FTDctName,FCScdDiscPerOrg,FCScdDiscAmtOrg,FNScdActSeqNo"
        tWhr = tWhr & "  ORDER BY FNScdActSeqNo ASC"
        tSql = "STP_CNxSQLFunction @ptFunc = 'FNDctNo,FTDctName,Sum(FCScdAmt) AS FCScdAmt,FCScdDiscPerOrg,FCScdDiscAmtOrg' ,@ptTable = '" & tVB_CNTblCD & "',@ptWhere = '" & tWhr & "'"
        
        Call SP_SQLvExecute(tSql, orsTemp)
        nRow = .FixedRows
        nSeqNo = 1
        Do While Not orsTemp.EOF
             If nRow >= .rows - 1 Then .rows = .rows + 1
            .TextMatrix(nRow, 0) = nSeqNo 'Seq
            .TextMatrix(nRow, 1) = SP_FEDtChkString(orsTemp, "FTDctName") 'Name
            .TextMatrix(nRow, 2) = SP_FMTcCurPoint(SP_FEDvChkNumeric(orsTemp, "FCScdDiscAmtOrg"), False, True)
            .TextMatrix(nRow, 3) = SP_FMTcCurPoint(SP_FEDvChkNumeric(orsTemp, "FCScdDiscPerOrg"), False, True)
            .TextMatrix(nRow, 4) = SP_FMTcCurPoint(SP_FEDvChkNumeric(orsTemp, "FCScdAmt"), False, True)
            .TextMatrix(nRow, 5) = SP_FEDvChkNumeric(orsTemp, "FNDctNo")
            nSeqNo = nSeqNo + 1
            nRow = nRow + 1
            orsTemp.MoveNext
        Loop
        Call SP_DATxRsNothing(orsTemp)
    End With
    'Refresh balance
    cVB_BalAftDisc = cVB_BalAftDisc - pcAmt
    olaBal.Caption = SP_FMTcCurPoint(cVB_BalAftDisc, False, True)
    '--------------------------------------------
    Call oDisc.DISxGetSubRegPro(cUT_SubReg, cUT_SubPro)
'    cUT_SubTotal = cUT_SubReg + cUT_SubPro
'*Eaw 56-07-31CommSheet TK-ISS3000-283 TR FAPL-063.xls แยกคิด ManualTax ออกจากค่าอื่น
    Select Case Val(Format(tVB_TransactionSub))
        Case tEN_TrnSaleManualTax, tEN_TrnReturnManualTax
'            cUT_SubReg = cUT_SubReg + cVB_GSTVatAmt
'            cUT_SubTotal = cUT_SubReg + cUT_SubPro
            '*Eaw 56-08-14
            tWhr = "  WHERE FTTmnNum=''" & tVB_CNTerminalNum & "''"
            tWhr = tWhr & "  AND FTShdTransNo=''" & tVB_Running & "''"
            tWhr = tWhr & "  AND FTShdTransType=''" & tVB_Transaction & "''"
            tWhr = tWhr & "  AND FDShdTransDate =''" & tUT_SaleDate & "''"
            tWhr = tWhr & "  AND FTSdtCancelFlag <> ''1''" '*Eaw 56-08-30 CommSheet TK-ISS3000-334-UAT-FAPL-109.xlsx
            tSql = "STP_CNxSQLFunction @ptFunc = 'Sum(FCSdtTotalBFVat) AS FCSdtTotalBFVat' ,@ptTable = '" & tVB_CNTblDT & "',@ptWhere = '" & tWhr & "'"
            Call SP_SQLvExecute(tSql, orsTemp)
            If Not (orsTemp.BOF And orsTemp.EOF) Then
                cUT_SubReg = SP_FEDtChkString(orsTemp, "FCSdtTotalBFvat")
                cUT_SubPro = 0
                cUT_SubTotal = cUT_SubReg + cUT_SubPro + cVB_GSTVatAmt
            End If
            Call SP_DATxRsNothing(orsTemp)
        Case Else
            cUT_SubTotal = cUT_SubReg + cUT_SubPro
    End Select

'*KT 51-09-16 ไม่แสดงให้แสดงค่าเดิม
'    olaTSubReg.Caption = SP_FMTcCurPoint(cUT_SubReg, False, True)
'    olaTSubPro.Caption = SP_FMTcCurPoint(cUT_SubPro, False, True)
'    olaTSubTotal.Caption = SP_FMTcCurPoint(cUT_SubTotal, False, True)
    Exit Sub
ErrHandle:
    
End Sub

Private Sub ocmHome_Click()
On Error GoTo ErrHandle
    Call Form_KeyDown(vbKeyHome, 0)
    Exit Sub
ErrHandle:
    If Not otmForm.Enabled Then otmForm.Enabled = True
    If Not otmLock.Enabled Then otmLock.Enabled = True
    Call SP_CRTxLogPrg("ErrorNumber:" & Err.Number & " ErrorDesc:" & Err.Description & " Form:" & Me.Name & " Control:ocmHome" & " Events:ocmHome_Click")
End Sub

Private Function W_PRCbChkItemDiscCpCpnAlw(ByVal pnDis As eEN_DiscType) As Boolean
'---------------------------------------------
'Call:
'Degit 1 : Disc Line Item,Degit 2 : Disc Trans,Degit 3 : Disc Line Mkd,Degit 4 : Disc Mkd Trans,Degit 5 : Coop Disc,
'Degit 6 : Coop Cpn,Degit 7 : Coop Brd,Degit 8 : Mall Cpn,Degit 9 : Mall/Emp ,Degit 10 : Auto Trans,Degit 11 : On Top,Degit 12 : Free Items
'Cmt: หาว่ารายการสินค้าที่ขายอยุ่มีสินค้าที่อนุญาตให้สามารถทำส่วนลด CoCpn/CoBnd
'--------------------------------------------
    Dim nI As Long
On Error GoTo ErrHandle
    W_PRCbChkItemDiscCpCpnAlw = False
    With wMain.ogdForm
        For nI = .FixedRows To .rows - 1
            If .TextMatrix(nI, .ColIndex("FNSdtSeqNo")) <> "" And .TextMatrix(nI, .ColIndex("FTSdtCancelFlag")) = "0" Then
                Select Case pnDis
                    Case nEN_Disc09CoopCpn
                        If Mid(.TextMatrix(nI, .ColIndex("FTSdtStaPdt")), 6, 1) = "1" Then W_PRCbChkItemDiscCpCpnAlw = True: Exit For
                    Case nEN_Disc10CoopBrand
                        If Mid(.TextMatrix(nI, .ColIndex("FTSdtStaPdt")), 7, 1) = "1" Then W_PRCbChkItemDiscCpCpnAlw = True: Exit For
                End Select
            End If
        Next nI
    End With
    Exit Function
ErrHandle:
    
End Function

Private Function W_PRCbChkItemDiscCpCpnAlwEx(ByVal pnDis As eEN_DiscType) As Boolean
'---------------------------------------------
'Call:
'Cmt: หาว่ารายการสินค้าที่ขายอยุ่มีสินค้าที่อนุญาตให้สามารถทำส่วนลด CoCpn/CoBnd
'--------------------------------------------
    Dim nI As Long
    Dim tSql$
    tSql = ""
    W_PRCbChkItemDiscCpCpnAlwEx = False

End Function

Private Function W_SUMcDiscByType(pnDiscType As eEN_DiscType) As Double
'---------------------------------------------
'Call:
'Cmt: หาผลรวมส่วนลดตามประเภทส่วนลดที่ส่งมา
'--------------------------------------------
    Dim nI As Long
On Error GoTo ErrHandle
    W_SUMcDiscByType = 0
    With ogdForm
        For nI = .FixedRows To .rows - 1
            If CLng(Val(Format(.TextMatrix(nI, 5)))) = pnDiscType Then
                W_SUMcDiscByType = W_SUMcDiscByType + Val(Format(.TextMatrix(nI, 4)))
            End If
        Next nI
    End With
    Exit Function
ErrHandle:
    
End Function

Private Sub otmLock_Timer()
    '*KT 51-10-31
    Dim bOnOff As Boolean
    Dim bOnOffCur  As Boolean
    Dim orsTemp As ADODB.Recordset
    Dim tSql As String
    Dim tErr$
    Dim tSta$
    Dim tStaOk$
On Error GoTo ErrHandle '*KT 52-01-09 ใส่ on error เพื่อดูว่าเกิด Error Number 91 ซึ่งส่วนใหญ่เกิดจากการลง SQLEXPRESS 2005 ไม่สมบูรณ์เมื่อมาทำงานบน Offline
    '*KT 65-08-07 0 Block timer interup process
    Select Case UCase(Screen.ActiveForm.Name)
        Case UCase("wCNSubDisc640")
            'none
        Case Else
            otmLock.Enabled = False
            Exit Sub
    End Select
     '*KT 51-02-27 ทำการ Update Ststusbar หน้าจอหลัก
    bOnOff = IIf(wMain.olaOnOff.Caption = "On-line", True, False)
    bOnOffCur = IIf(oDB.SQLbAvailableServer(oDB.OnServerName, oDB.OnComIP), True, False) ' oDB.IsCurOnline
    If bOnOffCur Then
        wMain.oimOnOff.Picture = wMain.oilOnOff.ListImages(1).Picture
        wMain.olaOnOff.Caption = "On-line"
        'move file ej offline to path online
        If Not bOnOff And bOnOffCur Then
            Call SaveSetting("AdaPosMall", "POSFront", "SentSale", "1")
            Call SP_MOVxEJToOnline
            Call SP_CHKbOnlineEnable 'PH 2.0.0
            Call SP_TMNxUpdateOpenOnline   'PH 2.0.0
        End If
    Else
        wMain.oimOnOff.Picture = wMain.oilOnOff.ListImages(2).Picture
        wMain.olaOnOff.Caption = "Off-line"
    End If
    tSta = GetSetting("AdaPosMall", "POSFront", "RepSta")
     If tSta = "0" Then
        wMain.ostForm.Panels(9).Picture = Nothing
        wMain.ostForm.Panels(9).Text = "REP"
     Else
        wMain.ostForm.Panels(9).Text = ""
        tStaOk = GetSetting("AdaPosMall", "POSFront", "RepOk")
        If tStaOk = "0" Then
             wMain.ostForm.Panels(9).Picture = wMain.oilOnOff.ListImages(5).Picture
        Else
             wMain.ostForm.Panels(9).Picture = wMain.oilOnOff.ListImages(4).Picture
        End If
     End If
    
'    If bOnOffCur And oTerminal.SaleMode = nEN_OnlineNormal Then
'         wMain.ostForm.Panels(1).Text = "SRV"
'     Else
'         wMain.ostForm.Panels(1).Text = "LOCAL"
'    End If

    tSta = GetSetting("AdaPosMall", "POSFront", "SentSale")
    If tSta = "1" Then
        wMain.ostForm.Panels(10).Picture = wMain.oilOnOff.ListImages(7).Picture
    Else
        wMain.ostForm.Panels(10).Picture = wMain.oilOnOff.ListImages(6).Picture
    End If
'    DoEvents '*Ao 53-01-14   DoEvents
    
    If Not bVB_CNSysDateLocal And bOnOffCur Then
        dVB_CurrentDate = SP_GETdCurrentDate
        'If ocnAll.ConnectionString <> oDB.ocnOffLine.ConnectionString Then 'PH 2.0.0 OLD
        If oDB.SQLbAvailableServer(oDB.OffServerName, oDB.OnComIP) Then
             Call SP_SYSxChangeSystemDate    '*KT 51-08-08  Communication Sheet ML-SF-434 Date On Terminal.xls เปลี่ยน System date ของเครื่อง POS ให้ตรงกันกับ Server
             bVB_CNSysDateLocal = True
         End If
    End If
    
    '*KT 51-07-24 Unlock
'    If bOnOffCur And bOnOff = False And oDB.IsReExecute = False Then        '*Ao 2013/10/18   V4.0.84
'        Call SP_TMNxUpdateEmpOnline
'        If SP_CHKbUnLock Then
'            Call SP_MSGnShowing("tms_CN005166", nCS_Error)
'            Unload Me
'            Call SP_PRCxCancelBillOnUnLock
'        End If
'    End If
    
    'Verify Printer by OPOS
    If bVB_CNPrnOPOS Then Call SP_SHWxErrPrinter
    If bVB_CNPrnOPOS And (tVB_CNTypePrinter = "0" Or tVB_CNTypePrinter = "2") Then Call SP_SHWxErrPrinter(False, tVB_CNTypePrinter) '*Eaw 56-07-23 CommSheet TK-ISS3000-274 TR FAPL-053 00.00.02.xls
    Call SP_DATxRsNothing(orsTemp)

    Exit Sub
ErrHandle:
    Call SP_DATxRsNothing(orsTemp)
     If Err.Number = -2147467259 Or Err.Number = -2147217871 Or Err.Number = 91 Or Err.Number = 3704 Then '*KT 54-04-07 ป้องกันโปรแกรมหลุด 3.0.56
        Call SP_MSGnShowing("tms_CN005199", nCS_Warning)
        Call SP_CRTxLogPrg("ErrorNumber:" & Err.Number & " ErrorDesc:" & Err.Description & " Form:" & Me.Name & " Control:otmLock" & " Events:Timer")
        Call SP_APPxEndProgram
    Else
        Call SP_MSGxShowError(Err.Number, Err.Description)
    End If
End Sub

Private Sub W_SetBackColor(ByVal FlexGrid As VSFlexGrid, ByVal RowNo As Integer, ByVal BkColor As ColorConstants)
    With FlexGrid
        If RowNo >= .rows Then
            MsgBox "Invalid RowNo"
        End If
        .FillStyle = flexFillRepeat '<----Without this line it will not work

        .Row = RowNo
        .Col = .FixedCols

        .RowSel = RowNo
        .ColSel = .Cols - 1

        '.CellBackColor = BkColor
        .CellForeColor = BkColor
    End With
End Sub

Public Function W_PRCxRecDiscount()
'---------------------------------------------
'Call:
'Cmt: '*Eaw 56-01-07  กู้ข้อมูลส่วนลดจากตาราง TPSCDRXXXXTmp
'--------------------------------------------
Dim tSql, tWhr$
Dim orsTemp As ADODB.Recordset
Dim cDctAmt As Double
Dim tUpd$
Dim tName As String '*Eaw 56-11-01
On Error GoTo ErrHandle
    tName = SP_STRtSiggleCode(tVB_CNUserName)  '*Eaw 56-11-01ให้ชื่อมีเขาเดี่ยวได้
    '*Eaw 56-04-10  CommSheet TK-ISS3000-129-IT2 00.00.01.xls
    tUpd = " SET FDDateUpd = '" & UT_tStringDate(Date) & "',"
    tUpd = tUpd & vbCrLf & " FTTimeUpd = '" & Time$ & "',"
    tUpd = tUpd & vbCrLf & " FTWhoUpd = '" & tName & "'," '*Eaw 56-11-01ให้ชื่อมีเขาเดี่ยวได้
    tUpd = tUpd & vbCrLf & " FDDateIns = '" & UT_tStringDate(Date) & "',"
    tUpd = tUpd & vbCrLf & " FTTimeIns = '" & Time$ & "',"
    tUpd = tUpd & vbCrLf & " FTWhoIns = '" & tName & "'" '*Eaw 56-11-01ให้ชื่อมีเขาเดี่ยวได้
    '----------------------------------------------------------------------------------------------------------
    bVB_CNDiscount = False  '*Eaw 56-07-02
    tWhr = "  WHERE FTTmnNum='" & tVB_CNTerminalNum & "'"
    tWhr = tWhr & "  AND FTShdTransNo='" & tVB_Running & "'"
    tWhr = tWhr & "  AND FTShdTransType='" & tVB_Transaction & "'"
    tWhr = tWhr & "  AND FDShdTransDate ='" & tUT_SaleDate & "'"
    tSql = "SELECT SUM(FCScdAmt)  AS FCScdAmt FROM " & tVB_CNTblCDRTmp
    tSql = tSql & tWhr
    tSql = tSql & "  AND FNDctNo IN (7,8,9,10,11,12,13,14,23,24,25,26,27)"
    Call SP_SQLvExecute(tSql, orsTemp, oDB.ocnOffLine)
    
    orsTemp.MoveFirst
    cDctAmt = SP_FMTcCurPoint(SP_FEDvChkNumeric(orsTemp, "FCScdAmt"), False, True)
'*Eaw 56-07-02 Test DTTmp
    tSql = "DELETE FROM " & tVB_CNTblDT
    tSql = tSql & tWhr
    Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
    tSql = "INSERT INTO  " & tVB_CNTblDT & " SELECT * FROM " & tVB_CNTblDTRTmp
    tSql = tSql & tWhr
    Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
    
    tSql = "INSERT INTO  " & tVB_CNTblCDTmp & " SELECT * FROM " & tVB_CNTblCDRTmp
    tSql = tSql & tWhr
    Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)

    tSql = "DELETE FROM " & tVB_CNTblCDTmp
    tSql = tSql & tWhr
    Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
    tSql = "INSERT INTO  " & tVB_CNTblCD & " SELECT * FROM " & tVB_CNTblCDRTmp
    tSql = tSql & tWhr
    Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)

    olaTSubReg.Caption = SP_FMTcCurPoint(cUT_SubReg)
    olaTSubPro.Caption = SP_FMTcCurPoint(cUT_SubPro)
    olaTSubTotal.Caption = SP_FMTcCurPoint(cUT_SubTotal)
    
    Call W_SHWxDiscountByTrans(cDctAmt)

    tSql = "DROP TABLE  " & tVB_CNTblCDRTmp
    Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
    
    tSql = "DROP TABLE  " & tVB_CNTblCDR
    Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
    
     tSql = "DROP TABLE  " & tVB_CNTblDTRTmp
    Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
    
     '*Eaw 56-04-10  CommSheet TK-ISS3000-129-IT2 00.00.01.xls
    tSql = "UPDATE " & tVB_CNTblDT & tUpd & tWhr
    Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
    tSql = "UPDATE " & tVB_CNTblDTTmp & tUpd & tWhr
    Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
    tSql = "UPDATE " & tVB_CNTblCD & tUpd & tWhr
    Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
    tSql = "UPDATE " & tVB_CNTblCDTmp & tUpd & tWhr
    Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
    '--------------------------------------------------------------------------------
    Call SP_DATxRsNothing(orsTemp)
'    bVB_IsRecovery = False '*Eaw 56-07-02 CommSheet TK-ISS3000-249-TR.xls ไม่ต้อง Recovery RC
    Exit Function
ErrHandle:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Function

Public Function W_PRCxClearCDRTmp()
'---------------------------------------------
'Call:
'Cmt: '*Eaw 56-01-07  ลบข้อมูลจากตาราง TPSCDRXXXXTmp กรณีมีการกด Esc
'--------------------------------------------
Dim tSql$, tWhr
On Error GoTo ErrHandle
    tWhr = " WHERE FTTmnNum='" & tVB_CNTerminalNum & "'"
    tWhr = tWhr & "  AND FTShdTransNo='" & tVB_Running & "'"
    tWhr = tWhr & "  AND FTShdTransType='" & tVB_Transaction & "'"
    tWhr = tWhr & "  AND FDShdTransDate ='" & tUT_SaleDate & "'"
    tWhr = tWhr & "  AND FNDctNo IN (7,8,9,10,11,12,13,14,23,24,25,26,27)"
    tSql = "DELETE  FROM " & tVB_CNTblCDTmp
    tSql = tSql & tWhr
    Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
    
    tSql = "DELETE  FROM " & tVB_CNTblCD
    tSql = tSql & tWhr
    Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
    
    If SP_TBLbChkExist(tVB_CNTblCDR, oDB.ocnOffLine) Then
        tSql = "DROP TABLE  " & tVB_CNTblCDR
        Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
    End If
    If SP_TBLbChkExist(tVB_CNTblCDRTmp, oDB.ocnOffLine) Then
        tSql = "DROP TABLE  " & tVB_CNTblCDRTmp
        Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
    End If
'*Eaw 56-07-02 CommSheet TK-ISS3000-249-TR.xls ไม่ต้อง Recovery RC
'    If SP_TBLbChkExist(tVB_CNTblRCR, oDB.ocnOffLine) Then
'        tSql = "DROP TABLE  " & tVB_CNTblRCR
'        Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
'    End If
    ogdForm.Clear
    bVB_CNDiscount = False
    Exit Function
ErrHandle:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Function

Public Function W_INSxTblCD()
'---------------------------------------------
'Call:
'Cmt: '*Eaw 56-01-07  เพิ่มข้อมูลลง TPSCDXXXX กรณีมีส่วนลดรายตัว แต่ไม่มีส่วนลดท้ายบิล
'--------------------------------------------
Dim tSql$, tWhr, tUpd
Dim tName As String '*Eaw 56-11-01
On Error GoTo ErrHandle
    tName = SP_STRtSiggleCode(tVB_CNUserName)  '*Eaw 56-11-01ให้ชื่อมีเขาเดี่ยวได้
    '*Eaw 56-04-10  CommSheet TK-ISS3000-129-IT2 00.00.01.xls
    tUpd = " SET FDDateUpd = '" & UT_tStringDate(Date) & "',"
    tUpd = tUpd & vbCrLf & " FTTimeUpd = '" & Time$ & "',"
    tUpd = tUpd & vbCrLf & " FTWhoUpd = '" & tName & "'," '*Eaw 56-11-01ให้ชื่อมีเขาเดี่ยวได้
    tUpd = tUpd & vbCrLf & " FDDateIns = '" & UT_tStringDate(Date) & "',"
    tUpd = tUpd & vbCrLf & " FTTimeIns = '" & Time$ & "',"
    tUpd = tUpd & vbCrLf & " FTWhoIns = '" & tName & "'" '*Eaw 56-11-01ให้ชื่อมีเขาเดี่ยวได้
    '----------------------------------------------------------------------------------------------------------
    bVB_CNDiscount = False
    tWhr = " WHERE FTTmnNum='" & tVB_CNTerminalNum & "'"
    tWhr = tWhr & "  AND FTShdTransNo='" & tVB_Running & "'"
    tWhr = tWhr & "  AND FTShdTransType='" & tVB_Transaction & "'"
    tWhr = tWhr & "  AND FDShdTransDate ='" & tUT_SaleDate & "'"
    If SP_TBLbChkExist(tVB_CNTblCDR, oDB.ocnOffLine) Then
        tSql = "INSERT INTO  " & tVB_CNTblCD & " SELECT * FROM " & tVB_CNTblCDR
        tSql = tSql & tWhr
        Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
        tSql = "DROP TABLE  " & tVB_CNTblCDR
        Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
    End If
    If SP_TBLbChkExist(tVB_CNTblCDRTmp, oDB.ocnOffLine) Then
        tSql = "INSERT INTO  " & tVB_CNTblCDTmp & " SELECT * FROM " & tVB_CNTblCDRTmp
        tSql = tSql & tWhr
        Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
        tSql = "DROP TABLE  " & tVB_CNTblCDRTmp
        Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
    End If
    '*Eaw 56-04-10  CommSheet TK-ISS3000-129-IT2 00.00.01.xls
    tSql = "UPDATE " & tVB_CNTblDT & tUpd & tWhr
    Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
    tSql = "UPDATE " & tVB_CNTblDTTmp & tUpd & tWhr
    Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
    tSql = "UPDATE " & tVB_CNTblCD & tUpd & tWhr
    Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
    tSql = "UPDATE " & tVB_CNTblCDTmp & tUpd & tWhr
    Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
    '--------------------------------------------------------------------------------
'*Eaw 56-07-02 CommSheet TK-ISS3000-249-TR.xls ไม่ต้อง Recovery RC
'    If Not bVB_CNPaid Then
'        If SP_TBLbChkExist(tVB_CNTblRCR, oDB.ocnOffLine) Then
'            tSql = "DROP TABLE  " & tVB_CNTblRCR
'            Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
'        End If
'    End If
    Exit Function
ErrHandle:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Function

Private Sub otmRec_Timer()
'-----------------------------------------------------
'  Call:
'  Cmt:  '*Ao 2014/05/07  PRJ10001TK-PH2-External Spec-Incomplete Transaction-00.02.04 FSBT
'-----------------------------------------------------
On Error GoTo ErrHandle
    otmRec.Enabled = False
    Call SleepEx(1000, 0)
    Call Form_KeyDown(vbKeyHome, 0)
Exit Sub
ErrHandle:
End Sub

