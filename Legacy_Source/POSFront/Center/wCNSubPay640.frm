VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.1#0"; "Codejock.SuiteCtrls.Unicode.v11.1.3.ocx"
Begin VB.Form wCNSubPay640 
   BorderStyle     =   0  'None
   Caption         =   "wCNSubPay640"
   ClientHeight    =   6015
   ClientLeft      =   1020
   ClientTop       =   2895
   ClientWidth     =   9585
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   9585
   ShowInTaskbar   =   0   'False
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
      Left            =   6120
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   5560
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
      Left            =   7800
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   5560
      Width           =   1455
   End
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   870
      Index           =   2
      Left            =   0
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   5455
      Width           =   9580
      _Version        =   720897
      _ExtentX        =   16907
      _ExtentY        =   1535
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
      PaintManager.Position=   2
      Begin VB.Timer otmEDC 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   1680
         Top             =   0
      End
      Begin VB.Timer otmForm 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   720
         Top             =   0
      End
      Begin VB.Timer otmLock 
         Interval        =   2000
         Left            =   1200
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
         Left            =   360
         TabIndex        =   21
         Top             =   120
         Visible         =   0   'False
         Width           =   4995
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid ogdFunc 
      Height          =   2355
      Left            =   15
      TabIndex        =   3
      Top             =   3100
      Width           =   9580
      _cx             =   16898
      _cy             =   4154
      _ConvInfo       =   1
      Appearance      =   2
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   14.25
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
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   8
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   17
      Cols            =   8
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   350
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"wCNSubPay640.frx":0000
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
      TabIndex        =   4
      Top             =   2700
      Width           =   9580
      _Version        =   720897
      _ExtentX        =   16898
      _ExtentY        =   1191
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
      PaintManager.Position=   2
      Begin VB.Label olaPayment 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Payment"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   495
         Left            =   720
         TabIndex        =   5
         Tag             =   "2;"
         Top             =   -15
         Width           =   8295
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid ogdTender 
      Height          =   1425
      Left            =   15
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1275
      Width           =   9580
      _cx             =   16898
      _cy             =   2514
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
      FocusRect       =   1
      HighLight       =   0
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   8
      GridLinesFixed  =   4
      GridLineWidth   =   1
      Rows            =   4
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   350
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"wCNSubPay640.frx":00DD
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
      Height          =   1545
      Index           =   1
      Left            =   15
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   0
      Width           =   4890
      _Version        =   720897
      _ExtentX        =   8625
      _ExtentY        =   2725
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
      PaintManager.Position=   2
      Begin VB.Label olaTotal 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "699.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   3720
         TabIndex        =   13
         Top             =   0
         Width           =   1080
      End
      Begin VB.Label olaTotalName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   345
         Left            =   120
         TabIndex        =   12
         Top             =   45
         Width           =   735
      End
      Begin VB.Label olaTotalDiscName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Discount"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   345
         Left            =   120
         TabIndex        =   11
         Top             =   465
         Width           =   2085
      End
      Begin VB.Label olaDiscTotal 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   4110
         TabIndex        =   10
         Top             =   405
         Width           =   690
      End
      Begin VB.Label olaNetSaleName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Net Sale"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   345
         Left            =   120
         TabIndex        =   9
         Top             =   855
         Width           =   1200
      End
      Begin VB.Label olaNetSale 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "699.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   3720
         TabIndex        =   8
         Top             =   780
         Width           =   1080
      End
   End
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   1545
      Index           =   3
      Left            =   4890
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   0
      Width           =   4695
      _Version        =   720897
      _ExtentX        =   8281
      _ExtentY        =   2725
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
      PaintManager.Position=   2
      Begin VB.Label olaAmtTender 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "699.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   2100
         TabIndex        =   17
         Top             =   405
         Width           =   2400
      End
      Begin VB.Label olaBalance 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "699.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   24
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   585
         Left            =   2850
         TabIndex        =   20
         Top             =   720
         Width           =   1650
      End
      Begin VB.Label olaBalanceName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Balance Due"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   345
         Index           =   4
         Left            =   120
         TabIndex        =   19
         Top             =   855
         Width           =   1755
      End
      Begin VB.Label olaDeposit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Deposit Amt"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   345
         Left            =   120
         TabIndex        =   16
         Top             =   45
         Width           =   1770
      End
      Begin VB.Label lblDepositAmt 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "699.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   2100
         TabIndex        =   15
         Top             =   0
         Width           =   2400
      End
      Begin VB.Label olaAmtTenderName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount Tendered"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   345
         Left            =   120
         TabIndex        =   18
         Top             =   465
         Width           =   2550
      End
   End
End
Attribute VB_Name = "wCNSubPay640"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*Ao 52-02-25
' ถ้ามีการแก้ Code ใน From wCNSubPay ให้แก้ใน From wCNSubPay640 ด้วย
'wCNSubPay ใช้ Resorution 800*600
'wCNSubPay640 ใช้ Resorution 640*480
'________________________________________________________
Private tW_HotKey As String
Private nMaxRow As Long

Private Sub W_PRCxCheckEMP()
'----------------------------------------------------------
'
'----------------------------------------------------------
Dim tSql As String
Dim orsTemp As ADODB.Recordset

On Error GoTo ErrHandle
    tSql = "SELECT * FROM " & tVB_CNTblCD
    tSql = tSql & vbCrLf & "WHERE FTTmnNum = '" & tVB_CNTerminalNum & "'"
    tSql = tSql & vbCrLf & "AND FTShdTransNo = '" & tVB_Running & "'"
    tSql = tSql & vbCrLf & "AND FDShdTransDate = '" & tUT_SaleDate & "'"
    tSql = tSql & vbCrLf & "AND FNDctNo = 12"
    Call SP_SQLbExecute(tSql, orsTemp)
    If Not (orsTemp.BOF And orsTemp.EOF) Then
        ocmCancel.Enabled = False
        otmEDC.Enabled = True
    End If
Exit Sub
ErrHandle:
End Sub

Private Sub W_CHKxDiscount91(ByVal pnKeyCode As Integer, ByVal pnShift As Integer)
'______________________________________
'   Call:
'   Cmt: '*Eaw 56-01-31 แก้ไขเรื่อง Hot Key ของ  Store Wide Discount
'______________________________________
Dim tSql$
Dim tProc$
Dim orsFunc As ADODB.Recordset
 On Error GoTo ErrHandle
    tSql = "SELECT FTSkbRef FROM TSysKeyboard "
    tSql = tSql & vbCrLf & " WHERE FNSkbKey='" & CStr(pnKeyCode) & "'"
    tSql = tSql & vbCrLf & " AND FNSkbShift = '" & CStr(pnShift) & "'"
    tSql = tSql & vbCrLf & " AND FTSkbwMain =  '1'"
    tSql = tSql & vbCrLf & " AND FTSkbUsage ='Y'"
'    Debug.Print pnKeyCode
'    Debug.Print pnShift
    Call SP_SQLvExecute(tSql, orsFunc, oDB.ocnOnLine, , , , , True, True) 'read master
    If Not (orsFunc.EOF And orsFunc.BOF) Then
         tProc = SP_FEDtChkString(orsFunc, "FTSkbRef")
         If tProc = "W_KBDiscount91" Then
                Call SP_MSGnShowing("tms_CN005245", nCS_Error)
                Call SP_DATxRsNothing(orsFunc)
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
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Sub

Private Function W_CHKbCheckTKAllowEmp(ByVal pnSeqNo As Double) As Boolean
'____________________________________
'   Call:
'   Cmt: '*Ao 2012/07/26        PRJ10001TK-External Specification Takashimaya Gift Voucher-00.01.00.doc
'____________________________________
Dim tSql As String
Dim orsTemp As ADODB.Recordset
Dim bChk As Boolean

On Error GoTo ErrHandle
    
    bChk = False
    
    If aVB_TenderMnu(nVB_CurTenderSeqNo).StaEmpDisc Then     'True ผ่าน
        W_CHKbCheckTKAllowEmp = bChk
        Exit Function
    End If

    tSql = "SELECT TOP 1 * FROM " & tVB_CNTblCD
    tSql = tSql & vbCrLf & "WHERE  FTTmnNum='" & tVB_CNTerminalNum & "' AND FTShdTransNo='" & tVB_Running & "'"
    tSql = tSql & vbCrLf & "AND  FTShdTransType='" & tVB_Transaction & "' AND FDShdTransDate='" & tUT_SaleDate & "'"
    tSql = tSql & vbCrLf & "AND FNDctNo = 12"           'Employee Card

    Call SP_SQLvExecute(tSql, orsTemp)
    If Not (orsTemp.BOF And orsTemp.EOF) Then
        bChk = True
    End If
    Call SP_DATxRsNothing(orsTemp)
    W_CHKbCheckTKAllowEmp = bChk
    Exit Function
ErrHandle:
    bChk = False
    W_CHKbCheckTKAllowEmp = bChk
End Function

Private Function W_DEFbReturn() As Boolean
    W_DEFbReturn = False
    If Val(Format(tVB_TransactionSub)) = tEN_TrnReturn Or Val(Format(tVB_TransactionSub)) = tEN_TrnReturnTaxInc Or Val(Format(tVB_TransactionSub)) = tEN_TrnReturnManualTax Then        '* Ao 51-06-27 การคืนแบบไม่มีเงื่อนไข ไม่ต้องมีส่วนลดใดๆเกิดขึ้น
        If bVB_Reference = False Then
            W_DEFbReturn = True
        End If
    End If
End Function

Private Sub W_FNCxAddFunc2Grid()
'--------------------------------------------------------
'Call :
'Cmt : *KT 51-01-02 ทำการแสดงรายการ Tender ทั้งหมดของระบบกรณีการขาย
'*KT 53-10-15 Commsheet หมายเลข PH3-ML-SF-131
'-------------------------------------------------------
    Dim nR%
    Dim nI As Long
On Error GoTo ErrHandle
    With ogdFunc
        nR = LBound(aVB_TenderMnu)
      .rows = UBound(aVB_TenderMnu)
        For nI = LBound(aVB_TenderMnu) To UBound(aVB_TenderMnu) - 1
'             If aVB_TransMnu(nI).Show Then
            If aVB_TenderMnu(nI).Show Then   '*Eaw 56-05-16 CommSheet TK-ISS3000-180-ST 00 00 01.xls
                Select Case Val(Format(tVB_TransactionSub))
                    Case tEN_TrnReturn, tEN_TrnCancelVoucher, tEN_TrnDepositCancel, tEN_TrnAddDepositCancel
                       '*KT 51-07-08 Communication Sheet ML-SF-137(2) Payment menu is not complete
                       ' If aVB_TenderMnu(nI).StaReturn Then
                       
                            'If aVB_TenderMnu(nI).ID <> "T003" Then   '*Eaw 55-11-26 ให้แสดงเมื่อเป็น T003 ด้วย
                                .TextMatrix(nR, 0) = nR + 1
'                                .TextMatrix(nR, 1) = nR + 1 & "." & aVB_TenderMnu(nI).Name '*KT 53-10-15 Commsheet หมายเลข PH3-ML-SF-131
                                .TextMatrix(nR, 1) = nR + 1 & "." & aVB_TenderMnu(nI).Menu
                                .TextMatrix(nR, 2) = aVB_TenderMnu(nI).SmnID
                                .TextMatrix(nR, 3) = aVB_TenderMnu(nI).ID
                                .TextMatrix(nR, 4) = aVB_TenderMnu(nI).DescReturn
                                .TextMatrix(nR, 5) = aVB_TenderMnu(nI).Name
                                .TextMatrix(nR, 6) = IIf(aVB_TenderMnu(nI).StaReturn, "1", "0")
                                .TextMatrix(nR, 7) = IIf(aVB_TenderMnu(nI).StaReEntry, "1", "0")
'                                '*Tao 60-09-01 ของเดิม
'                                If .TextMatrix(nR, 6) = "0" Then '55-07-20 Neung TK
'                                    Call W_SetBackColor(ogdFunc, nR, 13619151)
'                                End If

                                '*Tao 60-09-01 ถ้าเป็น WalletPay ไม่ต้องเปลี่ยนสี
                                Select Case aVB_TenderMnu(nI).ID
                                Case "T019", "T020", "T021", "T022"
                                    If Not bVB_AlwWalletPay Then Call W_SetBackColor(ogdFunc, nR, 13619151)
                                Case Else
                                    If .TextMatrix(nR, 6) = "0" Then '55-07-20 Neung TK
                                        Call W_SetBackColor(ogdFunc, nR, 13619151)
                                    End If
                                End Select
                                
                                nR = nR + 1
                            'End If   '*Eaw 55-11-26 ให้แสดงเมื่อเป็น T003 ด้วย
                       ' End If
                    Case tEN_TrnReturnTaxInc, tEN_TrnReturnManualTax
                           ' If aVB_TenderMnu(nI).ID <> "T003" Then   '*Eaw 55-11-26 ให้แสดงเมื่อเป็น T003 ด้วย
                                .TextMatrix(nR, 0) = nR + 1
                                .TextMatrix(nR, 1) = nR + 1 & "." & aVB_TenderMnu(nI).Menu
                                .TextMatrix(nR, 2) = aVB_TenderMnu(nI).SmnID
                                .TextMatrix(nR, 3) = aVB_TenderMnu(nI).ID
                                .TextMatrix(nR, 4) = aVB_TenderMnu(nI).DescReturn
                                .TextMatrix(nR, 5) = aVB_TenderMnu(nI).Name
                                .TextMatrix(nR, 6) = IIf(aVB_TenderMnu(nI).StaReturn, "1", "0")
                                .TextMatrix(nR, 7) = IIf(aVB_TenderMnu(nI).StaReEntry, "1", "0")
'                                '*Tao 60-09-01 ของเดิม
'                                If aVB_TenderMnu(nI).StaReEntry = False Then '55-07-20 Neung TK
'                                    Call W_SetBackColor(ogdFunc, nR, 13619151)
'                                End If
                                
                                '*Tao 60-09-01 ถ้าเป็น WalletPay ไม่ต้องเปลี่ยนสี
                                Select Case aVB_TenderMnu(nI).ID
                                Case "T019", "T020", "T021", "T022"
                                    If Not bVB_AlwWalletPay Then Call W_SetBackColor(ogdFunc, nR, 13619151)
                                Case Else
                                    If aVB_TenderMnu(nI).StaReEntry = False Then '55-07-20 Neung TK
                                        Call W_SetBackColor(ogdFunc, nR, 13619151)
                                    End If
                                End Select
                                
                                nR = nR + 1
                           ' End If  '*Eaw 55-11-26 ให้แสดงเมื่อเป็น T003 ด้วย
                    Case tEN_TrnSaleTaxInc, tEN_TrnSaleManualTax
                            'If aVB_TenderMnu(nI).ID <> "T003" Then   '*Eaw 55-11-26 ให้แสดงเมื่อเป็น T003 ด้วย
                                .TextMatrix(nR, 0) = nR + 1
'                                .TextMatrix(nR, 1) = nR + 1 & "." & aVB_TenderMnu(nI).Name '*KT 53-10-15 Commsheet หมายเลข PH3-ML-SF-131
                                .TextMatrix(nR, 1) = nR + 1 & "." & aVB_TenderMnu(nI).Menu
                                .TextMatrix(nR, 2) = aVB_TenderMnu(nI).SmnID
                                .TextMatrix(nR, 3) = aVB_TenderMnu(nI).ID
                                .TextMatrix(nR, 4) = aVB_TenderMnu(nI).DescSale
                                .TextMatrix(nR, 5) = aVB_TenderMnu(nI).Name
                                .TextMatrix(nR, 6) = IIf(aVB_TenderMnu(nI).StaSale, "1", "0")
                                .TextMatrix(nR, 7) = IIf(aVB_TenderMnu(nI).StaReEntry, "1", "0")
'                                '*Tao 60-09-01 ของเดิม
'                                If aVB_TenderMnu(nI).StaReEntry = False Then    '55-07-20 Neung TK
'                                    Call W_SetBackColor(ogdFunc, nR, 13619151)
'                                End If
                                
                                '*Tao 60-09-01 ถ้าเป็น WalletPay ไม่ต้องเปลี่ยนสี
                                Select Case aVB_TenderMnu(nI).ID
                                Case "T019", "T020", "T021", "T022"
                                    If Not bVB_AlwWalletPay Then Call W_SetBackColor(ogdFunc, nR, 13619151)
                                Case Else
                                    If aVB_TenderMnu(nI).StaReEntry = False Then '55-07-20 Neung TK
                                        Call W_SetBackColor(ogdFunc, nR, 13619151)
                                    End If
                                End Select
                                
                                nR = nR + 1
                          '  End If '*Eaw 55-11-26 ให้แสดงเมื่อเป็น T003 ด้วย
                    Case Else
                       '*KT 51-07-08 Communication Sheet ML-SF-137(2) Payment menu is not complete
                       ' If aVB_TenderMnu(nI).StaSale Then
                            'If aVB_TenderMnu(nI).ID <> "T003" Then  '*Eaw 55-11-26 ให้แสดงเมื่อเป็น T003 ด้วย
                                .TextMatrix(nR, 0) = nR + 1
'                                .TextMatrix(nR, 1) = nR + 1 & "." & aVB_TenderMnu(nI).Name '*KT 53-10-15 Commsheet หมายเลข PH3-ML-SF-131
                                .TextMatrix(nR, 1) = nR + 1 & "." & aVB_TenderMnu(nI).Menu
                                .TextMatrix(nR, 2) = aVB_TenderMnu(nI).SmnID
                                .TextMatrix(nR, 3) = aVB_TenderMnu(nI).ID
                                .TextMatrix(nR, 4) = aVB_TenderMnu(nI).DescSale
                                .TextMatrix(nR, 5) = aVB_TenderMnu(nI).Name
                                .TextMatrix(nR, 6) = IIf(aVB_TenderMnu(nI).StaSale, "1", "0")
                                .TextMatrix(nR, 7) = IIf(aVB_TenderMnu(nI).StaReEntry, "1", "0")
'                                '*Tao 60-09-01 ของเดิม
'                                If .TextMatrix(nR, 6) = "0" Then '55-07-20 Neung TK
'                                    Call W_SetBackColor(ogdFunc, nR, 13619151)
'                                End If
                                
                                '*Tao 60-09-01 ถ้าเป็น WalletPay ไม่ต้องเปลี่ยนสี
                                Select Case aVB_TenderMnu(nI).ID
                                Case "T019", "T020", "T021", "T022"
                                    If Not bVB_AlwWalletPay Then Call W_SetBackColor(ogdFunc, nR, 13619151)
                                Case Else
                                    If .TextMatrix(nR, 6) = "0" Then '55-07-20 Neung TK
                                        Call W_SetBackColor(ogdFunc, nR, 13619151)
                                    End If
                                End Select
                                
                                nR = nR + 1
                           ' End If '*Eaw 55-11-26 ให้แสดงเมื่อเป็น T003 ด้วย
                       ' End If
                End Select
            End If
        Next nI
        .rows = nR
    End With
    Exit Sub
ErrHandle:
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
     If bVB_CNSplash Then KeyCode = 0 '*KT 51-09-17 ป้องการ Key กรณีที่ Splash ขึ้น
     Call W_CHKxDiscount91(KeyCode, Shift)     '*Eaw 56-01-31 แก้เรื่อง Store Wide Discount
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Not bVB_CNTenderAct Then
            Call ocmOK_Click
        End If
'        DoEvents '*KT 53-01-14 DoEvents
    Else
        Select Case KeyAscii
            Case 48 To 58
                otmForm.Enabled = True
                With ogdFunc
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
'-----------------------------------------------------------------------------------------------
'*KT 52-07-24
'1. ในการแก้ไขต้องนำ Procedure  W_PRCxClearRC ขึ้นมาก่อนการทำการหาค่า Vendor Coupon
'2. ยกเลิกการ Req ที่ต้องสามารถจ่าย Tender Vendor Coupon พร้อมการ Scan สินค้า ที่ปัจจุบันมีสถานะ Hold ไว้
'3. ใน Procedure  W_PRCxClearRC  ให้ทำการ Clear ทุก Tendor
'4. ทำการ Mark Code กรณีที่ Sub ข้อมูล  Vendor
'------------------------------------------------------------------------------------------------
    Dim tPrint$
    Dim orsTemp As ADODB.Recordset
    Dim tSql$, tWhr$
    Dim nRow As Long
    Dim cVendorCpn As Double
    Dim tTdmType As String
    Dim nLoop As Long
    
On Error GoTo ErrHandle
    bVB_CheckVendor = False     '*Ao 53-10-13   CheckVendor
    Call SP_CRTxWriteLogTerminated("ErrorNumber:0" & " ErrorDesc:" & " Form:" & Me.Name & " Events:Form_Load[1]")
    tW_HotKey = ""
    Me.Top = IIf(bVB_Res640X480, cCS_Top680x480, cCS_Top800x600)
    Me.Left = IIf(bVB_Res640X480, cCS_Left680x480, cCS_Left800x600)
    If Not bVB_Res640X480 Then
        Call SP_FRMxGetCaptionFrmDB2Tag(Me)
        Call SP_FRMxSetCapForm(Me)
    End If
    Call W_FNCxAddFunc2Grid
    nMaxRow = ogdFunc.Aggregate(flexSTMax, 0, 0, ogdFunc.rows - 1, 0)
    Call W_SetGridHeading
    
   nVB_OnTopCount = 0
   ReDim aVB_CreditCrd(0) As String
   aVB_CreditCrd(0) = ""
   
   Call W_PRCxClearRC 'Clear ข้อมูล RC ยกเว้น Vendor Coupon ก่อนอันดับแรก '*KT 52-07-24   ยกเลิกการ Req ที่ต้องสามารถจ่าย Tender Vendor Coupon พร้อมการ Scan สินค้า ที่ปัจจุบันมีสถานะ Hold ไว้
    
    'Show Tender Vendor Coupon
    cVendorCpn = 0
    
'*KT 52-07-24   ยกเลิกการ Req ที่ต้องสามารถจ่าย Tender Vendor Coupon ที่ปัจจุบันมีสถานะ Hold ไว้
'*Ao 53-10-15   New Vendor
'------------------------------------------------------------------------------------------------------
    With ogdTender
        tWhr = " WHERE  FTTmnNum=''" & tVB_CNTerminalNum & "'' AND FTShdTransNo=''" & tVB_Running & "''"
        tWhr = tWhr & vbCrLf & " AND  FTShdTransType=''" & tVB_Transaction & "'' AND FDShdTransDate=''" & tUT_SaleDate & "''"
        tWhr = tWhr & vbCrLf & " AND  FTTdmCode = ''T006''"
        tWhr = tWhr & vbCrLf & " AND  ISNULL(FTSrcStaNoPay,'''') <> ''1''"      '*Ao 2011/01/17   CommSheet ML-PRV-020-Change Tender is wrong concept.xls
        tWhr = tWhr & vbCrLf & " ORDER BY  FNSrcSeqNo ASC"
        tSql = "STP_CNxSQLFunction @ptFunc = 'FNSrcSeqNo,FTTdmType,FCSrcNet' ,@ptTable = '" & tVB_CNTblRC & "',@ptWhere = '" & tWhr & "'"
        Call SP_SQLvExecute(tSql, orsTemp)
        If Not (orsTemp.BOF And orsTemp.EOF) Then
            orsTemp.MoveFirst
            tTdmType = SP_FEDtChkString(orsTemp, "FTTdmType")
            Do While Not orsTemp.EOF
                cVendorCpn = cVendorCpn + CDbl(Val(Format(SP_FEDvChkNumeric(orsTemp, "FCSrcNet"))))
                orsTemp.MoveNext
            Loop
                nRow = 1
                If nRow > .rows - 1 Then .rows = .rows + 1
                .TextMatrix(nRow, .ColIndex("FNSrcSeqNo")) = 1
                .TextMatrix(nRow, .ColIndex("FTTdmType")) = tTdmType
                .TextMatrix(nRow, .ColIndex("FCSrcNet")) = cVendorCpn
        End If
        Call SP_DATxRsNothing(orsTemp)
    End With
    '________________________________________________________________________________________________
    ''*Ao 53-10-16   ถ้าแก้ไขส่วนนี้ให้ไปแก้ไขที่ wCNShowVendor ด้วย
    Select Case Val(Format(tVB_TransactionSub))
        Case tEN_TrnDeposit, tEN_TrnAddDeposit
            'non ทำการหา cVB_RoundRcv ที่ PRCxDeposit/wCNDepositAmt แล้ว
            cVB_NetSaleRound = cVB_NetSale + cVB_NetDutyFreeDep + cVB_RoundRcv
            olaTotal.Caption = SP_FMTcCurPoint(cVB_Total, False, True)  ''*KT 51-04-10 Comsheet ML-SF-201 แสดงยอดรวมก่อนหักส่วนลดที่เกิดจาก Form wCNSubDisc2
            olaDiscTotal.Caption = SP_FMTcCurPoint(cVB_TotalDiscAft, False, True) 'SP_FMTcCurPoint(cVB_TotalDisc, False, True) '*KT 51-04-10 Comsheet ML-SF-201 แสดงส่วนลดเฉพาะที่เกิดจากหน้า wCNSubDisc2
            olaNetSale.Caption = SP_FMTcCurPoint(cVB_Total + cVB_NetDutyFreeDep - cVB_TotalDiscAft, False, True)
            olaAmtTender.Caption = SP_FMTtNum2Shw(0 + cVendorCpn, False, True)
            cVB_Balance = cVB_NetSaleRound - cVendorCpn
            olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance), False, True)
            lblDepositAmt.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_DepositAmt), False, True)
        Case tEN_TrnDepositCancel, tEN_TrnAddDepositCancel, tEN_TrnRedeem
            cVB_NetSaleRound = cVB_NetSale + cVB_NetDutyFreeDep + cVB_RoundRcv
            olaTotal.Caption = SP_FMTcCurPoint(cVB_Total, False, True)  ''*KT 51-04-10 Comsheet ML-SF-201 แสดงยอดรวมก่อนหักส่วนลดที่เกิดจาก Form wCNSubDisc2
            olaDiscTotal.Caption = SP_FMTcCurPoint(cVB_TotalDiscAft, False, True) 'SP_FMTcCurPoint(cVB_TotalDisc, False, True) '*KT 51-04-10 Comsheet ML-SF-201 แสดงส่วนลดเฉพาะที่เกิดจากหน้า wCNSubDisc2
            olaNetSale.Caption = SP_FMTcCurPoint(cVB_Total + cVB_NetDutyFreeDep - cVB_TotalDiscAft, False, True)
            olaAmtTender.Caption = SP_FMTtNum2Shw(0 + cVendorCpn, False, True)
            cVB_Balance = cVB_NetSaleRound - cVendorCpn
            olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance), False, True)
            lblDepositAmt.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_DepositAmt), False, True)
        Case Else
            'cVB_RoundDiffAmt = SP_FMTcCurPoint((cVB_NetSale + cVB_NetDutyFreeDep) - ((cVB_NetSale + cVB_NetDutyFreeDep) \ 1), True, False)
'            cVB_RoundRcv = IIf(bVB_CNAlwRnd, SP_FMTcCurPoint(UT_GetnRoundDiff(cVB_NetSale + cVB_NetDutyFreeDep), True, False), 0) '*KT 51-04-05 Comsheet ML-SF-172,ML-SF-173
            cVB_RoundRcv = 0 '*Eaw 56-01-03
            cVB_NetSaleRound = cVB_NetSale + cVB_RoundRcv
            olaTotal.Caption = SP_FMTcCurPoint(cVB_Total, False, True)  ''*KT 51-04-10 Comsheet ML-SF-201 แสดงยอดรวมก่อนหักส่วนลดที่เกิดจาก Form wCNSubDisc2
            olaDiscTotal.Caption = SP_FMTcCurPoint(cVB_TotalDiscAft, False, True) 'SP_FMTcCurPoint(cVB_TotalDisc, False, True) '*KT 51-04-10 Comsheet ML-SF-201 แสดงส่วนลดเฉพาะที่เกิดจากหน้า wCNSubDisc2
            olaNetSale.Caption = SP_FMTcCurPoint(cVB_NetSale, False, True)
            olaAmtTender.Caption = SP_FMTtNum2Shw(0 + cVendorCpn, False, True)
            cVB_Balance = cVB_NetSaleRound + cVB_NetDutyFreeDep - cVendorCpn '*KT 51-04-23 กรณีที่ขายแบบ Duty Free ต้องบวก Deposit Amt ด้วยครับ
'            Call SP_CRTxLogCashBalEvent(Me.Name, "Form_Load[1]", "Cash Tender", cVB_Balance)
            olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance), False, True)
            lblDepositAmt.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_NetDutyFreeDep), False, True)
    End Select
    nUT_SeqRC = 0
    '________________________________________________________________________________________________
    With ogdTender
        .HighLight = flexHighlightNever
        .Editable = flexEDNone
        .Enabled = False
    End With
    
    With ogdFunc '55-07-20 Neung TK
         .RowSel = .Row
         .Row = 0
    End With
    
    Select Case Val(Format(tVB_TransactionSub))
        Case tEN_TrnRedeem
            ocmCancel.Enabled = False
        Case tEN_TrnDepositCancel
            ocmCancel.Enabled = False
        Case tEN_TrnAddDeposit
            ocmCancel.Enabled = False
        Case Else
            ocmCancel.Enabled = True
    End Select
    'Display
    If bVB_OPOSLineDisplayEnable Then
        tPrint = "TOTAL : " & ";" & olaBalance.Caption
        Call SP_PrnPollDisplay(tPrint)
    End If
   ' ทำการเก็บข้อมูลการรูดบัตรเครดิตแล้วได้ Ontop ของ Deposit หรือ Add Deposit ก่อนหน้า
    Select Case Val(Format(tVB_TransactionSub))
        Case tEN_TrnAddDeposit
            Call W_GETxCountOntopDeposit
        Case tEN_TrnRedeem
            Call W_GETxCountOntopDeposit
    End Select

    bVB_CNTenderAct = False
    bVB_CNPayCrdNoSuccess = False
    '*KT 55-08-28 Takashimaya
    If Val(Format(tVB_TransactionSub)) = tEN_TrnReturn Then
'        ogdFunc.Select 8, 0, 8, ogdFunc.Cols - 1
'        ogdFunc.TopRow = 8
        For nLoop = 0 To nMaxRow
             If Trim(ogdFunc.TextMatrix(nLoop, 3)) = "T018" Then            '*Ao 2013/05/16   แก้ไขเรียง เมณู ใหม่
                ogdFunc.Select nLoop, 0, nLoop, ogdFunc.Cols - 1
                ogdFunc.TopRow = nLoop
                Exit For
             End If
        Next
    End If
'*Eaw 56-07-02 CommSheet TK-ISS3000-249-TR.xls ไม่ต้อง Recovery RC
'    '*Eaw 56-02-06  กู้ข้อมูลการจ่ายเงินของบิลที่ยังไม่จบ
'    If bVB_IsRecovery Then
'        If bVB_CNPaid Then
'            Call W_PRCxRecTender
'            bVB_CNPaid = False
'            bVB_IsRecovery = False
'        End If
'    End If
    '-----------------------------------------------------------------
    If tVB_MemPointID <> "" Then
        'Me.olaMEM.Caption = "BP Award ID:" & tVB_MemPointID     '*Ao 2013/02/27    TK POINT
        Me.olaMEM.Caption = "BP Award ID:" & tVB_MemPointID & tVB_MallCardIDRef  '*Tao 60-04-21 ถ้าได้ MallCardID จาก Manual ให้ใส่ *
        Me.olaMEM.Refresh '*Tao 60-05-02 แสดงเลขสมาชิก
    End If
    Call SP_DATxRsNothing(orsTemp)
    
    otmEDC.Enabled = False
    Select Case Val(Format(tVB_TransactionSub))
        Case tEN_TrnDepositCancel
            Call W_PRCxCheckEMP         '*Ao 2014/04/29    PRJ10001TK-PH2-External Spec-EDC Layaway Cancel-00.02.04 FSBT.docx
    End Select
    
    Call SP_CRTxWriteLogTerminated("ErrorNumber:0" & " ErrorDesc:" & " Form:" & Me.Name & " Events:Form_Load[2]")
    
    Exit Sub
ErrHandle:
    Call SP_DATxRsNothing(orsTemp)
    Call SP_CRTxWriteLogTerminated("ErrorNumber:" & Err.Number & " ErrorDesc:" & Err.Description & " Form:" & Me.Name & " Events:Form_Load[3]")
End Sub
'
'Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'    If UnloadMode = 0 Then Cancel = True
'End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrHandle
    otmForm.Enabled = False
    otmLock.Enabled = False
    Call SP_OBJxSetNothing(Me)
    Exit Sub
ErrHandle:
    Call SP_CRTxWriteLogTerminated("ErrorNumber:" & Err.Number & " ErrorDesc:" & Err.Description & " Form:" & Me.Name & " Events:Form_Unload[1]")
End Sub

Public Sub ocmCancel_Click()
    Dim tWhe$
    Dim cTender As Double
    Dim nKey As Long
    Dim tSql$
    Dim tWhrRC As String, tMsg As String
    
On Error GoTo ErrHandle

    If SP_CHKbDiscEmp Then                   '*Ao 2014/04/29 PRJ10001TK-PH2-External Spec-Staff Discount-00.02.03 FSBT.docx
        tMsg = "No ESC allowed with Staff Discount given"
        MsgBox tMsg, vbOKOnly, tVB_CNAppTitle
        Exit Sub
    End If
    tWhe = " WHERE  FTTmnNum='" & tVB_CNTerminalNum & "' AND FTShdTransNo='" & tVB_Running & "'"
    tWhe = tWhe & " AND  FTShdTransType='" & tVB_Transaction & "' AND FDShdTransDate='" & tUT_SaleDate & "'"
    tWhrRC = tWhe & " AND NOT(FTTdmCode='T006' AND FTVendorCnpFlag='1')"                 '*Ao 53-10-16   ถ้ามี Vendor Item ให้ ESC กลับได้
    cTender = SP_VALcText2Double(SP_SQLtFunction("SUM", "FCSrcNet", tVB_CNTblRC, tWhrRC))
    If cTender = 0 Then
        oDisc.DISxClearTender
        bVB_CNTrnDisc = False
        bVB_CNTrnDiscEx = False
        bVB_TextPromptAct = False
        bVB_FormPayAct = False
        Call SP_PRCxClearMemberPoint
        tVB_MemPointID = ""
        tVB_MallCardIDRef = "" '*Tao 60-04-21
        bVB_ChkEmpPoint = True '*Eaw 56-08-21
        Unload Me
        '*Ao 51-07-15
'        If W_DEFbReturn Then                                                                                           '*Ao 2013/06/05 Return ให้ทำงานเหมือน Sale
'            'Return
'            tVB_CNTblDT = "TPSDT" & tVB_CNTerminalNum
'            tVB_CNTblCD = "TPSCD" & tVB_CNTerminalNum
'        Else
'            '*KT 51-08-07 Communication Sheet ML-SF-327(3) Discount of Return Transaction.xls
'            '-----------------------------------------------------------------------------------------------------------------------------------
'            Select Case Val(Format(tVB_TransactionSub))
'                Case tEN_TrnReturn, tEN_TrnReturnTaxInc, tEN_TrnReturnManualTax
'                    If bVB_Reference Then
'                        nKey = SP_MSGnShowing("tms_CN005171", tCS_CNMsgYesNoDefYes)     'คุณต้องการคืนสินค้า โดยอ้างถึงเลขที่ใบเสร็จอย่างย่อหรือไม่
'                        Select Case nKey
'                            Case vbNo      'cancel
'                                tVB_CNTblDT = "TPSDT" & tVB_CNTerminalNum
'                                tVB_CNTblCD = "TPSCD" & tVB_CNTerminalNum
'                            Case vbYes
'                                'Clear
'                                tSql = "TRUNCATE TABLE " & tVB_CNTblDTTmp
'                                Call SP_SQLvExecute(tSql)
'                                tSql = "TRUNCATE TABLE " & tVB_CNTblCDTmp
'                                Call SP_SQLvExecute(tSql)
'
'                                Call UT_GenSaleNewDoc
'
'                                tVB_CNTblDT = "TPSDT" & tVB_CNTerminalNum
'                                tSql = "DELETE " & tVB_CNTblDT
'                                tSql = tSql & vbCrLf & tWhe
'                                Call SP_SQLvExecute(tSql)
'
'                                tVB_CNTblCD = "TPSCD" & tVB_CNTerminalNum
'                                tSql = "DELETE " & tVB_CNTblCD
'                                tSql = tSql & vbCrLf & tWhe
'                                Call SP_SQLvExecute(tSql)
'
'                                tVB_CNMnuCurrent = "1700"
'                                wCNBegin.Show vbModal
'                        End Select
'                    End If
'                Case Else
'                    '***KT 51-11-20 ป้องกันการไม่เข้า wMain.Form_Activate
'                    tVB_CNTblDT = "TPSDT" & tVB_CNTerminalNum
'                    tVB_CNTblCD = "TPSCD" & tVB_CNTerminalNum
'                    '*KT 52-05-02 PH 1.8 ML-SF-052  Other Receive เมื่อกด Total แล้วเข้าหน้าจอ Tender โดยที่ยังไม่ได้มีการชำระเงิน จะไม่สามารถกด ESC , Cancel  เพื่อกลับไปหน้าจอขายได้  แต่การขายธรรมดาทำได้
'                    Select Case Val(Format(tVB_TransactionSub))
'                        Case tEN_TrnOtherReceive, tEN_TrnOtherDisbursement '*KT 54-09-26 V.4.0.1 Req3 Phase4.0(Member Point) Member Point on Other Receive ให้หน้าจอ Tender สามารถ Back กลับไปที่หน้าจอ Discount ได้ในกรณี OtherReceive
'                            '*Bump 56-02-22 เพิ่ม tEN_TrnOtherReceive อ้างถึง TK-ISS3000-049 Other recive can not retrun from payment to entry page
'                            'none
'                        Case Else
'                            Call wMain.ogdForm_KeyDown(vbKeyHome, 0)
'                    End Select
'            End Select
'            '-----------------------------------------------------------------------------------------------------------------------------
'        End If
        '___________________________________
        tVB_CNTblDT = "TPSDT" & tVB_CNTerminalNum           '*Ao 2013/06/05 Return ให้ทำงานเหมือน Sale
        tVB_CNTblCD = "TPSCD" & tVB_CNTerminalNum
        '*KT 52-05-02 PH 1.8 ML-SF-052  Other Receive เมื่อกด Total แล้วเข้าหน้าจอ Tender โดยที่ยังไม่ได้มีการชำระเงิน จะไม่สามารถกด ESC , Cancel  เพื่อกลับไปหน้าจอขายได้  แต่การขายธรรมดาทำได้
        Select Case Val(Format(tVB_TransactionSub))
            Case tEN_TrnOtherReceive, tEN_TrnOtherDisbursement '*KT 54-09-26 V.4.0.1 Req3 Phase4.0(Member Point) Member Point on Other Receive ให้หน้าจอ Tender สามารถ Back กลับไปที่หน้าจอ Discount ได้ในกรณี OtherReceive
                '*Bump 56-02-22 เพิ่ม tEN_TrnOtherReceive อ้างถึง TK-ISS3000-049 Other recive can not retrun from payment to entry page
                'none
            Case Else
                Call wMain.ogdForm_KeyDown(vbKeyHome, 0)
        End Select
    Else
        Call SP_MSGnShowing("tms_CN005085", nCS_Error)
    End If
    
    Exit Sub
ErrHandle:
    Call SP_CRTxLogPrg("ErrorNumber:" & Err.Number & " ErrorDesc:" & Err.Description & " Form:" & Me.Name & " Control:ocmCancel" & " Events:ocmCancel_Click")
End Sub

Private Sub ocmOK_Click()
'*KT 51-11-13

    Dim tWhe$
    Dim tSql$
    Dim cMeaningDef As Double
    Dim orsTemp As ADODB.Recordset
    Dim nI As Long
    Dim tPrint As String '*Eaw 56-07-01 CommSheet TK-ISS3000-246-TR.xls
    Dim tTender As String, tTenderName As String
    Dim cWalletBalance As Double
    Dim tTmpRmk As String
    Dim dSrcDate As Date
    Dim dTransDate As Date
    Dim tTransID As String
    Dim nTmp As Integer
On Error GoTo ErrHandle
        
        If bVB_CNTenderAct Then
            Call SP_DATxRsNothing(orsTemp)
            Exit Sub
        End If
        
        If CDbl(Val(Format(olaBalance.Caption))) > cCS_CNMaxReceive Then
            Call SP_MSGnShowing("tms_CN005152", nCS_Error)
            bVB_CNTenderAct = False
            Call SP_DATxRsNothing(orsTemp)
            Exit Sub
        End If
        
        With ogdFunc
'            '*Tao 60-09-01 ของเดิม
'            '*KT 51-07-08 Communication Sheet ML-SF-137(2) Payment menu is not complete
'            If .TextMatrix(.Row, 6) <> "1" Then
'                Select Case Val(Format(tVB_TransactionSub))
'                    Case tEN_TrnReturn, tEN_TrnCancelVoucher, tEN_TrnDepositCancel, tEN_TrnReturnTaxInc, tEN_TrnReturnManualTax
'                        'Call SP_MSGnShowing("tms_CN005158", nCS_Error) '55-07-20 Neung TK เปลี่ยน Set สีใน Grid
'                    Case Else
'                        'Call SP_MSGnShowing("tms_CN005157", nCS_Error) '55-07-20 Neung TK เปลี่ยน Set สีใน Grid
'                End Select
'                bVB_CNTenderAct = False
'                Call SP_DATxRsNothing(orsTemp)
'                Exit Sub
'            End If
'
'            If .TextMatrix(.Row, 7) <> "1" And bVB_ReEntryMode Then
'                bVB_CNTenderAct = False
'                Call SP_DATxRsNothing(orsTemp)
'                Exit Sub
'            End If
            
            '*Tao 60-09-01 ถ้าเป็น WalletPay ไม่ต้องเช็ค
            Select Case .TextMatrix(.Row, 3)
            Case "T019", "T020", "T021", "T022"
                '
            Case Else
                If .TextMatrix(.Row, 6) <> "1" Then
                    bVB_CNTenderAct = False
                    Call SP_DATxRsNothing(orsTemp)
                    Exit Sub
                End If
                
                If .TextMatrix(.Row, 7) <> "1" And bVB_ReEntryMode Then
                    bVB_CNTenderAct = False
                    Call SP_DATxRsNothing(orsTemp)
                    Exit Sub
                End If
            End Select
            '==================================
            
            '*KT 54-08-03 V.4.0.1 Req.2 Phase4.0(Member Point) Member Point Based on Tender
            nVB_CurTenderSeqNo = 0
            For nI = 0 To UBound(aVB_TenderMnu)
                If UCase(aVB_TenderMnu(nI).ID) = Trim(UCase(.TextMatrix(.Row, 3))) Then
                    nVB_CurTenderSeqNo = nI
                    Exit For
                End If
            Next nI
            
            If SP_CHKbAllowMemEmp(nVB_CurTenderSeqNo, Trim(UCase(.TextMatrix(.Row, 3)))) Then
                Call SP_MSGnShowing("tms_CN008039", nCS_Error)      '*Ao 2013/03/15
                bVB_CNTenderAct = False
                Call SP_OBJxSetFocus(ogdFunc)
                Exit Sub
            End If
            
            otmForm.Enabled = False
            otmLock.Enabled = False
                                     
            tTender = Trim(.TextMatrix(.Row, 3))                                     '*Ao 2015/01/06
            tTenderName = Trim(.TextMatrix(.Row, 5))
            Call SP_CRTxLogPrg("Payment = " & tTender & ":" & tTenderName)
            
            Select Case Trim(UCase(.TextMatrix(.Row, 3)))
                    Case "T001" 'cash
                        '*Eaw 56-01-03 ก่อนเข้า  Tender ไม่มีการ Rounding เมื่อต้องการชำระด้วย cash ให้ใช้ค่า Balance+Rounding เป็นยอดที่ต้องชำระ
                        cVB_RoundRcv = SP_FMTcCurPoint(IIf(bVB_CNAlwRnd, SP_FMTcCurPoint(UT_GetnRoundDiff(cVB_Balance), True, True), 0), True, False)  '*KT 51-04-05 Comsheet ML-SF-172,ML-SF-173
                        cVB_Balance = cVB_Balance + cVB_RoundRcv
'                        Select Case Val(Format(tVB_TransactionSub))
'                            Case tEN_TrnDepositCancel, tEN_TrnAddDepositCancel
'                                olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance), False, True)  'แสดง Balance ที่มีการปัดเศษ เพราะ cash ต้องมีการ Rouding
'                            Case Else
'                                olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance), False, True) 'แสดง Balance ที่มีการปัดเศษ เพราะ cash ต้องมีการ Rouding
'                        End Select
                        olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance), False, True) '*Eaw 56-01-29   แสดง Balance ที่มีการปัดเศษ เพราะ cash ต้องมีการ Rouding
                        '*Eaw 56-07-01 CommSheet TK-ISS3000-246-TR.xls
                        'Display
                        If bVB_OPOSLineDisplayEnable Then
                            tPrint = "TOTAL : " & ";" & olaBalance.Caption
                            Call SP_PrnPollDisplay(tPrint)
                        End If
'                        cVB_RoundRcv = 0
                        '*Eaw 56-07-01 CommSheet TK-ISS3000-250-TR.xls
'                        If bVB_CNPayCrdNoSuccess Then
'                            MsgBox "Lock Tender", vbCritical
'                            '*Eaw 56-01-29   แสดง Balance ที่มีการปัดเศษ เพราะ cash ต้องมีการ Rouding
'                            cVB_Balance = cVB_Balance - cVB_RoundRcv
'                            cVB_RoundRcv = 0
'                             olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance), False, True)
'                             '-----------------------------------------------------------------------------------------------------------------------------------------
'                            bVB_CNTenderAct = False
'                            otmForm.Enabled = True
'                            otmLock.Enabled = True
'                            Call SP_DATxRsNothing(orsTemp)
'                            Call SP_OBJxSetFocus(ogdFunc)
'                            Exit Sub
'                        End If
                        bVB_CNTenderAct = True
                        'nVB_CurTenderSeqNo = 0 '*KT 54-08-03 V.4.0.1 Req.2 Phase4.0(Member Point) Member Point Based on Tender
                        tVB_CNMnuCurrent = .TextMatrix(.Row, 2)
                        If aVB_TenderMnu(nVB_CurTenderSeqNo).Autho Then
                                If Not SP_USRbMnuAccept(nEN_Usr2AthRead, tVB_CNMnuCurrent) Then
                                    bVB_CNTenderAct = False
                                    otmForm.Enabled = True
                                    otmLock.Enabled = True
                                    Call SP_DATxRsNothing(orsTemp)
                                    Call SP_OBJxSetFocus(ogdFunc) ' ogdFunc.SetFocus
                                    Exit Sub
                                End If
                        End If
                        tVB_TenderCode = Trim(UCase(.TextMatrix(.Row, 3)))
                        tVB_TenderType = .TextMatrix(.Row, 5)
                        tVB_TenderName = Trim(.TextMatrix(.Row, 4))
                        wCNPayCash.Caption = Trim(.TextMatrix(.Row, 4))
                        wCNPayCash.Show vbModal
                        If cVB_Pay = 0 And cVB_Balance <> 0 Then
                            bVB_CNTenderAct = False
                            otmForm.Enabled = True
                            otmLock.Enabled = True
                            Call SP_DATxRsNothing(orsTemp)
                            Call SP_OBJxSetFocus(ogdFunc)
'                            cVB_Balance = cVB_Balance - cVB_RoundRcv '*Eaw 56-07-04
                            olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance), False, True) '*Eaw 56-01-29  แสดง Balance ที่มีการปัดเศษ เพราะ cash ต้องมีการ Rouding
                            
                            cVB_RoundRcv = 0  '*Eaw 56-01-03 ก่อนเข้า  Tender ไม่มีการ Rounding เมื่อต้องการชำระด้วย cash ให้ใช้ค่า Balance+Rounding เป็นยอดที่ต้องชำระ
                            '*Eaw 56-07-01 CommSheet TK-ISS3000-246-TR.xls
                            'Display
                            If bVB_OPOSLineDisplayEnable Then
                                tPrint = "TOTAL : " & ";" & olaBalance.Caption
                                Call SP_PrnPollDisplay(tPrint)
                            End If
                            Exit Sub   '*KT 51-06-04 Comsheet ML-SF-320 ยอมให้รับชำระเป็น 0 กรณีที่มีการให้ส่วนลดทั้งหมด
                        End If
                        Call SP_OBJxSetFocus(ogdFunc)
                        Call W_PRCxAddPay
                        bVB_CNTenderAct = False
'*KT 53-01-06 ทำการย้ายอยู่ที่ Else เพราะหลังจากที่เรียกคำสั่ง Unload wSub แล้วเกิดมีการเรียก Timer ให้ Enable = True จะเข้า Form_Load อีกครั้ง
'                        otmForm.Enabled = True
'                        otmLock.Enabled = True
                        If cVB_Balance = 0 Then
'*KT 53-02-03 ป้องกันการ Terminated และใน Unload Me  มีการสั่ง otmForm.Enabled = False แล้ว
'                            otmForm.Enabled = False
'                            otmLock.Enabled = False
                            Call SP_DATxRsNothing(orsTemp)
                            Unload Me
                            Exit Sub
                        Else
                            otmForm.Enabled = True
                            otmLock.Enabled = True
                        End If
                    Case "T002", "T003" 'Credit card
'                        Select Case Val(Format(tVB_TransactionSub))
'                            Case tEN_TrnDepositCancel, tEN_TrnAddDepositCancel
'                                olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance - cVB_RoundRcv), False, True) 'แสดง Balance ที่มีการบวกเศษด้วย เพราะการจ่ายด้วยเครดิตสามารถจ่ายเป็นเศษได้
'                            Case Else
'                                olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance - cVB_RoundRcv), False, True) 'แสดง Balance ที่มีการบวกเศษด้วย เพราะการจ่ายด้วยเครดิตสามารถจ่ายเป็นเศษได้
'                        End Select
                        olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance), False, True)  '*Eaw 56-01-29  แสดง Balance ที่มีการปัดเศษ เพราะ cash ต้องมีการ Rouding
                        tVB_TypeOfEDC = "4"
                        bVB_CNTenderAct = True
                        '*KT 54-08-03 V.4.0.1 Req.2 Phase4.0(Member Point) Member Point Based on Tender
'                        If Trim(UCase(.TextMatrix(.Row, 3))) = "T002" Then
'                            nVB_CurTenderSeqNo = 1
'                        Else
'                            nVB_CurTenderSeqNo = 2
'                        End If
'                       '*Eaw  55-11-27   Add Case "T002","T003"
                        '----------------------------------------------------------------------------------------------------------------
                        Select Case Trim(UCase(.TextMatrix(.Row, 3)))
                            Case "T002"
                                tVB_CNMnuCurrent = "1817"
                                If aVB_TenderMnu(nVB_CurTenderSeqNo).Autho Then
'                                    If Not SP_USRbMnuAccept(nEN_Usr2AthRead, tVB_CNMnuCurrent) Then Exit Sub
                                    If Not SP_USRbMnuAccept(nEN_Usr2AthRead, tVB_CNMnuCurrent) Then  '*Eaw 56-07-02
                                        bVB_CNTenderAct = False
                                        otmForm.Enabled = True
                                        otmLock.Enabled = True
                                        Call SP_DATxRsNothing(orsTemp)
                                        Call SP_OBJxSetFocus(ogdFunc) ' ogdFunc.SetFocus
                                        Exit Sub
                                    End If
                                End If
                                tVB_TenderCode = "T002"
                                tVB_TenderType = aVB_TenderMnu(nVB_CurTenderSeqNo).Name '"Credit Card" '*KT 53-10-15 Commsheet หมายเลข PH3-ML-SF-131
'                                wCNPayCredit.Caption = aVB_TenderMnu(nVB_CurTenderSeqNo).Menu '"Credit Card" '*KT 53-10-15 Commsheet หมายเลข PH3-ML-SF-131
                                tVB_TenderName = aVB_TenderMnu(nVB_CurTenderSeqNo).DescSale
                                If SP_EDCbCheckEDC Then
'                                    wCNPayCreditPay.Show vbModal
                                    '*Eaw 56-06-11
                                    Select Case Val(Format(tVB_TransactionSub))
                                        Case tEN_TrnReturnManualTax, tEN_TrnReturnTaxInc, tEN_TrnSaleManualTax, tEN_TrnSaleTaxInc
'                                            wCNPayCredit.Show vbModal
                                            wCNPayCreditRE.Show vbModal '*Eaw 56-08-28 CommSheet TK-ISS3000-323-UAT.xlsx
                                        Case Else
                                            wCNPayCreditPay.Show vbModal
                                    End Select
                                Else
'                                    wCNPayCredit.Show vbModal
                                    '*Eaw 56-08-28 CommSheet TK-ISS3000-323-UAT.xlsx
                                    Select Case Val(Format(tVB_TransactionSub))
                                        Case tEN_TrnReturnManualTax, tEN_TrnReturnTaxInc, tEN_TrnSaleManualTax, tEN_TrnSaleTaxInc
                                            wCNPayCreditRE.Show vbModal
                                        Case Else
                                            wCNPayCredit.Show vbModal
                                    End Select
                                End If
'                                wCNPayCreditPay.Show vbModal        '*Ao 2013/01/03
'                                wCNPayCredit.Show vbModal
                            Case "T003"
                                tVB_CNMnuCurrent = "1818"
                                If aVB_TenderMnu(nVB_CurTenderSeqNo).Autho Then
'                                    If Not SP_USRbMnuAccept(nEN_Usr2AthRead, tVB_CNMnuCurrent) Then Exit Sub
                                    If Not SP_USRbMnuAccept(nEN_Usr2AthRead, tVB_CNMnuCurrent) Then  '*Eaw 56-07-02
                                        bVB_CNTenderAct = False
                                        otmForm.Enabled = True
                                        otmLock.Enabled = True
                                        Call SP_DATxRsNothing(orsTemp)
                                        Call SP_OBJxSetFocus(ogdFunc) ' ogdFunc.SetFocus
                                        Exit Sub
                                    End If
                                End If
                                tVB_TenderCode = "T003"
                                tVB_TenderType = aVB_TenderMnu(nVB_CurTenderSeqNo).Name     'Hire Purchase  For TK Project
'                                wCNPayCredit.Caption = aVB_TenderMnu(nVB_CurTenderSeqNo).Menu   'Hire Purchase  For TK Project
                                tVB_TenderName = aVB_TenderMnu(nVB_CurTenderSeqNo).DescSale
                                If SP_EDCbCheckEDC Then
'                                    wCNPayCreditPay.Show vbModal     '*Eaw 56-06-11
                                    Select Case Val(Format(tVB_TransactionSub))
                                        Case tEN_TrnReturnManualTax, tEN_TrnReturnTaxInc, tEN_TrnSaleManualTax, tEN_TrnSaleTaxInc
'                                            wCNPayCredit.Show vbModal
                                            wCNPayCreditRE.Show vbModal '*Eaw 56-08-28 CommSheet TK-ISS3000-323-UAT.xlsx
                                        Case Else
                                            wCNPayCreditPay.Show vbModal
                                    End Select
                                Else
'                                    wCNPayCredit.Show vbModal
                                    '*Eaw 56-08-28 CommSheet TK-ISS3000-323-UAT.xlsx
                                    Select Case Val(Format(tVB_TransactionSub))
                                        Case tEN_TrnReturnManualTax, tEN_TrnReturnTaxInc, tEN_TrnSaleManualTax, tEN_TrnSaleTaxInc
                                            wCNPayCreditRE.Show vbModal
                                        Case Else
                                            wCNPayCredit.Show vbModal
                                    End Select
                                End If
'                                wCNPayCreditPay.Show vbModal        '*Ao 2013/01/03
'                                wCNPayCredit.Show vbModal
                        End Select
                        '-------------------------------------------------------------------------------------------------------
                        
'                        DoEvents '*KT 53-01-14 DoEvents
                        'wCNPayCreditType.Show vbModal    '*Eaw 55-11-27

                        If cVB_Pay = 0 Then
                            If cVB_Balance > 0 Then
                                '*Eaw 56-01-04 แก้ไขการปัด Rounding
'                                olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance), False, True) 'แสดง Balance ที่มีปัดเศษเหมือนเดิมเพราะไม่มีการชำระด้วยเครดิต
'                                Select Case Val(Format(tVB_TransactionSub))
'                                    Case tEN_TrnDepositCancel, tEN_TrnAddDepositCancel
'                                        olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance - cVB_RoundRcv), False, True) 'แสดง Balance ที่มีการบวกเศษด้วย เพราะการจ่ายด้วยเครดิตสามารถจ่ายเป็นเศษได้
'                                    Case Else
'                                        olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance - cVB_RoundRcv), False, True) 'แสดง Balance ที่มีการบวกเศษด้วย เพราะการจ่ายด้วยเครดิตสามารถจ่ายเป็นเศษได้
'                                End Select
                                cVB_Balance = cVB_Balance - cVB_RoundRcv
                                olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance), False, True)  '*Eaw 56-01-29แสดง Balance ที่มีการปัดเศษ เพราะ cash ต้องมีการ Rouding
                                cVB_RoundRcv = 0
                            '________________________________________
                            Else
                                Call SP_GETxBalanceRnd  '*Eaw 56-07-01 ทำเป็นฟังก์ชันกลาง
                                '*KT 51-11-12 อย่าลืมทำที่ Credit Store ด้วย
                                '---------------------------------------------------------------------------------------------------------
'                                tSql = "SELECT FNMenMin,FNMenValue FROM TSysMeaning"
'                                Call SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine, , , , , True, True) 'read master
'                                If Not (orsTemp.BOF Or orsTemp.EOF) Then
'                                    cMeaningDef = CDbl(Val(Format(SP_FEDvChkNumeric(orsTemp, "FNMenValue"))) / 100)
'                                Else
'                                    cMeaningDef = 0
'                                End If
'                                Call SP_DATxRsNothing(orsTemp)
'                                If bVB_CNAlwRnd Then
'                                    Select Case nVB_CNTypeRnd
'                                        Case eRndUp
'                                            If SP_FMTcCurPoint(cVB_Balance - cVB_RoundRcv, True, False) < 1 Then
'                                                bVB_CNPayCrdNoSuccess = True
'                                                olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance - cVB_RoundRcv), False, True) 'แสดง Balance ที่มีการบวกเศษด้วย เพราะการจ่ายด้วยเครดิตสามารถจ่ายเป็นเศษได้
'                                            Else
'                                                bVB_CNPayCrdNoSuccess = False
'                                            End If
'                                        Case eRndDn
'                                            If SP_FMTcCurPoint(cVB_Balance - cVB_RoundRcv, True, False) < 1 Then
'                                                bVB_CNPayCrdNoSuccess = True
'                                                olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance - cVB_RoundRcv), False, True) 'แสดง Balance ที่มีการบวกเศษด้วย เพราะการจ่ายด้วยเครดิตสามารถจ่ายเป็นเศษได้
'                                            Else
'                                                bVB_CNPayCrdNoSuccess = False
'                                            End If
'                                        Case eRndMeaning
'                                            If SP_FMTcCurPoint(cVB_Balance - cVB_RoundRcv, True, False) <= cMeaningDef Then
'                                                bVB_CNPayCrdNoSuccess = True
'                                                olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance - cVB_RoundRcv), False, True) 'แสดง Balance ที่มีการบวกเศษด้วย เพราะการจ่ายด้วยเครดิตสามารถจ่ายเป็นเศษได้
'                                            Else
'                                                bVB_CNPayCrdNoSuccess = False
'                                            End If
'                                    End Select
'                                Else
'                                    If SP_FMTcCurPoint(cVB_Balance - cVB_RoundRcv, True, False) < 1 Then
'                                        bVB_CNPayCrdNoSuccess = True
'                                       olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance - cVB_RoundRcv), False, True) 'แสดง Balance ที่มีการบวกเศษด้วย เพราะการจ่ายด้วยเครดิตสามารถจ่ายเป็นเศษได้
'                                    Else
'                                        bVB_CNPayCrdNoSuccess = False
'                                    End If
'                                End If
                                '---------------------------------------------------------------------------------------------------------
                            End If
                            bVB_CNTenderAct = False
'*KT 53-02-03 ป้องกันการ Terminated และใน Unload Me  มีการสั่ง otmForm.Enabled = False แล้ว
'                            otmForm.Enabled = True
'                            otmLock.Enabled = True
                            Call SP_DATxRsNothing(orsTemp)
                            Call SP_OBJxSetFocus(ogdFunc)
                            Exit Sub
                        End If
                        Call SP_OBJxSetFocus(ogdFunc)
'                        Call W_PRCxAddPayCrdCard           '*Ao 2013/11/14     CommSheet TK-ISS3000-FAPL-152
                        '-------------------------------------------------------------------------
                        Select Case Val(Format(tVB_TransactionSub))
                            Case tEN_TrnReturnManualTax, tEN_TrnReturnTaxInc, tEN_TrnSaleManualTax, tEN_TrnSaleTaxInc
                                Call W_PRCxAddPayCrdCard
                            Case Else
                                If Trim(tVB_CrdCardNo) <> "" Then
                                    Call W_PRCxAddPayCrdCard
                                End If
                        End Select
                        '-------------------------------------------------------------------------
                        bVB_CNTenderAct = False
'*KT 53-01-06 ทำการย้ายอยู่ที่ Else เพราะหลังจากที่เรียกคำสั่ง Unload wSub แล้วเกิดมีการเรียก Timer ให้ Enable = True จะเข้า Form_Load อีกครั้ง
'                        otmForm.Enabled = True
'                        otmLock.Enabled = True
                        If cVB_Balance = 0 Then
'*KT 53-02-03 ป้องกันการ Terminated และใน Unload Me  มีการสั่ง otmForm.Enabled = False แล้ว
'                            otmForm.Enabled = False
'                            otmLock.Enabled = False
                            Call SP_DATxRsNothing(orsTemp)
                            Unload Me
                            Exit Sub
                        Else
                            otmForm.Enabled = True
                            otmLock.Enabled = True
                        End If
                    Case "T004" 'Gift Voucher
                        '*Eaw 56-01-04 แก้ไขการปัด Rounding
'                        Select Case Val(Format(tVB_TransactionSub))
'                            Case tEN_TrnDepositCancel, tEN_TrnAddDepositCancel
'                                olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance - cVB_RoundRcv), False, True) 'แสดง Balance ที่มีการบวกเศษด้วย เพราะการจ่ายด้วยเครดิตสามารถจ่ายเป็นเศษได้
'                            Case Else
'                                olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance - cVB_RoundRcv), False, True) 'แสดง Balance ที่มีการบวกเศษด้วย เพราะการจ่ายด้วยเครดิตสามารถจ่ายเป็นเศษได้
'                        End Select
                        olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance), False, True)  '*Eaw 56-01-29แสดง Balance ที่มีการปัดเศษ เพราะ cash ต้องมีการ Rouding
                        '----------------------------------------------------------------------------------------
                        '*Eaw 56-07-01 CommSheet TK-ISS3000-250-TR.xls
'                        If bVB_CNPayCrdNoSuccess Then
'                            MsgBox "Lock Tender", vbCritical
'                            bVB_CNTenderAct = False
'                            otmForm.Enabled = True
'                            otmLock.Enabled = True
'                            Call SP_DATxRsNothing(orsTemp)
'                            Call SP_OBJxSetFocus(ogdFunc)
'                            Exit Sub
'                        End If
                        bVB_CNTenderAct = True
                        'nVB_CurTenderSeqNo = 3 '*KT 54-08-03 V.4.0.1 Req.2 Phase4.0(Member Point) Member Point Based on Tender
                        tVB_CNMnuCurrent = .TextMatrix(.Row, 2)
                        If aVB_TenderMnu(nVB_CurTenderSeqNo).Autho Then
                                If Not SP_USRbMnuAccept(nEN_Usr2AthRead, tVB_CNMnuCurrent) Then
                                    bVB_CNTenderAct = False
                                    otmForm.Enabled = True
                                    otmLock.Enabled = True
                                    Call SP_DATxRsNothing(orsTemp)
                                    Call SP_OBJxSetFocus(ogdFunc)
                                    Exit Sub
                                End If
                         End If
                        tVB_TenderCode = Trim(UCase(.TextMatrix(.Row, 3)))
                        tVB_TenderType = .TextMatrix(.Row, 5)
                        tVB_TenderName = Trim(.TextMatrix(.Row, 4))
'                        wCNPayOth.tW_Name = "Gift Voucher No. :"
                        tVB_Tender = "Gift Voucher No. :"
'                        wCNPayOth.Caption = "Gift Voucher"
                        wCNPayOth.Show vbModal
                         If cVB_Pay = 0 Then
                            '*Eaw 56-01-04 แก้ไขการปัด Rounding
'                            Select Case Val(Format(tVB_TransactionSub))
'                                Case tEN_TrnDepositCancel, tEN_TrnAddDepositCancel
'                                    olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance - cVB_RoundRcv), False, True) 'แสดง Balance ที่มีการบวกเศษด้วย เพราะการจ่ายด้วยเครดิตสามารถจ่ายเป็นเศษได้
'                                Case Else
'                                    olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance - cVB_RoundRcv), False, True) 'แสดง Balance ที่มีการบวกเศษด้วย เพราะการจ่ายด้วยเครดิตสามารถจ่ายเป็นเศษได้
'                            End Select
                            cVB_Balance = cVB_Balance - cVB_RoundRcv
                            olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance), False, True)  '*Eaw 56-01-29แสดง Balance ที่มีการปัดเศษ เพราะ cash ต้องมีการ Rouding
                            cVB_RoundRcv = 0
                            '________________________________________
                            bVB_CNTenderAct = False
                            otmForm.Enabled = True
                            otmLock.Enabled = True
                            Call SP_DATxRsNothing(orsTemp)
                            Call SP_OBJxSetFocus(ogdFunc)
                            Exit Sub
                        End If
                        Call SP_OBJxSetFocus(ogdFunc)
                        Call W_PRCxAddPayGiftVoucher
                        bVB_CNTenderAct = False
'*KT 53-01-06 ทำการย้ายอยู่ที่ Else เพราะหลังจากที่เรียกคำสั่ง Unload wSub แล้วเกิดมีการเรียก Timer ให้ Enable = True จะเข้า Form_Load อีกครั้ง
'                        otmForm.Enabled = True
'                        otmLock.Enabled = True
                        If cVB_Balance = 0 Then
'*KT 53-02-03 ป้องกันการ Terminated และใน Unload Me  มีการสั่ง otmForm.Enabled = False แล้ว
'                            otmForm.Enabled = False
'                            otmLock.Enabled = False
                            Call SP_DATxRsNothing(orsTemp)
                            Unload Me
                            Exit Sub
                        Else
                            otmForm.Enabled = True
                            otmLock.Enabled = True
                        End If
                    Case "T005", "T014" 'Spacial Gift Voucher '*KT 54-07-29 V.4.0.1 Req.5 Phase4.0.1(Member Point) Add new Tender Type ,Add 1 Type of Special Gift Voucher apply payment same as “Special Gift Voucher"
                         '*Eaw 56-01-04 แก้ไขการปัด Rounding
'                        Select Case Val(Format(tVB_TransactionSub))
'                            Case tEN_TrnDepositCancel, tEN_TrnAddDepositCancel
'                                olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance - cVB_RoundRcv), False, True) 'แสดง Balance ที่มีการบวกเศษด้วย เพราะการจ่ายด้วยเครดิตสามารถจ่ายเป็นเศษได้
'                            Case Else
'                                olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance - cVB_RoundRcv), False, True) 'แสดง Balance ที่มีการบวกเศษด้วย เพราะการจ่ายด้วยเครดิตสามารถจ่ายเป็นเศษได้
'                        End Select
                        olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance), False, True)  '*Eaw 56-01-29แสดง Balance ที่มีการปัดเศษ เพราะ cash ต้องมีการ Rouding
                        '*Eaw 56-07-01 CommSheet TK-ISS3000-250-TR.xls
'                        If bVB_CNPayCrdNoSuccess Then
'                            MsgBox "Lock Tender", vbCritical
'                            bVB_CNTenderAct = False
'                            otmForm.Enabled = True
'                            otmLock.Enabled = True
'                            Call SP_DATxRsNothing(orsTemp)
'                            Call SP_OBJxSetFocus(ogdFunc)
'                            Exit Sub
'                        End If
                        If Trim(UCase(.TextMatrix(.Row, 3))) = "T005" Then  '*Ao 2012/07/27 PRJ10001TK-External Specification Takashimaya Gift Voucher-00.01.00.doc
                            If W_CHKbCheckTKAllowEmp(nVB_CurTenderSeqNo) Then
                                Call SP_MSGnShowing("tms_CN008034", nCS_Information)
                                bVB_CNTenderAct = False
                                otmForm.Enabled = True
                                otmLock.Enabled = True
                                Call SP_DATxRsNothing(orsTemp)
                                Call SP_OBJxSetFocus(ogdFunc)
                                Exit Sub
                            End If
                        End If
                        bVB_CNTenderAct = True
                        'ก่อนการชำระเงินถ้ามีการชำระประเภทอื่นที่ไม่ใช่Spacial Gift Voucher/Special Coupon จะไม่ยอมให้ชำระเงินประเภทนี้อีกแล้ว
                        'จะยกเลิกหน้านี้แล้วเข้ามาใหม่
                        tWhe = " WHERE  FTTmnNum='" & tVB_CNTerminalNum & "' AND FTShdTransNo='" & tVB_Running & "'"
                        tWhe = tWhe & " AND  FTShdTransType='" & tVB_Transaction & "' AND FDShdTransDate='" & tUT_SaleDate & "'"
                        tWhe = tWhe & " AND FTTdmCode NOT IN ('T005','T007','T014')  AND NOT (FTTdmCode IN ('T006') AND ISNULL(FTVendorCnpFlag,'') = '1')"  '*Ao 2011/01/26 CommSheet ML-UAT-PRV-032-Special Gift Voucher can't use with Vendor Coupon (item).xls
'                        If SP_VALcText2Double(SP_SQLtFunction("SUM", "FCSrcNet", tVB_CNTblRC, tWhe)) = 0 Then               '*Ao 2013/03/19 CommSheet TK-ISS3000-109 00.00.01.xlsx
                        If bVB_CNTenderAct Then
                            'nVB_CurTenderSeqNo = 4 '*KT 54-08-03 V.4.0.1 Req.2 Phase4.0(Member Point) Member Point Based on Tender
                            tVB_CNMnuCurrent = .TextMatrix(.Row, 2)
                            If aVB_TenderMnu(nVB_CurTenderSeqNo).Autho Then
                                If Not SP_USRbMnuAccept(nEN_Usr2AthRead, tVB_CNMnuCurrent) Then
                                    bVB_CNTenderAct = False
                                    otmForm.Enabled = True
                                    otmLock.Enabled = True
                                    Call SP_DATxRsNothing(orsTemp)
                                    Call SP_OBJxSetFocus(ogdFunc)
                                    Exit Sub
                                End If
                             End If
                            tVB_TenderCode = Trim(UCase(.TextMatrix(.Row, 3)))
                            tVB_TenderType = .TextMatrix(.Row, 5)
                            tVB_TenderName = Trim(.TextMatrix(.Row, 4))
                            '*Eaw 55-11-27 add if
                            '-----------------------------------------------------------------------
                            Select Case tVB_TenderCode
                            Case "T005"
'                                wCNPayOth.tW_Name = "Gift Voucher No. :"       '*Ao 2013/03/20  TK
                                tVB_Tender = "Gift Voucher No. :"
                                wCNPayOth.Show vbModal
                            Case "T014"
'                                wCNPayOth.tW_Name = "Coupon No. :"
                                tVB_Tender = "Coupon No. :"
                                wCNPayOth.Show vbModal
                            End Select
                            '------------------------------------------------------------------------
'                            wCNPayOth.tW_Name = "Gift Voucher No"      '*Eaw 55-11-27
'                            wCNPayOth.Caption = "Special Gift Voucher"     '*Eaw 55-11-27
'                            wCNPayOth.Show vbModal      '*Eaw 55-11-27
                             If cVB_Pay = 0 Then
                                    '*Eaw 56-01-04 แก้ไขการปัด Rounding
'                                    Select Case Val(Format(tVB_TransactionSub))
'                                        Case tEN_TrnDepositCancel, tEN_TrnAddDepositCancel
'                                            olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance - cVB_RoundRcv), False, True) 'แสดง Balance ที่มีการบวกเศษด้วย เพราะการจ่ายด้วยเครดิตสามารถจ่ายเป็นเศษได้
'                                        Case Else
'                                            olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance - cVB_RoundRcv), False, True) 'แสดง Balance ที่มีการบวกเศษด้วย เพราะการจ่ายด้วยเครดิตสามารถจ่ายเป็นเศษได้
'                                    End Select
                                    cVB_Balance = cVB_Balance - cVB_RoundRcv
                                    olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance), False, True)  '*Eaw 56-01-29แสดง Balance ที่มีการปัดเศษ เพราะ cash ต้องมีการ Rouding
                                    cVB_RoundRcv = 0
                                    '________________________________________
                                    bVB_CNTenderAct = False
                                    otmForm.Enabled = True
                                    otmLock.Enabled = True
                                    Call SP_DATxRsNothing(orsTemp)
                                    Call SP_OBJxSetFocus(ogdFunc)
                                Exit Sub
                            End If
                             Call SP_OBJxSetFocus(ogdFunc)
                             Call W_PRCxAddPaySpcGiftVoucher
                        Else
                            Call SP_MSGnShowing("tms_CN005083", nCS_Error)
                            Call SP_OBJxSetFocus(ogdFunc)
                        End If
                
                        bVB_CNTenderAct = False
'*KT 53-01-06 ทำการย้ายอยู่ที่ Else เพราะหลังจากที่เรียกคำสั่ง Unload wSub แล้วเกิดมีการเรียก Timer ให้ Enable = True จะเข้า Form_Load อีกครั้ง
'                        otmForm.Enabled = True
'                        otmLock.Enabled = True
                        If cVB_Balance = 0 Then
'*KT 53-02-03 ป้องกันการ Terminated และใน Unload Me  มีการสั่ง otmForm.Enabled = False แล้ว
'                            otmForm.Enabled = False
'                            otmLock.Enabled = False
                            Call SP_DATxRsNothing(orsTemp)
                            Unload Me
                            Exit Sub
                        Else
                            otmForm.Enabled = True
                            otmLock.Enabled = True
                        End If
                    Case "T006" 'Vendor Coupon
                    '*Eaw 56-07-01 CommSheet TK-ISS3000-250-TR.xls
'                        If bVB_CNPayCrdNoSuccess Then
'                            MsgBox "Lock Tender", vbCritical
'                            bVB_CNTenderAct = False
'                            otmForm.Enabled = True
'                            otmLock.Enabled = True
'                            Call SP_DATxRsNothing(orsTemp)
'                            Call SP_OBJxSetFocus(ogdFunc)
'                            Exit Sub
'                        End If
                        bVB_CNTenderAct = True
                       ' nVB_CurTenderSeqNo = 5 '*KT 54-08-03 V.4.0.1 Req.2 Phase4.0(Member Point) Member Point Based on Tender
                        tVB_CNMnuCurrent = .TextMatrix(.Row, 2)
                        If aVB_TenderMnu(nVB_CurTenderSeqNo).Autho Then
                                If Not SP_USRbMnuAccept(nEN_Usr2AthRead, tVB_CNMnuCurrent) Then
                                    bVB_CNTenderAct = False
                                    otmForm.Enabled = True
                                    otmLock.Enabled = True
                                    Call SP_DATxRsNothing(orsTemp)
                                    Call SP_OBJxSetFocus(ogdFunc)
                                    Exit Sub
                                End If
                        End If
                        tVB_TenderCode = Trim(UCase(.TextMatrix(.Row, 3)))
                        tVB_TenderType = .TextMatrix(.Row, 5)
                        tVB_TenderName = Trim(.TextMatrix(.Row, 4))
                        
'                        wCNPayVendor.Show vbModal                       '*Ao 53-10-13
                        bVB_VedFrom = True
                        
                        '*Ao 2012/04/02 Communication Sheet ML-PRV-080  ให้สามารถชำระเป็นเศษสตางค์
                        '---------------------------------------------------------------------------------------------------
'                        Select Case Val(Format(tVB_TransactionSub))
'                            Case tEN_TrnDepositCancel, tEN_TrnAddDepositCancel
'                                olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance - cVB_RoundRcv), False, True) 'แสดง Balance ที่มีการบวกเศษด้วย เพราะการจ่ายด้วยเครดิตสามารถจ่ายเป็นเศษได้
'                            Case Else
'                                olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance - cVB_RoundRcv), False, True) 'แสดง Balance ที่มีการบวกเศษด้วย เพราะการจ่ายด้วยเครดิตสามารถจ่ายเป็นเศษได้
'                        End Select
                        olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance), False, True)  '*Eaw 56-01-29แสดง Balance ที่มีการปัดเศษ เพราะ cash ต้องมีการ Rouding
                        '---------------------------------------------------------------------------------------------------
                        wCNPayNewVendor.Show vbModal
                        '_____________________________________________________
                         If tVB_VendorCpnNo = "" Then
                            bVB_CNTenderAct = False
                            otmForm.Enabled = True
                            otmLock.Enabled = True
                            Call SP_DATxRsNothing(orsTemp)
                            Call SP_OBJxSetFocus(ogdFunc)
'                            Exit Sub           '*Ao 2012/04/02 Communication Sheet ML-PRV-080  ให้สามารถชำระเป็นเศษสตางค์
                        End If
                        '*Ao 2012/04/02 Communication Sheet ML-PRV-080  ให้สามารถชำระเป็นเศษสตางค์
                        If cVB_Pay = 0 Then
                            If cVB_Balance > 0 Then
                                '*Eaw 56-01-04 แก้ไขการปัด Rounding
'                                olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance), False, True) 'แสดง Balance ที่มีปัดเศษเหมือนเดิมเพราะไม่มีการชำระด้วยเครดิต
'                                Select Case Val(Format(tVB_TransactionSub))
'                                    Case tEN_TrnDepositCancel, tEN_TrnAddDepositCancel
'                                        olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance - cVB_RoundRcv), False, True) 'แสดง Balance ที่มีการบวกเศษด้วย เพราะการจ่ายด้วยเครดิตสามารถจ่ายเป็นเศษได้
'                                    Case Else
'                                        olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance - cVB_RoundRcv), False, True) 'แสดง Balance ที่มีการบวกเศษด้วย เพราะการจ่ายด้วยเครดิตสามารถจ่ายเป็นเศษได้
'                                End Select
                                cVB_Balance = cVB_Balance - cVB_RoundRcv
                                olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance), False, True)  '*Eaw 56-01-29แสดง Balance ที่มีการปัดเศษ เพราะ cash ต้องมีการ Rouding
                                cVB_RoundRcv = 0
                                '________________________________________
                            Else
                                Call SP_GETxBalanceRnd  '*Eaw 56-07-01 ทำเป็นฟังก์ชันกลาง
                                '*KT 51-11-12 อย่าลืมทำที่ Credit Store ด้วย
                                '---------------------------------------------------------------------------------------------------------
'                                tSql = "SELECT FNMenMin,FNMenValue FROM TSysMeaning"
'                                Call SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine, , , , , True, True) 'read master
'                                If Not (orsTemp.BOF Or orsTemp.EOF) Then
'                                    cMeaningDef = CDbl(Val(Format(SP_FEDvChkNumeric(orsTemp, "FNMenValue"))) / 100)
'                                Else
'                                    cMeaningDef = 0
'                                End If
'                                Call SP_DATxRsNothing(orsTemp)
                                
'                                If bVB_CNAlwRnd Then
'                                    Select Case nVB_CNTypeRnd
'                                        Case eRndUp
'                                            If SP_FMTcCurPoint(cVB_Balance - cVB_RoundRcv, True, False) < 1 Then
'                                                bVB_CNPayCrdNoSuccess = True
'                                                olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance - cVB_RoundRcv), False, True) 'แสดง Balance ที่มีการบวกเศษด้วย เพราะการจ่ายด้วยเครดิตสามารถจ่ายเป็นเศษได้
'                                            Else
'                                                bVB_CNPayCrdNoSuccess = False
'                                            End If
'                                        Case eRndDn
'                                            If SP_FMTcCurPoint(cVB_Balance - cVB_RoundRcv, True, False) < 1 Then
'                                                bVB_CNPayCrdNoSuccess = True
'                                                olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance - cVB_RoundRcv), False, True) 'แสดง Balance ที่มีการบวกเศษด้วย เพราะการจ่ายด้วยเครดิตสามารถจ่ายเป็นเศษได้
'                                            Else
'                                                bVB_CNPayCrdNoSuccess = False
'                                            End If
'                                        Case eRndMeaning
'                                            If SP_FMTcCurPoint(cVB_Balance - cVB_RoundRcv, True, False) <= cMeaningDef Then
'                                                bVB_CNPayCrdNoSuccess = True
'                                                olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance - cVB_RoundRcv), False, True) 'แสดง Balance ที่มีการบวกเศษด้วย เพราะการจ่ายด้วยเครดิตสามารถจ่ายเป็นเศษได้
'                                            Else
'                                                bVB_CNPayCrdNoSuccess = False
'                                            End If
'                                    End Select
'                                Else
'                                    If SP_FMTcCurPoint(cVB_Balance - cVB_RoundRcv, True, False) < 1 Then
'                                        bVB_CNPayCrdNoSuccess = True
'                                       olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance - cVB_RoundRcv), False, True) 'แสดง Balance ที่มีการบวกเศษด้วย เพราะการจ่ายด้วยเครดิตสามารถจ่ายเป็นเศษได้
'                                    Else
'                                        bVB_CNPayCrdNoSuccess = False
'                                    End If
'                                End If
                                '---------------------------------------------------------------------------------------------------------
                            End If
                            bVB_CNTenderAct = False
                            otmForm.Enabled = True
                            otmLock.Enabled = True
                            Call SP_DATxRsNothing(orsTemp)
                            Call SP_OBJxSetFocus(ogdFunc)
                            Exit Sub
                        End If
                        '_______________________________________________________________
                        Call SP_OBJxSetFocus(ogdFunc)
                        Call W_PRCxAddPayVendorCpn
                        bVB_CNTenderAct = False
'*KT 53-01-06 ทำการย้ายอยู่ที่ Else เพราะหลังจากที่เรียกคำสั่ง Unload wSub แล้วเกิดมีการเรียก Timer ให้ Enable = True จะเข้า Form_Load อีกครั้ง
'                        otmForm.Enabled = True
'                        otmLock.Enabled = True
                         If cVB_Balance = 0 Then
'*KT 53-02-03 ป้องกันการ Terminated และใน Unload Me  มีการสั่ง otmForm.Enabled = False แล้ว
'                            otmForm.Enabled = False
'                            otmLock.Enabled = False
                            Call SP_DATxRsNothing(orsTemp)
                            Unload Me
                            Exit Sub
                        Else
                            otmForm.Enabled = True
                            otmLock.Enabled = True
                        End If
                    Case "T007" 'Special Coupon
                        '*KT 67-01-03 CR Mobile Voucher
                        If .TextMatrix(.Row, 5) <> "Special Coupon" Then
                            tVB_CNMnuCurrent = .TextMatrix(.Row, 2)
                            If aVB_TenderMnu(nVB_CurTenderSeqNo).Autho Then
                                If Not SP_USRbMnuAccept(nEN_Usr2AthRead, tVB_CNMnuCurrent) Then
                                    bVB_CNTenderAct = False
                                    otmForm.Enabled = True
                                    otmLock.Enabled = True
                                    Call SP_DATxRsNothing(orsTemp)
                                    Call SP_OBJxSetFocus(ogdFunc)
                                    Exit Sub
                                End If
                             End If
                        
                            olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance), False, True)
                            bVB_CNTenderAct = True
                            tVB_TenderCode = "T002"
                            tVB_TenderType = "Credit Card" '.TextMatrix(.Row, 5)
                            tVB_TenderName = Trim(.TextMatrix(.Row, 4))
                            tVB_CrdStaBPSale = IIf(aVB_TenderMnu(nVB_CurTenderSeqNo).StaBPSale, "Y", "N") '"N"
                            tVB_CrdStaBPReturn = IIf(aVB_TenderMnu(nVB_CurTenderSeqNo).StaBPReturn, "Y", "N") '"N"
                            tVB_CrdEvnID = ""
                            tVB_EDCDetail = ""
                            tVB_CreditGroup = ""
                            tVB_CrdType = ""
                            
                            wCNPayMVoucher.Show vbModal
                            If cVB_Pay = 0 Then
                                cVB_Balance = cVB_Balance - cVB_RoundRcv
                                olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance), False, True)  '*Eaw 56-01-29แสดง Balance ที่มีการปัดเศษ เพราะ cash ต้องมีการ Rouding
                                cVB_RoundRcv = 0
                                '________________________________________
                                bVB_CNTenderAct = False
                                otmForm.Enabled = True
                                otmLock.Enabled = True
                                Call SP_DATxRsNothing(orsTemp)
                                Call SP_OBJxSetFocus(ogdFunc)
                                Exit Sub
                            End If
                            Call SP_OBJxSetFocus(ogdFunc)
                            Call W_PRCxAddPayCrdCard
                        Else
                            '*Eaw 56-01-04 แก้ไขการปัด Rounding
    '                        Select Case Val(Format(tVB_TransactionSub))
    '                            Case tEN_TrnDepositCancel, tEN_TrnAddDepositCancel
    '                                olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance - cVB_RoundRcv), False, True) 'แสดง Balance ที่มีการบวกเศษด้วย เพราะการจ่ายด้วยเครดิตสามารถจ่ายเป็นเศษได้
    '                            Case Else
    '                                olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance - cVB_RoundRcv), False, True) 'แสดง Balance ที่มีการบวกเศษด้วย เพราะการจ่ายด้วยเครดิตสามารถจ่ายเป็นเศษได้
    '                        End Select
                            olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance), False, True)  '*Eaw 56-01-29แสดง Balance ที่มีการปัดเศษ เพราะ cash ต้องมีการ Rouding
                            '*Eaw 56-07-01 CommSheet TK-ISS3000-250-TR.xls
    '                        If bVB_CNPayCrdNoSuccess Then
    '                            MsgBox "Lock Tender", vbCritical
    '                            bVB_CNTenderAct = False
    '                            otmForm.Enabled = True
    '                            otmLock.Enabled = True
    '                            Call SP_DATxRsNothing(orsTemp)
    '                            Call SP_OBJxSetFocus(ogdFunc)
    '                            Exit Sub
    '                        End If
                            bVB_CNTenderAct = True
                            'ก่อนการชำระเงินถ้ามีการชำระประเภทอื่นที่ไม่ใช่Spacial Gift Voucher/Special Coupon จะไม่ยอมให้ชำระเงินประเภทนี้อีกแล้ว
                            'จะยกเลิกหน้านี้แล้วเข้ามาใหม่
                            tWhe = " WHERE  FTTmnNum='" & tVB_CNTerminalNum & "' AND FTShdTransNo='" & tVB_Running & "'"
                            tWhe = tWhe & " AND  FTShdTransType='" & tVB_Transaction & "' AND FDShdTransDate='" & tUT_SaleDate & "'"
                            tWhe = tWhe & " AND FTTdmCode NOT IN ('T005','T007','T014') AND NOT (FTTdmCode IN ('T006') AND ISNULL(FTVendorCnpFlag,'') = '1')"   '*Ao 2011/01/26 CommSheet ML-UAT-PRV-032-Special Gift Voucher can't use with Vendor Coupon (item).xls
    '                        If SP_VALcText2Double(SP_SQLtFunction("SUM", "FCSrcNet", tVB_CNTblRC, tWhe)) = 0 Then       '*Ao 2013/03/19 CommSheet TK-ISS3000-109 00.00.01.xlsx
                            If bVB_CNTenderAct Then
                                'nVB_CurTenderSeqNo = 6 '*KT 54-08-03 V.4.0.1 Req.2 Phase4.0(Member Point) Member Point Based on Tender
                                tVB_CNMnuCurrent = .TextMatrix(.Row, 2)
                                If aVB_TenderMnu(nVB_CurTenderSeqNo).Autho Then
                                    If Not SP_USRbMnuAccept(nEN_Usr2AthRead, tVB_CNMnuCurrent) Then
                                        bVB_CNTenderAct = False
                                        otmForm.Enabled = True
                                        otmLock.Enabled = True
                                        Call SP_DATxRsNothing(orsTemp)
                                        Call SP_OBJxSetFocus(ogdFunc)
                                        Exit Sub
                                    End If
                                 End If
                                tVB_TenderCode = Trim(UCase(.TextMatrix(.Row, 3)))
                                tVB_TenderType = .TextMatrix(.Row, 5)
                                tVB_TenderName = Trim(.TextMatrix(.Row, 4))
    '                            wCNPayOth.Caption = aVB_TenderMnu(nVB_CurTenderSeqNo).Menu   '*Eaw 55-11-29
    '                            wCNPayOth.tW_Name = "Special Cpn No."
                                tVB_Tender = "Special Cpn No."         '*Ao 2013/03/20  TK
    '                            wCNPayOth.Caption = "Special Coupon"
                                wCNPayOth.Show vbModal
                                 If cVB_Pay = 0 Then
                                    '*Eaw 56-01-04 แก้ไขการปัด Rounding
    '                                Select Case Val(Format(tVB_TransactionSub))
    '                                    Case tEN_TrnDepositCancel, tEN_TrnAddDepositCancel
    '                                        olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance - cVB_RoundRcv), False, True) 'แสดง Balance ที่มีการบวกเศษด้วย เพราะการจ่ายด้วยเครดิตสามารถจ่ายเป็นเศษได้
    '                                    Case Else
    '                                        olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance - cVB_RoundRcv), False, True) 'แสดง Balance ที่มีการบวกเศษด้วย เพราะการจ่ายด้วยเครดิตสามารถจ่ายเป็นเศษได้
    '                                End Select
                                    cVB_Balance = cVB_Balance - cVB_RoundRcv
                                    olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance), False, True)  '*Eaw 56-01-29แสดง Balance ที่มีการปัดเศษ เพราะ cash ต้องมีการ Rouding
                                    cVB_RoundRcv = 0
                                    '________________________________________
                                    bVB_CNTenderAct = False
                                    otmForm.Enabled = True
                                    otmLock.Enabled = True
                                    Call SP_DATxRsNothing(orsTemp)
                                    Call SP_OBJxSetFocus(ogdFunc)
                                    Exit Sub
                                End If
                                 Call SP_OBJxSetFocus(ogdFunc)
                                Call W_PRCxAddPaySpecialCpn
                            Else
                                Call SP_MSGnShowing("tms_CN005084", nCS_Error)
                                Call SP_OBJxSetFocus(ogdFunc)
                            End If
                        End If
                        bVB_CNTenderAct = False
'*KT 53-01-06 ทำการย้ายอยู่ที่ Else เพราะหลังจากที่เรียกคำสั่ง Unload wSub แล้วเกิดมีการเรียก Timer ให้ Enable = True จะเข้า Form_Load อีกครั้ง
'                        otmForm.Enabled = True
'                        otmLock.Enabled = True
                        If cVB_Balance = 0 Then
'*KT 53-02-03 ป้องกันการ Terminated และใน Unload Me  มีการสั่ง otmForm.Enabled = False แล้ว
'                            otmForm.Enabled = False
'                            otmLock.Enabled = False
                            Call SP_DATxRsNothing(orsTemp)
                            Unload Me
                            Exit Sub
                        Else
                            otmForm.Enabled = True
                            otmLock.Enabled = True
                        End If
                    Case "T008" 'Credit Store
                         bVB_CNTenderAct = True
                        'nVB_CurTenderSeqNo = 7 '*KT 54-08-03 V.4.0.1 Req.2 Phase4.0(Member Point) Member Point Based on Tender
                        tVB_CNMnuCurrent = .TextMatrix(.Row, 2)
                        If aVB_TenderMnu(nVB_CurTenderSeqNo).Autho Then
                                If Not SP_USRbMnuAccept(nEN_Usr2AthRead, tVB_CNMnuCurrent) Then
                                    bVB_CNTenderAct = False
                                    otmForm.Enabled = True
                                    otmLock.Enabled = True
                                    Call SP_DATxRsNothing(orsTemp)
                                    Call SP_OBJxSetFocus(ogdFunc)
                                    Exit Sub
                                End If
                         End If
                        tVB_TenderCode = Trim(UCase(.TextMatrix(.Row, 3)))
                        tVB_TenderType = .TextMatrix(.Row, 5)
                        tVB_TenderName = Trim(.TextMatrix(.Row, 4))
                        '*KT 51-09-2008 ให้สามารถชำระเป็นเศษสตางค์
                        '---------------------------------------------------------------------------------------------------
'                        Select Case Val(Format(tVB_TransactionSub))
'                            Case tEN_TrnDepositCancel, tEN_TrnAddDepositCancel
'                                olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance - cVB_RoundRcv), False, True) 'แสดง Balance ที่มีการบวกเศษด้วย เพราะการจ่ายด้วยเครดิตสามารถจ่ายเป็นเศษได้
'                            Case Else
'                                olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance - cVB_RoundRcv), False, True) 'แสดง Balance ที่มีการบวกเศษด้วย เพราะการจ่ายด้วยเครดิตสามารถจ่ายเป็นเศษได้
'                        End Select
                        olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance), False, True)  '*Eaw 56-01-29แสดง Balance ที่มีการปัดเศษ เพราะ cash ต้องมีการ Rouding
'                        DoEvents '*KT 53-01-14 DoEvents
                        '----------------------------------------------------------------------------------------------------
'                        wCNPayOth.tW_Name = "Credit Store"      '*Eaw 55-11-27
'                        wCNPayOth.Caption = "Credit Store"          '*Eaw 55-11-27
'                        wCNPayOth.tW_Name = "Account No. :"   '*Eaw 55-11-27
                        tVB_Tender = "Account No. :"        '*Ao 2013/03/20  TK
                        wCNPayOth.Show vbModal
                        If cVB_Pay = 0 Then
                            If cVB_Balance > 0 Then
                                '*Eaw 56-01-04 แก้ไขการปัด Rounding
'                                Select Case Val(Format(tVB_TransactionSub))
'                                    Case tEN_TrnDepositCancel, tEN_TrnAddDepositCancel
'                                        olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance - cVB_RoundRcv), False, True) 'แสดง Balance ที่มีการบวกเศษด้วย เพราะการจ่ายด้วยเครดิตสามารถจ่ายเป็นเศษได้
'                                    Case Else
'                                        olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance - cVB_RoundRcv), False, True) 'แสดง Balance ที่มีการบวกเศษด้วย เพราะการจ่ายด้วยเครดิตสามารถจ่ายเป็นเศษได้
'                                End Select
                                olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance), False, True)  '*Eaw 56-01-29แสดง Balance ที่มีการปัดเศษ เพราะ cash ต้องมีการ Rouding
'                                olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance), False, True) 'แสดง Balance ที่มีปัดเศษเหมือนเดิมเพราะไม่มีการชำระด้วยเครดิต
                            Else
                                Call SP_GETxBalanceRnd  '*Eaw 56-07-01 ทำเป็นฟังก์ชันกลาง
                                '*KT 51-11-12 อย่าลืมทำที่ Credit Store ด้วย
                                '---------------------------------------------------------------------------------------------------------
'                                tSql = "SELECT FNMenMin,FNMenValue FROM TSysMeaning"
'                                Call SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine, , , , , True, True) 'read master
'                                If Not (orsTemp.BOF Or orsTemp.EOF) Then
'                                    cMeaningDef = CDbl(Val(Format(SP_FEDvChkNumeric(orsTemp, "FNMenValue"))) / 100)
'                                Else
'                                    cMeaningDef = 0
'                                End If
'                                Call SP_DATxRsNothing(orsTemp)
'
'                                If bVB_CNAlwRnd Then
'                                    Select Case nVB_CNTypeRnd
'                                        Case eRndUp
'                                            If SP_FMTcCurPoint(cVB_Balance - cVB_RoundRcv, True, False) < 1 Then
'                                                bVB_CNPayCrdNoSuccess = True
'                                                olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance - cVB_RoundRcv), False, True) 'แสดง Balance ที่มีการบวกเศษด้วย เพราะการจ่ายด้วยเครดิตสามารถจ่ายเป็นเศษได้
'                                            Else
'                                                bVB_CNPayCrdNoSuccess = False
'                                            End If
'                                        Case eRndDn
'                                            If SP_FMTcCurPoint(cVB_Balance - cVB_RoundRcv, True, False) < 1 Then
'                                                bVB_CNPayCrdNoSuccess = True
'                                                olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance - cVB_RoundRcv), False, True) 'แสดง Balance ที่มีการบวกเศษด้วย เพราะการจ่ายด้วยเครดิตสามารถจ่ายเป็นเศษได้
'                                            Else
'                                                bVB_CNPayCrdNoSuccess = False
'                                            End If
'                                        Case eRndMeaning
'                                            If SP_FMTcCurPoint(cVB_Balance - cVB_RoundRcv, True, False) <= cMeaningDef Then
'                                                bVB_CNPayCrdNoSuccess = True
'                                                olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance - cVB_RoundRcv), False, True) 'แสดง Balance ที่มีการบวกเศษด้วย เพราะการจ่ายด้วยเครดิตสามารถจ่ายเป็นเศษได้
'                                            Else
'                                                bVB_CNPayCrdNoSuccess = False
'                                            End If
'                                    End Select
'                                Else
'                                    If SP_FMTcCurPoint(cVB_Balance - cVB_RoundRcv, True, False) < 1 Then
'                                        bVB_CNPayCrdNoSuccess = True
'                                       olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance - cVB_RoundRcv), False, True) 'แสดง Balance ที่มีการบวกเศษด้วย เพราะการจ่ายด้วยเครดิตสามารถจ่ายเป็นเศษได้
'                                    Else
'                                        bVB_CNPayCrdNoSuccess = False
'                                    End If
'                                End If
                                '---------------------------------------------------------------------------------------------------------
                            End If
                            bVB_CNTenderAct = False
                            otmForm.Enabled = True
                            otmLock.Enabled = True
                            Call SP_DATxRsNothing(orsTemp)
                            Call SP_OBJxSetFocus(ogdFunc)
                            Exit Sub
                        End If
                        Call SP_OBJxSetFocus(ogdFunc)
                        Call W_PRCxAddPayCrdStore
                        bVB_CNTenderAct = False
'*KT 53-01-06 ทำการย้ายอยู่ที่ Else เพราะหลังจากที่เรียกคำสั่ง Unload wSub แล้วเกิดมีการเรียก Timer ให้ Enable = True จะเข้า Form_Load อีกครั้ง
'                        otmForm.Enabled = True
'                        otmLock.Enabled = True
                        If cVB_Balance = 0 Then
'*KT 53-02-03 ป้องกันการ Terminated และใน Unload Me  มีการสั่ง otmForm.Enabled = False แล้ว
'                            otmForm.Enabled = False
'                            otmLock.Enabled = False
                            Call SP_DATxRsNothing(orsTemp)
                            Unload Me
                            Exit Sub
                        Else
                            otmForm.Enabled = True
                            otmLock.Enabled = True
                        End If
                    Case "T009" ' 'GIFT SHOCK
                    '*Eaw 56-01-04 แก้ไขการปัด Rounding
'                        Select Case Val(Format(tVB_TransactionSub))
'                            Case tEN_TrnDepositCancel, tEN_TrnAddDepositCancel
'                                olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance - cVB_RoundRcv), False, True) 'แสดง Balance ที่มีการบวกเศษด้วย เพราะการจ่ายด้วยเครดิตสามารถจ่ายเป็นเศษได้
'                            Case Else
'                                olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance - cVB_RoundRcv), False, True) 'แสดง Balance ที่มีการบวกเศษด้วย เพราะการจ่ายด้วยเครดิตสามารถจ่ายเป็นเศษได้
'                        End Select
                        olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance), False, True)  '*Eaw 56-01-29แสดง Balance ที่มีการปัดเศษ เพราะ cash ต้องมีการ Rouding
                        '--------------------------------------------------------------------------------------------------
                        '*Eaw 56-07-01 CommSheet TK-ISS3000-250-TR.xls
'                        If bVB_CNPayCrdNoSuccess Then
'                            MsgBox "Lock Tender", vbCritical
'                            bVB_CNTenderAct = False
'                            otmForm.Enabled = True
'                            otmLock.Enabled = True
'                            Call SP_DATxRsNothing(orsTemp)
'                            Call SP_OBJxSetFocus(ogdFunc)
'                            Exit Sub
'                        End If
                        bVB_CNTenderAct = True
                        'nVB_CurTenderSeqNo = 8 '*KT 54-08-03 V.4.0.1 Req.2 Phase4.0(Member Point) Member Point Based on Tender
                        tVB_CNMnuCurrent = .TextMatrix(.Row, 2)
                        If aVB_TenderMnu(nVB_CurTenderSeqNo).Autho Then
                            If Not SP_USRbMnuAccept(nEN_Usr2AthRead, tVB_CNMnuCurrent) Then
                                bVB_CNTenderAct = False
                                otmForm.Enabled = True
                                otmLock.Enabled = True
                                Call SP_DATxRsNothing(orsTemp)
                                Call SP_OBJxSetFocus(ogdFunc)
                                Exit Sub
                            End If
                        End If
                        tVB_TenderCode = Trim(UCase(.TextMatrix(.Row, 3)))
                        tVB_TenderType = .TextMatrix(.Row, 5)
                        tVB_TenderName = Trim(.TextMatrix(.Row, 4))
'                        wCNPayOth.tW_Name = "Store Cpn No"     '*Ao 2013/03/20  TK
                        tVB_Tender = "Store Cpn No"
'                        wCNPayOth.Caption = "Store Coupon"
                        wCNPayOth.Show vbModal
                         If cVB_Pay = 0 Then
                            '*Eaw 56-01-04 แก้ไขการปัด Rounding
'                            Select Case Val(Format(tVB_TransactionSub))
'                                Case tEN_TrnDepositCancel, tEN_TrnAddDepositCancel
'                                    olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance - cVB_RoundRcv), False, True) 'แสดง Balance ที่มีการบวกเศษด้วย เพราะการจ่ายด้วยเครดิตสามารถจ่ายเป็นเศษได้
'                                Case Else
'                                    olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance - cVB_RoundRcv), False, True) 'แสดง Balance ที่มีการบวกเศษด้วย เพราะการจ่ายด้วยเครดิตสามารถจ่ายเป็นเศษได้
'                            End Select
                            cVB_Balance = cVB_Balance - cVB_RoundRcv
                            olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance), False, True)  '*Eaw 56-01-29แสดง Balance ที่มีการปัดเศษ เพราะ cash ต้องมีการ Rouding
                            cVB_RoundRcv = 0
                            '________________________________________
                            bVB_CNTenderAct = False
                            otmForm.Enabled = True
                            otmLock.Enabled = True
                            Call SP_DATxRsNothing(orsTemp)
                            Call SP_OBJxSetFocus(ogdFunc)
                            Exit Sub
                        End If
                         Call SP_OBJxSetFocus(ogdFunc)
                         Call W_PRCxAddPayStoreCoupon
                         bVB_CNTenderAct = False
'*KT 53-01-06 ทำการย้ายอยู่ที่ Else เพราะหลังจากที่เรียกคำสั่ง Unload wSub แล้วเกิดมีการเรียก Timer ให้ Enable = True จะเข้า Form_Load อีกครั้ง
'                         otmForm.Enabled = True
'                         otmLock.Enabled = True
                        If cVB_Balance = 0 Then
'*KT 53-02-03 ป้องกันการ Terminated และใน Unload Me  มีการสั่ง otmForm.Enabled = False แล้ว
'                            otmForm.Enabled = False
'                            otmLock.Enabled = False
                            Call SP_DATxRsNothing(orsTemp)
                            Unload Me
                            Exit Sub
                        Else
                            otmForm.Enabled = True
                            otmLock.Enabled = True
                        End If
                    Case "T010" 'VOUCHER
                    '*Eaw 56-01-04 แก้ไขการปัด Rounding
'                        Select Case Val(Format(tVB_TransactionSub))
'                            Case tEN_TrnDepositCancel, tEN_TrnAddDepositCancel
'                                olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance - cVB_RoundRcv), False, True) 'แสดง Balance ที่มีการบวกเศษด้วย เพราะการจ่ายด้วยเครดิตสามารถจ่ายเป็นเศษได้
'                            Case Else
'                                olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance - cVB_RoundRcv), False, True) 'แสดง Balance ที่มีการบวกเศษด้วย เพราะการจ่ายด้วยเครดิตสามารถจ่ายเป็นเศษได้
'                        End Select
                        olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance), False, True)  '*Eaw 56-01-29แสดง Balance ที่มีการปัดเศษ เพราะ cash ต้องมีการ Rouding
                        '---------------------------------------------------------------------------------
                        '*Eaw 56-07-01 CommSheet TK-ISS3000-250-TR.xls
'                        If bVB_CNPayCrdNoSuccess Then
'                            MsgBox "Lock Tender", vbCritical
'                            bVB_CNTenderAct = False
'                            otmForm.Enabled = True
'                            otmLock.Enabled = True
'                            Call SP_DATxRsNothing(orsTemp)
'                            Call SP_OBJxSetFocus(ogdFunc)
'                            Exit Sub
'                        End If
                        bVB_CNTenderAct = True
                        'nVB_CurTenderSeqNo = 9 '*KT 54-08-03 V.4.0.1 Req.2 Phase4.0(Member Point) Member Point Based on Tender
                        tVB_CNMnuCurrent = .TextMatrix(.Row, 2)
                        If aVB_TenderMnu(nVB_CurTenderSeqNo).Autho Then
                                If Not SP_USRbMnuAccept(nEN_Usr2AthRead, tVB_CNMnuCurrent) Then
                                    bVB_CNTenderAct = False
                                    otmForm.Enabled = True
                                    otmLock.Enabled = True
                                    Call SP_DATxRsNothing(orsTemp)
                                    Call SP_OBJxSetFocus(ogdFunc) 'ogdFunc.SetFocus
                                    Exit Sub
                                End If
                         End If
                        tVB_TenderCode = Trim(UCase(.TextMatrix(.Row, 3)))
                        tVB_TenderType = .TextMatrix(.Row, 5)
                        tVB_TenderName = Trim(.TextMatrix(.Row, 4))
'                        wCNPayOth.tW_Name = "Voucher No"
'                        wCNPayOth.Caption = "Voucher"
                        tVB_Tender = "Voucher No"
                        wCNPayOth.Show vbModal
                        If cVB_Pay = 0 Then
                            '*Eaw 56-01-04 แก้ไขการปัด Rounding
'                            Select Case Val(Format(tVB_TransactionSub))
'                                Case tEN_TrnDepositCancel, tEN_TrnAddDepositCancel
'                                    olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance - cVB_RoundRcv), False, True) 'แสดง Balance ที่มีการบวกเศษด้วย เพราะการจ่ายด้วยเครดิตสามารถจ่ายเป็นเศษได้
'                                Case Else
'                                    olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance - cVB_RoundRcv), False, True) 'แสดง Balance ที่มีการบวกเศษด้วย เพราะการจ่ายด้วยเครดิตสามารถจ่ายเป็นเศษได้
'                            End Select
                            cVB_Balance = cVB_Balance - cVB_RoundRcv
                            olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance), False, True)  '*Eaw 56-01-29แสดง Balance ที่มีการปัดเศษ เพราะ cash ต้องมีการ Rouding
                            cVB_RoundRcv = 0
                            '________________________________________
                            bVB_CNTenderAct = False
                            otmForm.Enabled = True
                            otmLock.Enabled = True
                            Call SP_DATxRsNothing(orsTemp)
                            Call SP_OBJxSetFocus(ogdFunc)
                            Exit Sub
                        End If
                        Call SP_OBJxSetFocus(ogdFunc)
                        Call W_PRCxAddPayVoucher
                        bVB_CNTenderAct = False
'*KT 53-01-06 ทำการย้ายอยู่ที่ Else เพราะหลังจากที่เรียกคำสั่ง Unload wSub แล้วเกิดมีการเรียก Timer ให้ Enable = True จะเข้า Form_Load อีกครั้ง
'                        otmForm.Enabled = True
'                        otmLock.Enabled = True
                        If cVB_Balance = 0 Then
'*KT 53-02-03 ป้องกันการ Terminated และใน Unload Me  มีการสั่ง otmForm.Enabled = False แล้ว
'                            otmForm.Enabled = False
'                            otmLock.Enabled = False
                            Call SP_DATxRsNothing(orsTemp)
                            Unload Me
                            Exit Sub
                        Else
                            otmForm.Enabled = True
                            otmLock.Enabled = True
                        End If
                    Case "T011" 'CREDIT NOTE
                        '*Eaw 56-01-04 แก้ไขการปัด Rounding
'                        Select Case Val(Format(tVB_TransactionSub))
'                            Case tEN_TrnDepositCancel, tEN_TrnAddDepositCancel
'                                olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance - cVB_RoundRcv), False, True) 'แสดง Balance ที่มีการบวกเศษด้วย เพราะการจ่ายด้วยเครดิตสามารถจ่ายเป็นเศษได้
'                            Case Else
'                                olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance - cVB_RoundRcv), False, True) 'แสดง Balance ที่มีการบวกเศษด้วย เพราะการจ่ายด้วยเครดิตสามารถจ่ายเป็นเศษได้
'                        End Select
                        olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance), False, True)  '*Eaw 56-01-29แสดง Balance ที่มีการปัดเศษ เพราะ cash ต้องมีการ Rouding
                        '-----------------------------------------------------------------------------------------------
                        '*Eaw 56-07-01 CommSheet TK-ISS3000-250-TR.xls
'                        If bVB_CNPayCrdNoSuccess Then
'                            MsgBox "Lock Tender", vbCritical
'                            bVB_CNTenderAct = False
'                            otmForm.Enabled = True
'                            otmLock.Enabled = True
'                            Call SP_DATxRsNothing(orsTemp)
'                            Call SP_OBJxSetFocus(ogdFunc)
'                            Exit Sub
'                        End If
                        bVB_CNTenderAct = True
                        'nVB_CurTenderSeqNo = 10 '*KT 54-08-03 V.4.0.1 Req.2 Phase4.0(Member Point) Member Point Based on Tender
                        tVB_CNMnuCurrent = .TextMatrix(.Row, 2)
                        If aVB_TenderMnu(nVB_CurTenderSeqNo).Autho Then
                                If Not SP_USRbMnuAccept(nEN_Usr2AthRead, tVB_CNMnuCurrent) Then
                                    bVB_CNTenderAct = False
                                    otmForm.Enabled = True
                                    otmLock.Enabled = True
                                    Call SP_DATxRsNothing(orsTemp)
                                    Call SP_OBJxSetFocus(ogdFunc)   'ogdFunc.SetFocus
                                    Exit Sub
                                End If
                         End If
                        tVB_TenderCode = Trim(UCase(.TextMatrix(.Row, 3)))
                        tVB_TenderType = .TextMatrix(.Row, 5)
                        tVB_TenderName = Trim(.TextMatrix(.Row, 4))
'                        wCNPayOth.tW_Name = "Credit Not No"
                        tVB_Tender = "Credit Note No"                   '*Ao 2013/03/20  TK
'                        wCNPayOth.Caption = "Credit Note"
                        wCNPayOth.Show vbModal
                        If cVB_Pay = 0 Then
                            '*Eaw 56-01-04 แก้ไขการปัด Rounding
'                            Select Case Val(Format(tVB_TransactionSub))
'                                Case tEN_TrnDepositCancel, tEN_TrnAddDepositCancel
'                                    olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance - cVB_RoundRcv), False, True) 'แสดง Balance ที่มีการบวกเศษด้วย เพราะการจ่ายด้วยเครดิตสามารถจ่ายเป็นเศษได้
'                                Case Else
'                                    olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance - cVB_RoundRcv), False, True) 'แสดง Balance ที่มีการบวกเศษด้วย เพราะการจ่ายด้วยเครดิตสามารถจ่ายเป็นเศษได้
'                            End Select
                            cVB_Balance = cVB_Balance - cVB_RoundRcv
                            olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance), False, True)  '*Eaw 56-01-29แสดง Balance ที่มีการปัดเศษ เพราะ cash ต้องมีการ Rouding
                            cVB_RoundRcv = 0
                            '________________________________________
                            bVB_CNTenderAct = False
                            otmForm.Enabled = True
                            otmLock.Enabled = True
                            Call SP_DATxRsNothing(orsTemp)
                            Call SP_OBJxSetFocus(ogdFunc)
                            Exit Sub
                        End If
                        Call SP_OBJxSetFocus(ogdFunc)
                        Call W_PRCxAddPayCreditNote
                        bVB_CNTenderAct = False
'*KT 53-01-06 ทำการย้ายอยู่ที่ Else เพราะหลังจากที่เรียกคำสั่ง Unload wSub แล้วเกิดมีการเรียก Timer ให้ Enable = True จะเข้า Form_Load อีกครั้ง
'                        otmForm.Enabled = True
'                        otmLock.Enabled = True
                        If cVB_Balance = 0 Then
'*KT 53-02-03 ป้องกันการ Terminated และใน Unload Me  มีการสั่ง otmForm.Enabled = False แล้ว
'                            otmForm.Enabled = False
'                            otmLock.Enabled = False
                            Call SP_DATxRsNothing(orsTemp)
                            Unload Me
                            Exit Sub
                        Else
                            otmForm.Enabled = True
                            otmLock.Enabled = True
                        End If
                    
                    Case "T012" 'FOREIGN CURRENCY
                        '*Eaw 56-03-19 CommSheet TK-ISS3000-105 00.00.01.xlsx ไม่ต้อง Rounding ก่อนเข้า
'                        '*Eaw 56-01-03 ก่อนเข้า  Tender ไม่มีการ Rounding เมื่อต้องการชำระด้วย cash ให้ใช้ค่า Balance+Rounding เป็นยอดที่ต้องชำระ
'                        cVB_RoundRcv = SP_FMTcCurPoint(IIf(bVB_CNAlwRnd, SP_FMTcCurPoint(UT_GetnRoundDiff(cVB_Balance), True, True), 0), True, False) '*KT 51-04-05 Comsheet ML-SF-172,ML-SF-173
'                        cVB_Balance = cVB_Balance + cVB_RoundRcv
''                        Select Case Val(Format(tVB_TransactionSub))
''                            Case tEN_TrnDepositCancel, tEN_TrnAddDepositCancel
''                                olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance), False, True)  'แสดง Balance ที่มีการปัดเศษ เพราะ cash ต้องมีการ Rouding
''                            Case Else
''                                olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance), False, True)  'แสดง Balance ที่มีการปัดเศษ เพราะ cash ต้องมีการ Rouding
''                        End Select
                        olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance), False, True)  '*Eaw 56-01-29แสดง Balance ที่มีการปัดเศษ เพราะ cash ต้องมีการ Rouding
                        '--------------------------------------------------------------------------------------
                        '*Eaw 56-07-01 CommSheet TK-ISS3000-250-TR.xls
'                        If bVB_CNPayCrdNoSuccess Then
'                            MsgBox "Lock Tender", vbCritical
'                            bVB_CNTenderAct = False
'                            otmForm.Enabled = True
'                            otmLock.Enabled = True
'                            Call SP_DATxRsNothing(orsTemp)
'                            Call SP_OBJxSetFocus(ogdFunc)
'                            Exit Sub
'                        End If
                        bVB_CNTenderAct = True
                        'nVB_CurTenderSeqNo = 11 '*KT 54-08-03 V.4.0.1 Req.2 Phase4.0(Member Point) Member Point Based on Tender
                        tVB_CNMnuCurrent = .TextMatrix(.Row, 2)
                        If aVB_TenderMnu(nVB_CurTenderSeqNo).Autho Then
                                If Not SP_USRbMnuAccept(nEN_Usr2AthRead, tVB_CNMnuCurrent) Then
                                    bVB_CNTenderAct = False
                                    otmForm.Enabled = True
                                    otmLock.Enabled = True
                                    Call SP_DATxRsNothing(orsTemp)
                                    Call SP_OBJxSetFocus(ogdFunc)
                                    Exit Sub
                                End If
                         End If
                        tVB_TenderCode = Trim(UCase(.TextMatrix(.Row, 3)))
                        tVB_TenderType = .TextMatrix(.Row, 5)
                        tVB_TenderName = Trim(.TextMatrix(.Row, 4))
                        wCNCurrency.Show vbModal
                        If cVB_Pay = 0 Then
                            '*Eaw 56-01-03 ก่อนเข้า  Tender ไม่มีการ Rounding เมื่อต้องการชำระด้วย cash ให้ใช้ค่า Balance+Rounding เป็นยอดที่ต้องชำระ
'                            cVB_RoundRcv = SP_FMTcCurPoint(IIf(bVB_CNAlwRnd, SP_FMTcCurPoint(UT_GetnRoundDiff(cVB_Balance), True, True), 0), True, False) '*KT 51-04-05 Comsheet ML-SF-172,ML-SF-173
'                            cVB_Balance = cVB_Balance - cVB_RoundRcv '*Eaw 56-03-19 CommSheet TK-ISS3000-105 00.00.01.xlsx ไม่ต้อง Rounding ก่อนเข้า
'                            Select Case Val(Format(tVB_TransactionSub))
'                                Case tEN_TrnDepositCancel, tEN_TrnAddDepositCancel
'                                    olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance), False, True) 'แสดง Balance ที่มีการปัดเศษ เพราะ cash ต้องมีการ Rouding
'                                Case Else
'                                    olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance), False, True) 'แสดง Balance ที่มีการปัดเศษ เพราะ cash ต้องมีการ Rouding
'                            End Select
                            olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance), False, True)  '*Eaw 56-01-29  แสดง Balance ที่มีการปัดเศษ เพราะ cash ต้องมีการ Rouding
'                            cVB_RoundRcv = 0  '*Eaw 56-01-03 '*Eaw 56-03-19 CommSheet TK-ISS3000-105 00.00.01.xlsx ไม่ต้อง Rounding ก่อนเข้า
                            '----------------------------------------------------------------------------------------------
                            bVB_CNTenderAct = False
                            otmForm.Enabled = True
                            otmLock.Enabled = True
                            Call SP_DATxRsNothing(orsTemp)
                            Call SP_OBJxSetFocus(ogdFunc)
                            Exit Sub
                        End If
                        Call SP_OBJxSetFocus(ogdFunc)
                        tVB_TenderName = tVB_ForeignCurName
                        Call W_PRCxAddPayForeignCur
                        bVB_CNTenderAct = False
'*KT 53-01-06 ทำการย้ายอยู่ที่ Else เพราะหลังจากที่เรียกคำสั่ง Unload wSub แล้วเกิดมีการเรียก Timer ให้ Enable = True จะเข้า Form_Load อีกครั้ง
'                        otmForm.Enabled = True
'                        otmLock.Enabled = True
                        If cVB_Balance = 0 Then
'*KT 53-02-03 ป้องกันการ Terminated และใน Unload Me  มีการสั่ง otmForm.Enabled = False แล้ว
'                            otmForm.Enabled = False
'                            otmLock.Enabled = False
                            Call SP_DATxRsNothing(orsTemp)
                            Unload Me
                            Exit Sub
                        Else
                            otmForm.Enabled = True
                            otmLock.Enabled = True
                        End If
                Case "T013" 'CASH CARD   '*Eaw 55-11-27 Add for TK copy from case T015
                         tVB_TypeOfEDC = "4"
                        bVB_CNTenderAct = True
                        '*Eaw 56-01-03 ก่อนเข้า  Tender ไม่มีการ Rounding ใช้ค่า Balance ที่มีอยู่ได้เลย
'                        Select Case Val(Format(tVB_TransactionSub))
'                            Case tEN_TrnDepositCancel, tEN_TrnAddDepositCancel
'                                olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance - cVB_RoundRcv), False, True) 'แสดง Balance ที่มีการบวกเศษด้วย เพราะการจ่ายด้วยเครดิตสามารถจ่ายเป็นเศษได้
'                            Case Else
'                                olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance - cVB_RoundRcv), False, True) 'แสดง Balance ที่มีการบวกเศษด้วย เพราะการจ่ายด้วยเครดิตสามารถจ่ายเป็นเศษได้
'                        End Select
                        olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance), False, True)  '*Eaw 56-01-29แสดง Balance ที่มีการปัดเศษ เพราะ cash ต้องมีการ Rouding
                        '----------------------------------------------------------------------------------------------
                        tVB_CNMnuCurrent = "1828"
                        If aVB_TenderMnu(nVB_CurTenderSeqNo).Autho Then
'                            If Not SP_USRbMnuAccept(nEN_Usr2AthRead, tVB_CNMnuCurrent) Then Exit Sub
                            If Not SP_USRbMnuAccept(nEN_Usr2AthRead, tVB_CNMnuCurrent) Then  '*Eaw 56-07-02
                                bVB_CNTenderAct = False
                                otmForm.Enabled = True
                                otmLock.Enabled = True
                                Call SP_DATxRsNothing(orsTemp)
                                Call SP_OBJxSetFocus(ogdFunc) ' ogdFunc.SetFocus
                                Exit Sub
                            End If
                        End If
                        tVB_TenderCode = "T013"
                        tVB_TenderType = aVB_TenderMnu(nVB_CurTenderSeqNo).Name
'                        wCNPayCredit.Caption = aVB_TenderMnu(nVB_CurTenderSeqNo).Menu  '*Eaw 55-11-27
'                        wCNPayNETS.Caption = aVB_TenderMnu(nVB_CurTenderSeqNo).Menu  '*Eaw 55-11-27
                        tVB_TenderName = aVB_TenderMnu(nVB_CurTenderSeqNo).DescSale
'                        wCNPayCredit.Show vbModal  '*Eaw 55-11-27
'                        wCNPayNETS.Show vbModal
                        If SP_NETbCheckPort Then   '*Eaw 56-06-17
'                            wCNPayCreditPay.Show vbModal
                            '*Eaw 56-06-11
                            Select Case Val(Format(tVB_TransactionSub))
                                Case tEN_TrnReturnManualTax, tEN_TrnReturnTaxInc, tEN_TrnSaleManualTax, tEN_TrnSaleTaxInc, tEN_TrnReturn, tEN_TrnDepositCancel, tEN_TrnAddDepositCancel, tEN_TrnCancelVoucher, tEN_TrnOtherDisbursement  '*Eaw 56-11-26 เพิ่มเคสจากโหมดปกติเข้าไป CommSheet TK-ISS3000-363-FAPL-129.xls
                                   wCNPayNETS.Show vbModal
                                Case Else
                                    wCNPayCreditPay.Show vbModal
                            End Select
                        Else
                            wCNPayNETS.Show vbModal
                        End If
                        If cVB_Pay = 0 Then
                            If cVB_Balance > 0 Then
                                '*Eaw 56-01-04 แก้ไขการปัด Rounding
'                                Select Case Val(Format(tVB_TransactionSub))
'                                    Case tEN_TrnDepositCancel, tEN_TrnAddDepositCancel
'                                        olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance - cVB_RoundRcv), False, True) 'แสดง Balance ที่มีการบวกเศษด้วย เพราะการจ่ายด้วยเครดิตสามารถจ่ายเป็นเศษได้
'                                    Case Else
'                                        olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance - cVB_RoundRcv), False, True) 'แสดง Balance ที่มีการบวกเศษด้วย เพราะการจ่ายด้วยเครดิตสามารถจ่ายเป็นเศษได้
'                                End Select
                                cVB_Balance = cVB_Balance - cVB_RoundRcv
                                olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance), False, True)  '*Eaw 56-01-29แสดง Balance ที่มีการปัดเศษ เพราะ cash ต้องมีการ Rouding
                                cVB_RoundRcv = 0
                                '________________________________________
'                                olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance), False, True) 'แสดง Balance ที่มีปัดเศษเหมือนเดิมเพราะไม่มีการชำระด้วยเครดิต
                            Else
                                Call SP_GETxBalanceRnd  '*Eaw 56-07-01 ทำเป็นฟังก์ชันกลาง
                                '*KT 51-11-12 อย่าลืมทำที่ Credit Store ด้วย
                                '---------------------------------------------------------------------------------------------------------
'                                tSql = "SELECT FNMenMin,FNMenValue FROM TSysMeaning"
'                                Call SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine, , , , , True, True) 'read master
'                                If Not (orsTemp.BOF Or orsTemp.EOF) Then
'                                    cMeaningDef = CDbl(Val(Format(SP_FEDvChkNumeric(orsTemp, "FNMenValue"))) / 100)
'                                Else
'                                    cMeaningDef = 0
'                                End If
'                                Call SP_DATxRsNothing(orsTemp)
'                                If bVB_CNAlwRnd Then
'                                    Select Case nVB_CNTypeRnd
'                                        Case eRndUp
'                                            If SP_FMTcCurPoint(cVB_Balance - cVB_RoundRcv, True, False) < 1 Then
'                                                bVB_CNPayCrdNoSuccess = True
'                                                olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance - cVB_RoundRcv), False, True) 'แสดง Balance ที่มีการบวกเศษด้วย เพราะการจ่ายด้วยเครดิตสามารถจ่ายเป็นเศษได้
'                                            Else
'                                                bVB_CNPayCrdNoSuccess = False
'                                            End If
'                                        Case eRndDn
'                                            If SP_FMTcCurPoint(cVB_Balance - cVB_RoundRcv, True, False) < 1 Then
'                                                bVB_CNPayCrdNoSuccess = True
'                                                olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance - cVB_RoundRcv), False, True) 'แสดง Balance ที่มีการบวกเศษด้วย เพราะการจ่ายด้วยเครดิตสามารถจ่ายเป็นเศษได้
'                                            Else
'                                                bVB_CNPayCrdNoSuccess = False
'                                            End If
'                                        Case eRndMeaning
'                                            If SP_FMTcCurPoint(cVB_Balance - cVB_RoundRcv, True, False) <= cMeaningDef Then
'                                                bVB_CNPayCrdNoSuccess = True
'                                                olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance - cVB_RoundRcv), False, True) 'แสดง Balance ที่มีการบวกเศษด้วย เพราะการจ่ายด้วยเครดิตสามารถจ่ายเป็นเศษได้
'                                            Else
'                                                bVB_CNPayCrdNoSuccess = False
'                                            End If
'                                    End Select
'                                Else
'                                    If SP_FMTcCurPoint(cVB_Balance - cVB_RoundRcv, True, False) < 1 Then
'                                        bVB_CNPayCrdNoSuccess = True
'                                       olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance - cVB_RoundRcv), False, True) 'แสดง Balance ที่มีการบวกเศษด้วย เพราะการจ่ายด้วยเครดิตสามารถจ่ายเป็นเศษได้
'                                    Else
'                                        bVB_CNPayCrdNoSuccess = False
'                                    End If
'                                End If
                                '---------------------------------------------------------------------------------------------------------
                            End If
                            bVB_CNTenderAct = False
                            Call SP_DATxRsNothing(orsTemp)
                            Call SP_OBJxSetFocus(ogdFunc)
                            Exit Sub
                        End If
                        Call SP_OBJxSetFocus(ogdFunc)
                        Call W_PRCxAddPayCrdCard
                        bVB_CNTenderAct = False
                        If cVB_Balance = 0 Then
                            Call SP_DATxRsNothing(orsTemp)
                            Unload Me
                            Exit Sub
                        Else
                            otmForm.Enabled = True
                            otmLock.Enabled = True
                        End If
                         
                  Case "T015" 'NETS
                        tVB_TypeOfEDC = "4"
                        bVB_CNTenderAct = True
                        '*Eaw 56-01-03 ก่อนเข้า  Tender ไม่มีการ Rounding ใช้ค่า Balance ที่มีอยู่ได้เลย
'                        Select Case Val(Format(tVB_TransactionSub))
'                            Case tEN_TrnDepositCancel, tEN_TrnAddDepositCancel
'                                olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance - cVB_RoundRcv), False, True) 'แสดง Balance ที่มีการบวกเศษด้วย เพราะการจ่ายด้วยเครดิตสามารถจ่ายเป็นเศษได้
'                            Case Else
'                                olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance - cVB_RoundRcv), False, True) 'แสดง Balance ที่มีการบวกเศษด้วย เพราะการจ่ายด้วยเครดิตสามารถจ่ายเป็นเศษได้
'                        End Select
                        olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance), False, True)  '*Eaw 56-01-29แสดง Balance ที่มีการปัดเศษ เพราะ cash ต้องมีการ Rouding
                        '-----------------------------------------------------------------------------------------------------
                        tVB_CNMnuCurrent = "1832"
                        If aVB_TenderMnu(nVB_CurTenderSeqNo).Autho Then
'                            If Not SP_USRbMnuAccept(nEN_Usr2AthRead, tVB_CNMnuCurrent) Then Exit Sub
                            If Not SP_USRbMnuAccept(nEN_Usr2AthRead, tVB_CNMnuCurrent) Then  '*Eaw 56-07-02
                                bVB_CNTenderAct = False
                                otmForm.Enabled = True
                                otmLock.Enabled = True
                                Call SP_DATxRsNothing(orsTemp)
                                Call SP_OBJxSetFocus(ogdFunc) ' ogdFunc.SetFocus
                                Exit Sub
                            End If
                        End If
                        tVB_TenderCode = "T015"
                        tVB_TenderType = aVB_TenderMnu(nVB_CurTenderSeqNo).Name
'                        wCNPayCredit.Caption = aVB_TenderMnu(nVB_CurTenderSeqNo).Menu  '*Eaw 55-11-27
'                        wCNPayNETS.Caption = aVB_TenderMnu(nVB_CurTenderSeqNo).Menu  '*Eaw 55-11-27
                        tVB_TenderName = aVB_TenderMnu(nVB_CurTenderSeqNo).DescSale
'                        wCNPayCredit.Show vbModal  '*Eaw 55-11-27
'                        wCNPayNETS.Show vbModal  '*Eaw 55-11-27
                         wCNPayNETSType.Show vbModal             '*Ao 2013/01/04
                        If cVB_Pay = 0 Then
                            If cVB_Balance > 0 Then
                                '*Eaw 56-01-04 แก้ไขการปัด Rounding
'                                Select Case Val(Format(tVB_TransactionSub))
'                                    Case tEN_TrnDepositCancel, tEN_TrnAddDepositCancel
'                                        olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance - cVB_RoundRcv), False, True) 'แสดง Balance ที่มีการบวกเศษด้วย เพราะการจ่ายด้วยเครดิตสามารถจ่ายเป็นเศษได้
'                                    Case Else
'                                        olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance - cVB_RoundRcv), False, True) 'แสดง Balance ที่มีการบวกเศษด้วย เพราะการจ่ายด้วยเครดิตสามารถจ่ายเป็นเศษได้
'                                End Select
                                cVB_Balance = cVB_Balance - cVB_RoundRcv
                                olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance), False, True)  '*Eaw 56-01-29แสดง Balance ที่มีการปัดเศษ เพราะ cash ต้องมีการ Rouding
                                cVB_RoundRcv = 0
                                '________________________________________
'                                olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance), False, True) 'แสดง Balance ที่มีปัดเศษเหมือนเดิมเพราะไม่มีการชำระด้วยเครดิต
                            Else
                                Call SP_GETxBalanceRnd  '*Eaw 56-07-01 ทำเป็นฟังก์ชันกลาง
                                '*KT 51-11-12 อย่าลืมทำที่ Credit Store ด้วย
                                '---------------------------------------------------------------------------------------------------------
'                                tSql = "SELECT FNMenMin,FNMenValue FROM TSysMeaning"
'                                Call SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine, , , , , True, True) 'read master
'                                If Not (orsTemp.BOF Or orsTemp.EOF) Then
'                                    cMeaningDef = CDbl(Val(Format(SP_FEDvChkNumeric(orsTemp, "FNMenValue"))) / 100)
'                                Else
'                                    cMeaningDef = 0
'                                End If
'                                Call SP_DATxRsNothing(orsTemp)
'                                If bVB_CNAlwRnd Then
'                                    Select Case nVB_CNTypeRnd
'                                        Case eRndUp
'                                            If SP_FMTcCurPoint(cVB_Balance - cVB_RoundRcv, True, False) < 1 Then
'                                                bVB_CNPayCrdNoSuccess = True
'                                                olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance - cVB_RoundRcv), False, True) 'แสดง Balance ที่มีการบวกเศษด้วย เพราะการจ่ายด้วยเครดิตสามารถจ่ายเป็นเศษได้
'                                            Else
'                                                bVB_CNPayCrdNoSuccess = False
'                                            End If
'                                        Case eRndDn
'                                            If SP_FMTcCurPoint(cVB_Balance - cVB_RoundRcv, True, False) < 1 Then
'                                                bVB_CNPayCrdNoSuccess = True
'                                                olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance - cVB_RoundRcv), False, True) 'แสดง Balance ที่มีการบวกเศษด้วย เพราะการจ่ายด้วยเครดิตสามารถจ่ายเป็นเศษได้
'                                            Else
'                                                bVB_CNPayCrdNoSuccess = False
'                                            End If
'                                        Case eRndMeaning
'                                            If SP_FMTcCurPoint(cVB_Balance - cVB_RoundRcv, True, False) <= cMeaningDef Then
'                                                bVB_CNPayCrdNoSuccess = True
'                                                olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance - cVB_RoundRcv), False, True) 'แสดง Balance ที่มีการบวกเศษด้วย เพราะการจ่ายด้วยเครดิตสามารถจ่ายเป็นเศษได้
'                                            Else
'                                                bVB_CNPayCrdNoSuccess = False
'                                            End If
'                                    End Select
'                                Else
'                                    If SP_FMTcCurPoint(cVB_Balance - cVB_RoundRcv, True, False) < 1 Then
'                                        bVB_CNPayCrdNoSuccess = True
'                                       olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance - cVB_RoundRcv), False, True) 'แสดง Balance ที่มีการบวกเศษด้วย เพราะการจ่ายด้วยเครดิตสามารถจ่ายเป็นเศษได้
'                                    Else
'                                        bVB_CNPayCrdNoSuccess = False
'                                    End If
'                                End If
                                '---------------------------------------------------------------------------------------------------------
                            End If
                            bVB_CNTenderAct = False
                            Call SP_DATxRsNothing(orsTemp)
                            Call SP_OBJxSetFocus(ogdFunc)
                            Exit Sub
                        End If
                        Call SP_OBJxSetFocus(ogdFunc)
                        Call W_PRCxAddPayCrdCard
                        bVB_CNTenderAct = False
                        If cVB_Balance = 0 Then
                            Call SP_DATxRsNothing(orsTemp)
                            Unload Me
                            Exit Sub
                        Else
                            otmForm.Enabled = True
                            otmLock.Enabled = True
                        End If
                  Case "T016" 'Ezlink  change to Flash Pay  '*Eaw 55-11-27
                        tVB_TypeOfEDC = "5"
                        bVB_CNTenderAct = True
                        '*Eaw 56-01-03 ก่อนเข้า  Tender ไม่มีการ Rounding ใช้ค่า Balance ที่มีอยู่ได้เลย
'                        Select Case Val(Format(tVB_TransactionSub))
'                            Case tEN_TrnDepositCancel, tEN_TrnAddDepositCancel
'                                olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance - cVB_RoundRcv), False, True) 'แสดง Balance ที่มีการบวกเศษด้วย เพราะการจ่ายด้วยเครดิตสามารถจ่ายเป็นเศษได้
'                            Case Else
'                                olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance - cVB_RoundRcv), False, True) 'แสดง Balance ที่มีการบวกเศษด้วย เพราะการจ่ายด้วยเครดิตสามารถจ่ายเป็นเศษได้
'                        End Select
                        olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance), False, True)  '*Eaw 56-01-29แสดง Balance ที่มีการปัดเศษ เพราะ cash ต้องมีการ Rouding
                        '--------------------------------------------------------------------------------------
                        tVB_CNMnuCurrent = "1833"
                        If aVB_TenderMnu(nVB_CurTenderSeqNo).Autho Then
'                            If Not SP_USRbMnuAccept(nEN_Usr2AthRead, tVB_CNMnuCurrent) Then Exit Sub
                            If Not SP_USRbMnuAccept(nEN_Usr2AthRead, tVB_CNMnuCurrent) Then  '*Eaw 56-07-02
                                bVB_CNTenderAct = False
                                otmForm.Enabled = True
                                otmLock.Enabled = True
                                Call SP_DATxRsNothing(orsTemp)
                                Call SP_OBJxSetFocus(ogdFunc) ' ogdFunc.SetFocus
                                Exit Sub
                            End If
                        End If
                        tVB_TenderCode = "T016"
                        tVB_TenderType = aVB_TenderMnu(nVB_CurTenderSeqNo).Name
'                        wCNPayCredit.Caption = aVB_TenderMnu(nVB_CurTenderSeqNo).Menu  '*Eaw 55-11-27
'                        wCNPayNETS.Caption = aVB_TenderMnu(nVB_CurTenderSeqNo).Menu  '*Eaw 55-11-27
                        tVB_TenderName = aVB_TenderMnu(nVB_CurTenderSeqNo).DescSale
                        'wCNPayCredit.Show vbModal    '*Eaw 55-11-27
'                        wCNPayNETS.Show vbModal
                        If SP_NETbCheckPort Then   '*Eaw 56-06-17
'                            wCNPayCreditPay.Show vbModal
                            '*Eaw 56-06-11
                            Select Case Val(Format(tVB_TransactionSub))
                                Case tEN_TrnReturnManualTax, tEN_TrnReturnTaxInc, tEN_TrnSaleManualTax, tEN_TrnSaleTaxInc
                                   wCNPayNETS.Show vbModal
                                Case Else
                                    wCNPayCreditPay.Show vbModal
                            End Select
                        Else
                            wCNPayNETS.Show vbModal
                        End If
                        If cVB_Pay = 0 Then
                            If cVB_Balance > 0 Then
                                '*Eaw 56-01-04 แก้ไขการปัด Rounding
'                                Select Case Val(Format(tVB_TransactionSub))
'                                    Case tEN_TrnDepositCancel, tEN_TrnAddDepositCancel
'                                        olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance - cVB_RoundRcv), False, True) 'แสดง Balance ที่มีการบวกเศษด้วย เพราะการจ่ายด้วยเครดิตสามารถจ่ายเป็นเศษได้
'                                    Case Else
'                                        olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance - cVB_RoundRcv), False, True) 'แสดง Balance ที่มีการบวกเศษด้วย เพราะการจ่ายด้วยเครดิตสามารถจ่ายเป็นเศษได้
'                                End Select
                                cVB_Balance = cVB_Balance - cVB_RoundRcv
                                olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance), False, True)  '*Eaw 56-01-29แสดง Balance ที่มีการปัดเศษ เพราะ cash ต้องมีการ Rouding
                                cVB_RoundRcv = 0
                                '________________________________________
'                                olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance), False, True) 'แสดง Balance ที่มีปัดเศษเหมือนเดิมเพราะไม่มีการชำระด้วยเครดิต
                            Else
                                Call SP_GETxBalanceRnd  '*Eaw 56-07-01 ทำเป็นฟังก์ชันกลาง
                                '*KT 51-11-12 อย่าลืมทำที่ Credit Store ด้วย
                                '---------------------------------------------------------------------------------------------------------
'                                tSql = "SELECT FNMenMin,FNMenValue FROM TSysMeaning"
'                                Call SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine, , , , , True, True) 'read master
'                                If Not (orsTemp.BOF Or orsTemp.EOF) Then
'                                    cMeaningDef = CDbl(Val(Format(SP_FEDvChkNumeric(orsTemp, "FNMenValue"))) / 100)
'                                Else
'                                    cMeaningDef = 0
'                                End If
'                                Call SP_DATxRsNothing(orsTemp)
'                                If bVB_CNAlwRnd Then
'                                    Select Case nVB_CNTypeRnd
'                                        Case eRndUp
'                                            If SP_FMTcCurPoint(cVB_Balance - cVB_RoundRcv, True, False) < 1 Then
'                                                bVB_CNPayCrdNoSuccess = True
'                                                olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance - cVB_RoundRcv), False, True) 'แสดง Balance ที่มีการบวกเศษด้วย เพราะการจ่ายด้วยเครดิตสามารถจ่ายเป็นเศษได้
'                                            Else
'                                                bVB_CNPayCrdNoSuccess = False
'                                            End If
'                                        Case eRndDn
'                                            If SP_FMTcCurPoint(cVB_Balance - cVB_RoundRcv, True, False) < 1 Then
'                                                bVB_CNPayCrdNoSuccess = True
'                                                olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance - cVB_RoundRcv), False, True) 'แสดง Balance ที่มีการบวกเศษด้วย เพราะการจ่ายด้วยเครดิตสามารถจ่ายเป็นเศษได้
'                                            Else
'                                                bVB_CNPayCrdNoSuccess = False
'                                            End If
'                                        Case eRndMeaning
'                                            If SP_FMTcCurPoint(cVB_Balance - cVB_RoundRcv, True, False) <= cMeaningDef Then
'                                                bVB_CNPayCrdNoSuccess = True
'                                                olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance - cVB_RoundRcv), False, True) 'แสดง Balance ที่มีการบวกเศษด้วย เพราะการจ่ายด้วยเครดิตสามารถจ่ายเป็นเศษได้
'                                            Else
'                                                bVB_CNPayCrdNoSuccess = False
'                                            End If
'                                    End Select
'                                Else
'                                    If SP_FMTcCurPoint(cVB_Balance - cVB_RoundRcv, True, False) < 1 Then
'                                        bVB_CNPayCrdNoSuccess = True
'                                       olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance - cVB_RoundRcv), False, True) 'แสดง Balance ที่มีการบวกเศษด้วย เพราะการจ่ายด้วยเครดิตสามารถจ่ายเป็นเศษได้
'                                    Else
'                                        bVB_CNPayCrdNoSuccess = False
'                                    End If
'                                End If
                                '---------------------------------------------------------------------------------------------------------
                            End If
                            bVB_CNTenderAct = False
                            Call SP_DATxRsNothing(orsTemp)
                            Call SP_OBJxSetFocus(ogdFunc)
                            Exit Sub
                        End If
                        Call SP_OBJxSetFocus(ogdFunc)
                        Call W_PRCxAddPayCrdCard
                        bVB_CNTenderAct = False
                        If cVB_Balance = 0 Then
                            Call SP_DATxRsNothing(orsTemp)
                            Unload Me
                            Exit Sub
                        Else
                            otmForm.Enabled = True
                            otmLock.Enabled = True
                        End If
                  Case "T017" 'TC Card
                        tVB_TypeOfEDC = ""
                        bVB_CNTenderAct = True
                        '*Eaw 56-01-03 ก่อนเข้า  Tender ไม่มีการ Rounding ใช้ค่า Balance ที่มีอยู่ได้เลย
'                        Select Case Val(Format(tVB_TransactionSub))
'                            Case tEN_TrnDepositCancel, tEN_TrnAddDepositCancel
'                                olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance - cVB_RoundRcv), False, True) 'แสดง Balance ที่มีการบวกเศษด้วย เพราะการจ่ายด้วยเครดิตสามารถจ่ายเป็นเศษได้
'                            Case Else
'                                olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance - cVB_RoundRcv), False, True) 'แสดง Balance ที่มีการบวกเศษด้วย เพราะการจ่ายด้วยเครดิตสามารถจ่ายเป็นเศษได้
'                        End Select
                        olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance), False, True)  '*Eaw 56-01-29แสดง Balance ที่มีการปัดเศษ เพราะ cash ต้องมีการ Rouding
                        '------------------------------------------------------------------------------------------
                        tVB_CNMnuCurrent = "1831"
                        If aVB_TenderMnu(nVB_CurTenderSeqNo).Autho Then
'                        If Not SP_USRbMnuAccept(nEN_Usr2AthRead, tVB_CNMnuCurrent) Then Exit Sub
                            If Not SP_USRbMnuAccept(nEN_Usr2AthRead, tVB_CNMnuCurrent) Then  '*Eaw 56-07-02
                                    bVB_CNTenderAct = False
                                    otmForm.Enabled = True
                                    otmLock.Enabled = True
                                    Call SP_DATxRsNothing(orsTemp)
                                    Call SP_OBJxSetFocus(ogdFunc) ' ogdFunc.SetFocus
                                    Exit Sub
                            End If
                        End If
                        tVB_TenderCode = "T017"
                        tVB_TenderType = aVB_TenderMnu(nVB_CurTenderSeqNo).Name
'                        wCNPayCredit.Caption = aVB_TenderMnu(nVB_CurTenderSeqNo).Menu          '*Ao 2013/03/20  TK
                        tVB_TenderName = aVB_TenderMnu(nVB_CurTenderSeqNo).DescSale
'                        wCNPayCredit.Show vbModal
                        '*Eaw 56-08-28 CommSheet TK-ISS3000-323-UAT.xlsx
                        Select Case Val(Format(tVB_TransactionSub))
                            Case tEN_TrnReturnManualTax, tEN_TrnReturnTaxInc, tEN_TrnSaleManualTax, tEN_TrnSaleTaxInc, tEN_TrnReturn, tEN_TrnDepositCancel, tEN_TrnAddDepositCancel, tEN_TrnCancelVoucher, tEN_TrnOtherDisbursement  '*Eaw 56-11-26 เพิ่มเคสจากโหมดปกติเข้าไป CommSheet TK-ISS3000-363-FAPL-129.xls
                                wCNPayCreditRE.Show vbModal
                            Case Else
                                wCNPayCredit.Show vbModal
                        End Select
                        If cVB_Pay = 0 Then
                            If cVB_Balance > 0 Then
                                '*Eaw 56-01-04 แก้ไขการปัด Rounding
'                                Select Case Val(Format(tVB_TransactionSub))
'                                    Case tEN_TrnDepositCancel, tEN_TrnAddDepositCancel
'                                        olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance - cVB_RoundRcv), False, True) 'แสดง Balance ที่มีการบวกเศษด้วย เพราะการจ่ายด้วยเครดิตสามารถจ่ายเป็นเศษได้
'                                    Case Else
'                                        olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance - cVB_RoundRcv), False, True) 'แสดง Balance ที่มีการบวกเศษด้วย เพราะการจ่ายด้วยเครดิตสามารถจ่ายเป็นเศษได้
'                                End Select
                                cVB_Balance = cVB_Balance - cVB_RoundRcv
                                olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance), False, True)  '*Eaw 56-01-29แสดง Balance ที่มีการปัดเศษ เพราะ cash ต้องมีการ Rouding
                                cVB_RoundRcv = 0
                                '________________________________________
'                                olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance), False, True) 'แสดง Balance ที่มีปัดเศษเหมือนเดิมเพราะไม่มีการชำระด้วยเครดิต
                            Else
                                Call SP_GETxBalanceRnd  '*Eaw 56-07-01 ทำเป็นฟังก์ชันกลาง
                                '*KT 51-11-12 อย่าลืมทำที่ Credit Store ด้วย
                                '---------------------------------------------------------------------------------------------------------
'                                tSql = "SELECT FNMenMin,FNMenValue FROM TSysMeaning"
'                                Call SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine, , , , , True, True) 'read master
'                                If Not (orsTemp.BOF Or orsTemp.EOF) Then
'                                    cMeaningDef = CDbl(Val(Format(SP_FEDvChkNumeric(orsTemp, "FNMenValue"))) / 100)
'                                Else
'                                    cMeaningDef = 0
'                                End If
'                                Call SP_DATxRsNothing(orsTemp)
'                                If bVB_CNAlwRnd Then
'                                    Select Case nVB_CNTypeRnd
'                                        Case eRndUp
'                                            If SP_FMTcCurPoint(cVB_Balance - cVB_RoundRcv, True, False) < 1 Then
'                                                bVB_CNPayCrdNoSuccess = True
'                                                olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance - cVB_RoundRcv), False, True) 'แสดง Balance ที่มีการบวกเศษด้วย เพราะการจ่ายด้วยเครดิตสามารถจ่ายเป็นเศษได้
'                                            Else
'                                                bVB_CNPayCrdNoSuccess = False
'                                            End If
'                                        Case eRndDn
'                                            If SP_FMTcCurPoint(cVB_Balance - cVB_RoundRcv, True, False) < 1 Then
'                                                bVB_CNPayCrdNoSuccess = True
'                                                olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance - cVB_RoundRcv), False, True) 'แสดง Balance ที่มีการบวกเศษด้วย เพราะการจ่ายด้วยเครดิตสามารถจ่ายเป็นเศษได้
'                                            Else
'                                                bVB_CNPayCrdNoSuccess = False
'                                            End If
'                                        Case eRndMeaning
'                                            If SP_FMTcCurPoint(cVB_Balance - cVB_RoundRcv, True, False) <= cMeaningDef Then
'                                                bVB_CNPayCrdNoSuccess = True
'                                                olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance - cVB_RoundRcv), False, True) 'แสดง Balance ที่มีการบวกเศษด้วย เพราะการจ่ายด้วยเครดิตสามารถจ่ายเป็นเศษได้
'                                            Else
'                                                bVB_CNPayCrdNoSuccess = False
'                                            End If
'                                    End Select
'                                Else
'                                    If SP_FMTcCurPoint(cVB_Balance - cVB_RoundRcv, True, False) < 1 Then
'                                        bVB_CNPayCrdNoSuccess = True
'                                       olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance - cVB_RoundRcv), False, True) 'แสดง Balance ที่มีการบวกเศษด้วย เพราะการจ่ายด้วยเครดิตสามารถจ่ายเป็นเศษได้
'                                    Else
'                                        bVB_CNPayCrdNoSuccess = False
'                                    End If
'                                End If
                                '---------------------------------------------------------------------------------------------------------
                            End If
                            bVB_CNTenderAct = False
                            Call SP_DATxRsNothing(orsTemp)
                            Call SP_OBJxSetFocus(ogdFunc)
                            Exit Sub
                        End If
                        Call SP_OBJxSetFocus(ogdFunc)
                        Call W_PRCxAddPayCrdCard
                        bVB_CNTenderAct = False
                        If cVB_Balance = 0 Then
                            Call SP_DATxRsNothing(orsTemp)
                            Unload Me
                            Exit Sub
                        Else
                            otmForm.Enabled = True
                            otmLock.Enabled = True
                        End If
                Case "T018" 'Credit Voucher
'*Eaw 56-06-17 CV สามารถสร้างขณะ Offline ได้
'                         If Not oDB.SQLbAvailableServer(oDB.OffServerName, oDB.OnComIP) Then  '*Eaw 56-06-10 CommSheet TK-ISS3000-FAPL-025-Eaw.xls
'                            Call SP_MSGnShowing("tms_CN005247", nCS_Warning)
'                            Exit Sub
'                         End If
                        bVB_CNTenderAct = True
                        
                        Select Case Val(Format(tVB_TransactionSub))
                            Case tEN_TrnReturn, tEN_TrnReturnTaxInc, tEN_TrnReturnManualTax
                                If tVB_CNCVNo <> "" Then
                                    Call SP_MSGnShowing("tms_CN005235", nCS_Warning)
                                    bVB_CNTenderAct = False
                                    Call SP_DATxRsNothing(orsTemp)
                                    Call SP_OBJxSetFocus(ogdFunc)
                                    Exit Sub
                                End If
                            Case tEN_TrnAddDepositCancel, tEN_TrnDepositCancel
                                If tVB_CNCVNo <> "" Then
                                    Call SP_MSGnShowing("tms_CN005235", nCS_Warning)
                                    bVB_CNTenderAct = False
                                    Call SP_DATxRsNothing(orsTemp)
                                    Call SP_OBJxSetFocus(ogdFunc)
                                    Exit Sub
                                End If
                            Case Else
                                If tVB_CNCVNoReIssue <> "" Then
                                    Call SP_MSGnShowing("tms_CN005235", nCS_Warning)
                                    bVB_CNTenderAct = False
                                    Call SP_DATxRsNothing(orsTemp)
                                    Call SP_OBJxSetFocus(ogdFunc)
                                    Exit Sub
                                End If
                        End Select
                        '*Eaw 56-01-03 ก่อนเข้า  Tender ไม่มีการ Rounding ใช้ค่า Balance ที่มีอยู่ได้เลย
'                        Select Case Val(Format(tVB_TransactionSub))
'                            Case tEN_TrnDepositCancel, tEN_TrnAddDepositCancel
'                                olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance - cVB_RoundRcv), False, True) 'แสดง Balance ที่มีการบวกเศษด้วย เพราะการจ่ายด้วยเครดิตสามารถจ่ายเป็นเศษได้
'                            Case Else
'                                olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance - cVB_RoundRcv), False, True) 'แสดง Balance ที่มีการบวกเศษด้วย เพราะการจ่ายด้วยเครดิตสามารถจ่ายเป็นเศษได้
'                        End Select
                        olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance), False, True)  '*Eaw 56-01-29แสดง Balance ที่มีการปัดเศษ เพราะ cash ต้องมีการ Rouding
                        '--------------------------------------------------------------------------------------------
                        tVB_CNMnuCurrent = "1831"
                        If aVB_TenderMnu(nVB_CurTenderSeqNo).Autho Then
'                            If Not SP_USRbMnuAccept(nEN_Usr2AthRead, tVB_CNMnuCurrent) Then Exit Sub
                            If Not SP_USRbMnuAccept(nEN_Usr2AthRead, tVB_CNMnuCurrent) Then  '*Eaw 56-07-02
                                bVB_CNTenderAct = False
                                otmForm.Enabled = True
                                otmLock.Enabled = True
                                Call SP_DATxRsNothing(orsTemp)
                                Call SP_OBJxSetFocus(ogdFunc) ' ogdFunc.SetFocus
                                Exit Sub
                            End If
                        End If
                        tVB_TenderCode = "T018"
                         tVB_TenderType = aVB_TenderMnu(nVB_CurTenderSeqNo).Name
                        tVB_TenderName = aVB_TenderMnu(nVB_CurTenderSeqNo).DescSale
                        wCNPayCV.Show vbModal
                       If cVB_Pay = 0 Then
                            If cVB_Balance > 0 Then
                                '*Eaw 56-01-04 แก้ไขการปัด Rounding
'                                Select Case Val(Format(tVB_TransactionSub))
'                                    Case tEN_TrnDepositCancel, tEN_TrnAddDepositCancel
'                                        olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance - cVB_RoundRcv), False, True) 'แสดง Balance ที่มีการบวกเศษด้วย เพราะการจ่ายด้วยเครดิตสามารถจ่ายเป็นเศษได้
'                                    Case Else
'                                        olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance - cVB_RoundRcv), False, True) 'แสดง Balance ที่มีการบวกเศษด้วย เพราะการจ่ายด้วยเครดิตสามารถจ่ายเป็นเศษได้
'                                End Select
                                cVB_Balance = cVB_Balance - cVB_RoundRcv
                                olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance), False, True)  '*Eaw 56-01-29แสดง Balance ที่มีการปัดเศษ เพราะ cash ต้องมีการ Rouding
                                cVB_RoundRcv = 0
                                '________________________________________
'                                olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance), False, True) 'แสดง Balance ที่มีปัดเศษเหมือนเดิมเพราะไม่มีการชำระด้วยเครดิต
                            Else
                                Call SP_GETxBalanceRnd  '*Eaw 56-07-01 ทำเป็นฟังก์ชันกลาง
                                '*KT 51-11-12 อย่าลืมทำที่ Credit Store ด้วย
                                '---------------------------------------------------------------------------------------------------------
'                                tSql = "SELECT FNMenMin,FNMenValue FROM TSysMeaning"
'                                Call SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine, , , , , True, True) 'read master
'                                If Not (orsTemp.BOF Or orsTemp.EOF) Then
'                                    cMeaningDef = CDbl(Val(Format(SP_FEDvChkNumeric(orsTemp, "FNMenValue"))) / 100)
'                                Else
'                                    cMeaningDef = 0
'                                End If
'                                Call SP_DATxRsNothing(orsTemp)
'                                If bVB_CNAlwRnd Then
'                                    Select Case nVB_CNTypeRnd
'                                        Case eRndUp
'                                            If SP_FMTcCurPoint(cVB_Balance - cVB_RoundRcv, True, False) < 1 Then
'                                                bVB_CNPayCrdNoSuccess = True
'                                                olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance - cVB_RoundRcv), False, True) 'แสดง Balance ที่มีการบวกเศษด้วย เพราะการจ่ายด้วยเครดิตสามารถจ่ายเป็นเศษได้
'                                            Else
'                                                bVB_CNPayCrdNoSuccess = False
'                                            End If
'                                        Case eRndDn
'                                            If SP_FMTcCurPoint(cVB_Balance - cVB_RoundRcv, True, False) < 1 Then
'                                                bVB_CNPayCrdNoSuccess = True
'                                                olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance - cVB_RoundRcv), False, True) 'แสดง Balance ที่มีการบวกเศษด้วย เพราะการจ่ายด้วยเครดิตสามารถจ่ายเป็นเศษได้
'                                            Else
'                                                bVB_CNPayCrdNoSuccess = False
'                                            End If
'                                        Case eRndMeaning
'                                            If SP_FMTcCurPoint(cVB_Balance - cVB_RoundRcv, True, False) <= cMeaningDef Then
'                                                bVB_CNPayCrdNoSuccess = True
'                                                olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance - cVB_RoundRcv), False, True) 'แสดง Balance ที่มีการบวกเศษด้วย เพราะการจ่ายด้วยเครดิตสามารถจ่ายเป็นเศษได้
'                                            Else
'                                                bVB_CNPayCrdNoSuccess = False
'                                            End If
'                                    End Select
'                                Else
'                                    If SP_FMTcCurPoint(cVB_Balance - cVB_RoundRcv, True, False) < 1 Then
'                                        bVB_CNPayCrdNoSuccess = True
'                                       olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance - cVB_RoundRcv), False, True) 'แสดง Balance ที่มีการบวกเศษด้วย เพราะการจ่ายด้วยเครดิตสามารถจ่ายเป็นเศษได้
'                                    Else
'                                        bVB_CNPayCrdNoSuccess = False
'                                    End If
'                                End If
                                '---------------------------------------------------------------------------------------------------------
                            End If
                            bVB_CNTenderAct = False
                            Call SP_DATxRsNothing(orsTemp)
                            Call SP_OBJxSetFocus(ogdFunc)
                            Exit Sub
                        End If
                        Call SP_OBJxSetFocus(ogdFunc)
                        Call W_PRCxAddPayCV
                        bVB_CNTenderAct = False
                        If cVB_Balance = 0 Then
                            
                            Call SP_DATxRsNothing(orsTemp)
                            Unload Me
                            Exit Sub
                        Else
                            otmForm.Enabled = True
                            otmLock.Enabled = True
                        End If
                Case "T019" '*Tao 60-08-25 WalletPay
                        bVB_CNTenderAct = True
                        tVB_CNMnuCurrent = .TextMatrix(.Row, 2)
                        If aVB_TenderMnu(nVB_CurTenderSeqNo).Autho Then
                            If Not SP_USRbMnuAccept(nEN_Usr2AthRead, tVB_CNMnuCurrent) Then
                                bVB_CNTenderAct = False
                                otmForm.Enabled = True
                                otmLock.Enabled = True
                                Call SP_DATxRsNothing(orsTemp)
                                Call SP_OBJxSetFocus(ogdFunc)
                                Exit Sub
                            End If
                        End If
                        tVB_TenderCode = Trim(UCase(.TextMatrix(.Row, 3)))
                        tVB_TenderType = Trim(UCase(.TextMatrix(.Row, 5)))
                        tVB_TenderName = Trim(UCase(.TextMatrix(.Row, 4)))
                        
                        '*Tao 60-09-04 Allow Wallet Payment
                        If Not bVB_AlwWalletPay Then
                            Call SP_MSGnShowing("tms_CN999048", nCS_Warning, , , , tVB_TenderName) 'ไม่อนุญาตให้ใช้ Wallet Payment
                            bVB_CNTenderAct = False
                            otmForm.Enabled = True
                            otmLock.Enabled = True
                            Call SP_DATxRsNothing(orsTemp)
                            Call SP_OBJxSetFocus(ogdFunc)
                            Exit Sub
                        End If
                        '===============================
                                               
                        Select Case SP_VALnText2Double(tVB_TransactionSub)
                        Case tEN_TrnDepositCancel, tEN_TrnAddDepositCancel 'Deposit ไม่ให้คืนด้วย WalletPay
                            Call SP_MSGnShowing("tms_CN999036", nCS_Warning) 'ไม่สามารถทำคืนสินค้าได้
                            bVB_CNTenderAct = False
                            otmForm.Enabled = True
                            otmLock.Enabled = True
                            Call SP_DATxRsNothing(orsTemp)
                            Call SP_OBJxSetFocus(ogdFunc)
                            Exit Sub
                        Case tEN_TrnIPVTaxExceptSale, tEN_TrnCancelVoucher, tEN_TrnIPVDutyFreeSale, tEN_TrnIPVVIPService, tEN_TrnIPVReturn, tEN_TrnIPVRedeem 'ไม่ให้คืนด้วย WalletPay
                            Call SP_MSGnShowing("tms_CN999036", nCS_Warning) 'ไม่สามารถทำคืนสินค้าได้
                            bVB_CNTenderAct = False
                            otmForm.Enabled = True
                            otmLock.Enabled = True
                            Call SP_DATxRsNothing(orsTemp)
                            Call SP_OBJxSetFocus(ogdFunc)
                            Exit Sub
                        Case tEN_TrnDeposit, tEN_TrnAddDeposit, tEN_TrnRedeem 'Deposit ไม่ให้ชำระด้วย WalletPay
                            Call SP_MSGnShowing("tms_CN999037", nCS_Warning) 'ไม่อนุญาตให้ชำระด้วย WalletPay
                            bVB_CNTenderAct = False
                            otmForm.Enabled = True
                            otmLock.Enabled = True
                            Call SP_DATxRsNothing(orsTemp)
                            Call SP_OBJxSetFocus(ogdFunc)
                            Exit Sub
                        Case tEN_TrnVIPSale, tEN_TrnVoucher, tEN_TrnDutyFreeSale 'ไม่ให้ชำระด้วย WalletPay
                            Call SP_MSGnShowing("tms_CN999037", nCS_Warning) 'ไม่อนุญาตให้ชำระด้วย WalletPay
                            bVB_CNTenderAct = False
                            otmForm.Enabled = True
                            otmLock.Enabled = True
                            Call SP_DATxRsNothing(orsTemp)
                            Call SP_OBJxSetFocus(ogdFunc)
                            Exit Sub
                        Case tEN_TrnReturn, tEN_TrnDepositCancel, tEN_TrnAddDepositCancel, tEN_TrnIPVTaxExceptSale, tEN_TrnCancelVoucher, tEN_TrnIPVDutyFreeSale, tEN_TrnIPVVIPService, tEN_TrnIPVReturn, tEN_TrnIPVRedeem, tEN_TrnReturnManualTax, tEN_TrnReturnTaxInc
                        
                            If cVB_Balance = 0 Then
                                Call SP_MSGnShowing("tms_CN999034", nCS_Warning) 'ไม่มียอดเงินที่จะทำคืนสินค้า
                                bVB_CNTenderAct = False
                                otmForm.Enabled = True
                                otmLock.Enabled = True
                                Call SP_DATxRsNothing(orsTemp)
                                Call SP_OBJxSetFocus(ogdFunc)
                                Exit Sub
                            End If
                            
                            'ถ้าเคยคืนสินค้าด้วย WalletPay แล้ว ไม่ใหคืนสินค้าด้วย WalletPay ได้อีก
                            tWhe = " WHERE  FTTmnNum='" & tVB_CNTerminalNum & "' AND FTShdTransNo='" & tVB_Running & "'"
                            tWhe = tWhe & " AND  FTShdTransType='" & tVB_Transaction & "' AND FDShdTransDate='" & tUT_SaleDate & "'"
                            tWhe = tWhe & " AND FTTdmCode IN ('T019','T020','T021','T022')"
                            nTmp = SP_VALcText2Double(SP_SQLtFunction("COUNT", "FTShdTransNo", tVB_CNTblRC, tWhe))
                            If nTmp > 0 Then
                                Call SP_MSGnShowing("tms_CN999043", nCS_Warning) 'มีการคืนสินค้าด้วย WalletPay แล้ว
                                bVB_CNTenderAct = False
                                otmForm.Enabled = True
                                otmLock.Enabled = True
                                Call SP_DATxRsNothing(orsTemp)
                                Call SP_OBJxSetFocus(ogdFunc)
                                Exit Sub
                            End If
                            '========================================================
                            
                            'ถ้าไม่ใช่การชำระด้วย WalletPay ไม่ให้คืนด้วย WalletPay
                            Select Case SP_VALnText2Double(tVB_TransactionSub)
                            Case tEN_TrnCancelVoucher 'Cancel Voucher
                                tWhe = " WITH(NOLOCK) WHERE  FTTmnNum='" & tVB_CancelTmnNumRef & "' AND FTShdTransNo='" & tVB_CancelTransNoRef & "'"
                                tWhe = tWhe & " AND  FTShdTransType='" & tEN_TrnVoucher & "' AND FDShdTransDate='" & Format(dVB_CancelTransDateRef, "YYYY/MM/DD") & "'"
                                tWhe = tWhe & " AND FTTdmCode IN ('T019','T020','T021','T022')"
                                tSql = "SELECT FCSrcNet, FTSrcRmk, FDDateIns, FTTimeIns,FDShdTransDate,FTSrcRetDocRef,FTSrcCardNo,FTTdmCode FROM " & tVB_CNTblRC & tWhe
                                Call SP_SQLvExecute(tSql, orsTemp)
                            Case tEN_TrnReturn 'Return
                                tWhe = " WITH(NOLOCK) WHERE  FTTmnNum='" & tVB_CNTerminalNum & "' AND FTShdTransNo='" & tVB_TransNoRet & "'"
                                tWhe = tWhe & " AND  FTShdTransType='" & tVB_TransTypeRet & "' AND FDShdTransDate='" & Format(dVB_TransDateRet, "YYYY/MM/DD") & "'"
                                tWhe = tWhe & " AND FTTdmCode IN ('T019','T020','T021','T022')"
                                tSql = "SELECT FCSrcNet, FTSrcRmk, FDDateIns, FTTimeIns,FDShdTransDate,FTSrcRetDocRef,FTSrcCardNo,FTTdmCode,FTTdmType,FTSrcAbb FROM " & tVB_CNTblRC & tWhe
                                Call SP_SQLvExecute(tSql, orsTemp)
                            Case tEN_TrnIPVReturn 'IPV Return
                                tWhe = " WITH(NOLOCK) WHERE  FTTmnNum='" & tVB_TerminalNoDptCancel & "' AND FTShdTransNo='" & tVB_TransNoDptCancel & "'"
                                tWhe = tWhe & " AND  FTShdTransType='" & tVB_TransTypeDptCancel & "' AND FDShdTransDate='" & Format(dVB_TransDateDptCancel, "YYYY/MM/DD") & "'"
                                tWhe = tWhe & " AND FTTdmCode IN ('T019','T020','T021','T022')"
                                tSql = "SELECT FCSrcNet, FTSrcRmk, FDDateIns, FTTimeIns,FDShdTransDate,FTSrcRetDocRef,FTSrcCardNo,FTTdmCode,FTTdmType,FTSrcAbb FROM " & tVB_CNTblRC & tWhe
                                Call SP_SQLvExecute(tSql, orsTemp)
                            Case tEN_TrnDepositCancel, tEN_TrnAddDepositCancel 'Deposit Cancel
                                bVB_Return = True
                                bVB_Reference = True
                                tWhe = " WITH(NOLOCK) WHERE  FTTmnNum='" & tVB_TerminalNoDptCancel & "' AND FTShdTransNo='" & tVB_TransNoDptCancel & "'"
                                tWhe = tWhe & " AND  FTShdTransType='" & tVB_TransTypeDptCancel & "' AND FDShdTransDate='" & Format(dVB_TransDateDptCancel, "YYYY/MM/DD") & "'"
                                tWhe = tWhe & " AND FTTdmCode IN ('T019','T020','T021','T022')"
                                tSql = "SELECT FCSrcNet, FTSrcRmk, FDDateIns, FTTimeIns,FDShdTransDate,FTSrcRetDocRef,FTSrcCardNo,FTTdmCode,FTTdmType,FTSrcAbb FROM TPSTDepositRC " & tWhe
                                Call SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine, , , , , True, True)
                            Case Else
                                tWhe = " WITH(NOLOCK) WHERE  FTTmnNum='" & tVB_CNTerminalNum & "' AND FTShdTransNo='" & tVB_TransNoRet & "'"
                                tWhe = tWhe & " AND  FTShdTransType='" & tVB_TransTypeRet & "' AND FDShdTransDate='" & Format(dVB_TransDateRet, "YYYY/MM/DD") & "'"
                                tWhe = tWhe & " AND FTTdmCode IN ('T019','T020','T021','T022')"
                                tSql = "SELECT FCSrcNet, FTSrcRmk, FDDateIns, FTTimeIns,FDShdTransDate,FTSrcRetDocRef,FTSrcCardNo,FTTdmCode,FTTdmType,FTSrcAbb FROM " & tVB_CNTblRC & tWhe
                                Call SP_SQLvExecute(tSql, orsTemp)
                            End Select
                            If Not (orsTemp.BOF Or orsTemp.EOF) Then
                                cWalletBalance = SP_FEDvChkNumeric(orsTemp, "FCSrcNet")
                                tTmpRmk = SP_FEDtChkString(orsTemp, "FTSrcRmk")
                                dSrcDate = SP_FEDdChkDate(orsTemp, "FDDateIns") & " " & TimeValue(SP_FEDtChkString(orsTemp, "FTTimeIns"))
                                dTransDate = SP_FEDdChkDate(orsTemp, "FDShdTransDate")
                                tTransID = SP_FEDtChkString(orsTemp, "FTSrcCardNo")
                                tVB_TenderCode = SP_FEDtChkString(orsTemp, "FTTdmCode")
                                Select Case tVB_TenderCode
                                Case "T019"
                                    tVB_TenderType = aVB_WalletType(0).TenderType
                                    tVB_TenderName = aVB_WalletType(0).TenderReturn
                                    If Not aVB_WalletType(0).StaReturn Then
                                        Call SP_MSGnShowing("tms_CN999048", nCS_Warning, , , , aVB_WalletType(0).TenderSale) 'ไม่อนุญาตคืนสินค้าด้วย Alipay
                                        tVB_TenderCode = Trim(UCase(.TextMatrix(.Row, 3)))
                                        tVB_TenderType = cUT.UT_STRtToken(Trim(.TextMatrix(.Row, 1)), ".", 2)
                                        tVB_TenderName = cUT.UT_STRtToken(Trim(.TextMatrix(.Row, 1)), ".", 2) & " Return"
                                        bVB_CNTenderAct = False
                                        otmForm.Enabled = True
                                        otmLock.Enabled = True
                                        Call SP_DATxRsNothing(orsTemp)
                                        Call SP_OBJxSetFocus(ogdFunc)
                                        Exit Sub
                                    End If
                                Case "T020"
                                    tVB_TenderType = aVB_WalletType(1).TenderType
                                    tVB_TenderName = aVB_WalletType(1).TenderReturn
                                    If Not aVB_WalletType(1).StaReturn Then
                                        Call SP_MSGnShowing("tms_CN999048", nCS_Warning, , , , aVB_WalletType(1).TenderSale) 'ไม่อนุญาตคืนสินค้าด้วย WeChatpay
                                        tVB_TenderCode = Trim(UCase(.TextMatrix(.Row, 3)))
                                        tVB_TenderType = cUT.UT_STRtToken(Trim(.TextMatrix(.Row, 1)), ".", 2)
                                        tVB_TenderName = cUT.UT_STRtToken(Trim(.TextMatrix(.Row, 1)), ".", 2) & " Return"
                                        bVB_CNTenderAct = False
                                        otmForm.Enabled = True
                                        otmLock.Enabled = True
                                        Call SP_DATxRsNothing(orsTemp)
                                        Call SP_OBJxSetFocus(ogdFunc)
                                        Exit Sub
                                    End If
                                Case "T021"
                                    tVB_TenderType = aVB_WalletType(2).TenderType
                                    tVB_TenderName = aVB_WalletType(2).TenderReturn
                                    If Not aVB_WalletType(2).StaReturn Then
                                        Call SP_MSGnShowing("tms_CN999048", nCS_Warning, , , , aVB_WalletType(2).TenderSale) 'ไม่อนุญาตคืนสินค้าด้วย Bluepay
                                        tVB_TenderCode = Trim(UCase(.TextMatrix(.Row, 3)))
                                        tVB_TenderType = cUT.UT_STRtToken(Trim(.TextMatrix(.Row, 1)), ".", 2)
                                        tVB_TenderName = cUT.UT_STRtToken(Trim(.TextMatrix(.Row, 1)), ".", 2) & " Return"
                                        bVB_CNTenderAct = False
                                        otmForm.Enabled = True
                                        otmLock.Enabled = True
                                        Call SP_DATxRsNothing(orsTemp)
                                        Call SP_OBJxSetFocus(ogdFunc)
                                        Exit Sub
                                    End If
                                Case "T022"
                                    tVB_TenderType = aVB_WalletType(3).TenderType
                                    tVB_TenderName = aVB_WalletType(3).TenderReturn
                                    If Not aVB_WalletType(3).StaReturn Then
                                        Call SP_MSGnShowing("tms_CN999048", nCS_Warning, , , , aVB_WalletType(3).TenderSale) 'ไม่อนุญาตคืนสินค้าด้วย Others
                                        tVB_TenderCode = Trim(UCase(.TextMatrix(.Row, 3)))
                                        tVB_TenderType = cUT.UT_STRtToken(Trim(.TextMatrix(.Row, 1)), ".", 2)
                                        tVB_TenderName = cUT.UT_STRtToken(Trim(.TextMatrix(.Row, 1)), ".", 2) & " Return"
                                        bVB_CNTenderAct = False
                                        otmForm.Enabled = True
                                        otmLock.Enabled = True
                                        Call SP_DATxRsNothing(orsTemp)
                                        Call SP_OBJxSetFocus(ogdFunc)
                                        Exit Sub
                                    End If
                                End Select
                            Else
                                tVB_TenderCode = Trim(UCase(.TextMatrix(.Row, 3)))
                                tVB_TenderType = cUT.UT_STRtToken(Trim(.TextMatrix(.Row, 1)), ".", 2)
                                tVB_TenderName = cUT.UT_STRtToken(Trim(.TextMatrix(.Row, 1)), ".", 2) & " Refund"
                                bVB_Return = True
                                bVB_Reference = False
                            End If
                            Call SP_DATxRsNothing(orsTemp)
                            
                            If (bVB_Return) Then 'Return
                                If cWalletBalance > cVB_Balance Then cWalletBalance = cVB_Balance
                                If Not bVB_Reference Then
                                    cWalletBalance = cVB_Balance
                                    tTmpRmk = "" 'ไม่อ้างบิลขาย
                                    tTransID = ""
                                End If
                            End If
                                
    '                            'ถ้ายอดชำระ กับ ยอดคืน ไม่เท่ากัน จะไม่ยอมให้คืน
    '                            If cWalletBalance = cVB_Balance Then
    
                                    '*Tao 60-09-04 ทุก Tender ให้ Return Manual
                                    wFunctionKB.tW_StaSend = "N" 'Y=ส่ง , N=ไม่ได้ส่ง
                                    GoTo ManualReturn 'ไม่ต้องส่งไป WalletPay
                                    '====================================
                                    
                                    wFunctionKB.tW_No = ""
                                    wFunctionKB.tW_BnkCode = ""
                                    wFunctionKB.tW_BnkName = ""
                                    wFunctionKB.tW_AppCode = ""
                                    wFunctionKB.tW_Rmk = ""
                                    wFunctionKB.tW_TransID = ""
                                    wFunctionKB.tW_WPSTransID = ""
                                    wFunctionKB.tW_WalletName = ""
                                    wFunctionKB.tW_WalletTransID = ""
                                    wFunctionKB.tW_StaSend = ""
                                                                                                                                                
                                    Select Case SP_VALnText2Double(tVB_TransactionSub)
                                    Case tEN_TrnCancelVoucher 'ถ้า Cancel Voucher ให้ส่ง Cancel
                                        If (DateValue(dSrcDate) = Date) Then 'ถ้าคืนสินค้าวันเดียวกัน กับวันที่ชำระเงิน
                                            If (Now <= CDate(Date & " " & TimeValue(tVB_WalletAllowCancelTime))) Then 'ถ้ายังไม่ถึง 23:00:00
                                                wCNWalletPay.nW_WalletMode = 3 'Void
                                                wFunctionKB.tW_StaSend = "Y" 'Y=ส่ง , N=ไม่ได้ส่ง
                                            Else
                                                'wCNWalletPay.nW_WalletPayMode = 4 'Refund
                                                wFunctionKB.tW_StaSend = "N" 'Y=ส่ง , N=ไม่ได้ส่ง
                                                wFunctionKB.tW_TransID = tTransID
                                                cVB_Pay = cWalletBalance
                                                GoTo AutoReturn 'ไม่ต้องส่งไป WalletPay (คืนได้เลย)
                                                'If (Date > DateValue(dSrcDate)) Then GoTo ManualReturn 'ถ้าเกิน 23:00:00 หรือ ข้ามวันแล้ว ไม่ต้องส่งไป WalletPay (คืนได้เลย)
                                                'If (dVB_SaleDate > dTransDate) Then GoTo ManualReturn 'ถ้าเกิน 23:00:00 หรือ ข้ามวันแล้ว ไม่ต้องส่งไป WalletPay (คืนได้เลย)
                                            End If
                                        Else
                                            'wCNWalletPay.nW_WalletPayMode = 4 'Refund
                                            wFunctionKB.tW_StaSend = "N" '*Tao 60-05-24 Y=ส่ง , N=ไม่ได้ส่ง
                                            wFunctionKB.tW_TransID = tTransID
                                            cVB_Pay = cWalletBalance
                                            GoTo AutoReturn '*Tao 60-05-23 ไม่ต้องส่งไป WalletPay (คืนได้เลย)
                                            'If (Date > DateValue(dSrcDate)) Then GoTo ManualReturn '*Tao 60-05-18 ถ้าเกิน 23:00:00 หรือ ข้ามวันแล้ว ไม่ต้องส่งไป WalletPay (คืนได้เลย)
                                            'If (dVB_SaleDate > dTransDate) Then GoTo ManualReturn '*Tao 60-05-18 ถ้าเกิน 23:00:00 หรือ ข้ามวันแล้ว ไม่ต้องส่งไป WalletPay (คืนได้เลย)
                                        End If
                                    Case tEN_TrnDepositCancel, tEN_TrnAddDepositCancel  '*Tao 60-05-12 ถ้า Deposit Cancel ให้ส่ง Cancel
                                        If (DateValue(dSrcDate) = Date) Then 'ถ้าคืนสินค้าวันเดียวกัน กับวันที่ชำระเงิน
                                            If (Now <= CDate(Date & " " & TimeValue(tVB_WalletAllowCancelTime))) Then  'ถ้ายังไม่ถึง 23:00:00
                                                wCNWalletPay.nW_WalletMode = 3 'Void
                                                wFunctionKB.tW_StaSend = "Y" '*Tao 60-05-24 Y=ส่ง , N=ไม่ได้ส่ง
                                            Else
                                                'wCNAliPay.nW_AlipayMode = 4 'Refund
                                                wFunctionKB.tW_StaSend = "N" '*Tao 60-05-24 Y=ส่ง , N=ไม่ได้ส่ง
                                                wFunctionKB.tW_TransID = tTransID
                                                cVB_Pay = cWalletBalance
                                                GoTo AutoReturn '*Tao 60-05-23 ไม่ต้องส่งไป Alipay (คืนได้เลย)
                                                'If (Date > DateValue(dSrcDate)) Then GoTo ManualReturn '*Tao 60-05-18 ถ้าเกิน 23:00:00 หรือ ข้ามวันแล้ว ไม่ต้องส่งไป Alipay (คืนได้เลย)
                                                'If (dVB_SaleDate > dTransDate) Then GoTo ManualReturn '*Tao 60-05-18 ถ้าเกิน 23:00:00 หรือ ข้ามวันแล้ว ไม่ต้องส่งไป Alipay (คืนได้เลย)
                                            End If
                                        Else
                                            'wCNWalletPay.nW_WalletPayMode = 4 'Refund
                                            wFunctionKB.tW_StaSend = "N" '*Tao 60-05-24 Y=ส่ง , N=ไม่ได้ส่ง
                                            wFunctionKB.tW_TransID = tTransID
                                            cVB_Pay = cWalletBalance
                                            GoTo AutoReturn 'ไม่ต้องส่งไป WalletPay (คืนได้เลย)
                                            'If (Date > DateValue(dSrcDate)) Then GoTo ManualReturn '*Tao 60-05-18 ถ้าเกิน 23:00:00 หรือ ข้ามวันแล้ว ไม่ต้องส่งไป WalletPay (คืนได้เลย)
                                            'If (dVB_SaleDate > dTransDate) Then GoTo ManualReturn '*Tao 60-05-18 ถ้าเกิน 23:00:00 หรือ ข้ามวันแล้ว ไม่ต้องส่งไป WalletPay (คืนได้เลย)
                                        End If
                                    Case tEN_TrnIPVReturn  'IPV Return
                                        'wCNWalletPay.nW_WalletPayMode = 4 'Refund
                                        wFunctionKB.tW_StaSend = "N" 'Y=ส่ง , N=ไม่ได้ส่ง
                                        wFunctionKB.tW_TransID = tTransID
                                        cVB_Pay = cWalletBalance
                                        GoTo AutoReturn 'ไม่ต้องส่งไป WalletPay (คืนได้เลย)
                                    Case Else
                                        'wCNWalletPay.nW_WalletPayMode = 4 'Refund
                                        wFunctionKB.tW_StaSend = "N" 'Y=ส่ง , N=ไม่ได้ส่ง
                                        GoTo ManualReturn 'ไม่ต้องส่งไป WalletPay
                                        'If (Date > DateValue(dSrcDate)) Then GoTo ManualReturn 'ถ้าเกิน 23:00:00 หรือ ข้ามวันแล้ว ไม่ต้องส่งไป WalletPay (คืนได้เลย)
                                        'If (dVB_SaleDate > dTransDate) Then GoTo ManualReturn '23:00:00 หรือ ข้ามวันแล้ว ไม่ต้องส่งไป WalletPay (คืนได้เลย)
                                    End Select
                                    '===========================================
                                    
                                    If tTmpRmk = "" Then
                                        Call SP_MSGnShowing("tms_CN999033", nCS_Warning) 'ไม่มีข้อมูลจากบิลขายด้วย WalletPay
                                        bVB_CNTenderAct = False
                                        otmForm.Enabled = True
                                        otmLock.Enabled = True
                                        Call SP_DATxRsNothing(orsTemp)
                                        Call SP_OBJxSetFocus(ogdFunc)
                                        Exit Sub
                                    End If
                                                                                               
                                    wCNWalletPay.tW_InVoiceID = UT_tToken(tTmpRmk, ";", 1, True)
                                    wCNWalletPay.tW_RefDT = UT_tToken(tTmpRmk, ";", 2, True)
                                    wCNWalletPay.tW_RefPartnerTransID = UT_tToken(tTmpRmk, ";", 3, True)
                                    'wCNWalletPay.tW_RefAmount = SP_FMTcCurPoint(SP_VALcText2Double(UT_tToken(tTmpRmk, ";", 4, True)), True, True)
                                    'wCNWalletPay.tW_RefAmount = SP_FMTcCurPoint(cVB_Balance - cVB_RoundRcv - cVB_RoundChg, True, True)
                                    wCNWalletPay.tW_RefAmount = SP_FMTcCurPoint(cWalletBalance, True, True)
                                    wCNWalletPay.tW_RefBrand = UT_tToken(tTmpRmk, ";", 5, True)
                                    wCNWalletPay.tW_RefStoreID = UT_tToken(tTmpRmk, ";", 6, True)
                                    wCNWalletPay.tW_RefDeviceID = UT_tToken(tTmpRmk, ";", 7, True)
                                    wCNWalletPay.tW_RefWalletName = UT_tToken(tTmpRmk, ";", 8, True)
                                    wCNWalletPay.tW_RefWalletTransID = UT_tToken(tTmpRmk, ";", 9, True)
                                    
                                    wCNWalletPay.Show vbModal
                                    If (wFunctionKB.tW_Rmk = "") Then
                                        cVB_Pay = 0
                                    Else
                                        cVB_Pay = cWalletBalance
                                    End If
                                    
                                    If cVB_Pay = 0 Then
                                        wFunctionKB.tW_No = ""
                                        wFunctionKB.tW_BnkCode = ""
                                        wFunctionKB.tW_BnkName = ""
                                        wFunctionKB.tW_AppCode = ""
                                        wFunctionKB.tW_Rmk = ""
                                        wFunctionKB.tW_TransID = ""
                                        wFunctionKB.tW_WPSTransID = ""
                                        wFunctionKB.tW_WalletName = ""
                                        wFunctionKB.tW_WalletTransID = ""
                                        wFunctionKB.tW_StaSend = ""
                                        bVB_CNTenderAct = False
                                        otmForm.Enabled = True
                                        otmLock.Enabled = True
                                        Call SP_DATxRsNothing(orsTemp)
                                        Call SP_OBJxSetFocus(ogdFunc)
                                        Exit Sub
                                    End If
ManualReturn:
                                    'หากไม่ส่งไป WalletPay ให้ใส่ TransID ก่อน
                                    If wFunctionKB.tW_StaSend = "N" Then
                                        If tTransID <> "" Then
                                            aVB_WalletLog.PartnerTransID = tTransID
                                        Else
                                            aVB_WalletLog.PartnerTransID = IIf(tTmpRmk <> "", UT_tToken(tTmpRmk, ";", 3, True), "")
                                        End If
                                        aVB_WalletLog.ReqAmt = SP_FMTcCurPoint(cWalletBalance, True, True)
                                        wCNWalletReturn.Caption = tVB_TenderName
                                        wCNWalletReturn.Show vbModal
                                        If (Not bVB_CNReturnOk) Or (cVB_CNDiscAmt = 0) Then
                                            wFunctionKB.tW_No = ""
                                            wFunctionKB.tW_BnkCode = ""
                                            wFunctionKB.tW_BnkName = ""
                                            wFunctionKB.tW_AppCode = ""
                                            wFunctionKB.tW_Rmk = ""
                                            wFunctionKB.tW_TransID = ""
                                            wFunctionKB.tW_WPSTransID = ""
                                            wFunctionKB.tW_WalletName = ""
                                            wFunctionKB.tW_WalletTransID = ""
                                            wFunctionKB.tW_StaSend = ""
                                            bVB_CNTenderAct = False
                                            otmForm.Enabled = True
                                            otmLock.Enabled = True
                                            Call SP_DATxRsNothing(orsTemp)
                                            Call SP_OBJxSetFocus(ogdFunc)
                                            Exit Sub
                                        End If
                                        cVB_Pay = cVB_CNDiscAmt
                                        cVB_CNDiscAmt = 0
                                        '*Tao 60-09-06 เปลี่ยน Tender Type ตามที่ WPA ส่งมา
                                        Select Case UCase(wFunctionKB.tW_WalletName)
                                        Case UCase(aVB_WalletType(0).TenderType)
                                            tVB_TenderCode = aVB_WalletType(0).TenderCode
                                            tVB_TenderType = aVB_WalletType(0).TenderType
                                            tVB_TenderName = aVB_WalletType(0).TenderReturn
                                        Case UCase(aVB_WalletType(1).TenderType)
                                            tVB_TenderCode = aVB_WalletType(1).TenderCode
                                            tVB_TenderType = aVB_WalletType(1).TenderType
                                            tVB_TenderName = aVB_WalletType(1).TenderReturn
                                        Case UCase(aVB_WalletType(2).TenderType)
                                            tVB_TenderCode = aVB_WalletType(2).TenderCode
                                            tVB_TenderType = aVB_WalletType(2).TenderType
                                            tVB_TenderName = aVB_WalletType(2).TenderReturn
                                        Case UCase(aVB_WalletType(3).TenderType)
                                            tVB_TenderCode = aVB_WalletType(3).TenderCode
                                            tVB_TenderType = aVB_WalletType(3).TenderType
                                            tVB_TenderName = aVB_WalletType(3).TenderReturn
                                        Case Else
                                            tVB_TenderCode = aVB_WalletType(3).TenderCode
                                            tVB_TenderType = aVB_WalletType(3).TenderType
                                            tVB_TenderName = aVB_WalletType(3).TenderReturn
                                        End Select
                                    End If
                                    '============================================
                                    
AutoReturn:
                                    Call SP_OBJxSetFocus(ogdFunc)
                                    Call W_PRCxAddWalletPay
    '                            Else
    '                                Call SP_MSGnShowing("tms_CN005234", nCS_Warning) 'ต้องคืนเต็มจำนวน
    '                                Call SP_OBJxSetFocus(ogdFunc)
    '                            End If
                        Case tEN_TrnSale, tEN_TrnDeposit, tEN_TrnAddDeposit, tEN_TrnRedeem, tEN_TrnTaxExceptSale, tEN_TrnVIPSale, tEN_TrnVoucher, tEN_TrnDutyFreeSale, tEN_TrnSaleManualTax, tEN_TrnSaleTaxInc
                            'ถ้าเคยชำระด้วย Wallet แล้ว ไม่ให้ชำระด้วย Wallet ได้อีก
                            tWhe = " WHERE FTTmnNum='" & tVB_CNTerminalNum & "' AND FTShdTransNo='" & tVB_Running & "'"
                            tWhe = tWhe & " AND FTShdTransType='" & tVB_Transaction & "' AND FDShdTransDate='" & tUT_SaleDate & "'"
                            tWhe = tWhe & " AND FTTdmCode IN ('T019','T020','T021','T022')"
                            nTmp = SP_VALcText2Double(SP_SQLtFunction("COUNT", "FTShdTransNo", tVB_CNTblRC, tWhe))
                            If nTmp = 0 Then
                                 
                                olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance), False, True)
                                
                                tVB_TenderCode = Trim(UCase(.TextMatrix(.Row, 3)))
                                tVB_TenderType = cUT.UT_STRtToken(Trim(.TextMatrix(.Row, 1)), ".", 2)
                                tVB_TenderName = cUT.UT_STRtToken(Trim(.TextMatrix(.Row, 1)), ".", 2)
                                'ให้แสดงจำนวนเงินก่อน
                                wCNPayCash.Caption = tVB_TenderName
                                wCNPayCash.Show vbModal
                                If cVB_Pay = 0 And cVB_Balance <> 0 Then
                                    'Call SP_GETxBalanceRnd
                                    bVB_CNTenderAct = False
                                    otmForm.Enabled = True
                                    otmLock.Enabled = True
                                    Call SP_DATxRsNothing(orsTemp)
                                    Call SP_OBJxSetFocus(ogdFunc)
                                    Exit Sub
                                Else
                                    If cVB_Pay > (cVB_Balance) Then 'ชำระเกินจำนวน
                                        Call SP_MSGnShowing("tms_CN005066", nCS_Warning)
                                        bVB_CNTenderAct = False
                                        otmForm.Enabled = True
                                        otmLock.Enabled = True
                                        Call SP_DATxRsNothing(orsTemp)
                                        Call SP_OBJxSetFocus(ogdFunc)
                                        Exit Sub
                                    End If
                                End If
                                                                
                                wFunctionKB.tW_No = ""
                                wFunctionKB.tW_BnkCode = ""
                                wFunctionKB.tW_BnkName = ""
                                wFunctionKB.tW_AppCode = ""
                                wFunctionKB.tW_Rmk = ""
                                wFunctionKB.tW_TransID = ""
                                wFunctionKB.tW_WPSTransID = ""
                                wFunctionKB.tW_WalletName = ""
                                wFunctionKB.tW_WalletTransID = ""
                                wFunctionKB.tW_StaSend = ""
                                
                                wCNWalletPay.nW_WalletMode = 0 'Payment
                                wCNWalletPay.tW_InVoiceID = tVB_CNStoreCode & "-" & tVB_CNTerminalNum & tVB_Running
                                'wCNWalletPay.tW_Amount = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance - cVB_RoundRcv - cVB_RoundChg), True, True)
                                wCNWalletPay.tW_Amount = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Pay), True, True)
                                
                                wCNWalletPay.Show vbModal
'                                wFunctionKB.tW_Rmk = "SP11-0057000297;20170518145428;110000003476;12345.00;THEMALL;38;00001;201799999999999999999999999999"  'ทดสอบ อย่าลืมเอาออก
'                                wFunctionKB.tW_TransID = "110000003476"
'                                wFunctionKB.tW_WalletTransID = "201799999999999999999999999999"
                                If (wFunctionKB.tW_Rmk = "") Then cVB_Pay = 0
                                If cVB_Pay = 0 Then
                                    wFunctionKB.tW_No = ""
                                    wFunctionKB.tW_BnkCode = ""
                                    wFunctionKB.tW_BnkName = ""
                                    wFunctionKB.tW_AppCode = ""
                                    wFunctionKB.tW_Rmk = ""
                                    wFunctionKB.tW_TransID = ""
                                    wFunctionKB.tW_WPSTransID = ""
                                    wFunctionKB.tW_WalletName = ""
                                    wFunctionKB.tW_WalletTransID = ""
                                    wFunctionKB.tW_StaSend = ""
                                    olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance), False, True)
                                    bVB_CNTenderAct = False
                                    otmForm.Enabled = True
                                    otmLock.Enabled = True
                                    Call SP_DATxRsNothing(orsTemp)
                                    Call SP_OBJxSetFocus(ogdFunc)
                                    Exit Sub
                                Else
                                    If wFunctionKB.tW_WalletTransID <> "" Then '*Tao 60-09-06 ถ้าเป็นค่าว่าง แสดงว่า Pay OffLine
                                        wFunctionKB.tW_StaSend = "Y" 'Y=ส่ง , N=ไม่ได้ส่ง
                                        '*Tao 60-08-31 เปลี่ยน Tender Type ตามที่ WPA ส่งมา
                                        Select Case UCase(Left(wFunctionKB.tW_WalletName, 3))
                                        Case UCase(Left(aVB_WalletType(0).TenderType, 3))
                                            tVB_TenderCode = aVB_WalletType(0).TenderCode
                                            tVB_TenderType = aVB_WalletType(0).TenderType
                                            tVB_TenderName = aVB_WalletType(0).TenderSale
                                        Case UCase(Left(aVB_WalletType(1).TenderType, 3))
                                            tVB_TenderCode = aVB_WalletType(1).TenderCode
                                            tVB_TenderType = aVB_WalletType(1).TenderType
                                            tVB_TenderName = aVB_WalletType(1).TenderSale
                                        Case UCase(Left(aVB_WalletType(2).TenderType, 3))
                                            tVB_TenderCode = aVB_WalletType(2).TenderCode
                                            tVB_TenderType = aVB_WalletType(2).TenderType
                                            tVB_TenderName = aVB_WalletType(2).TenderSale
                                        Case Else
                                            tVB_TenderCode = aVB_WalletType(3).TenderCode
                                            tVB_TenderType = aVB_WalletType(3).TenderType
                                            tVB_TenderName = aVB_WalletType(3).TenderSale
                                        End Select
                                    Else
                                        wFunctionKB.tW_StaSend = "N" 'Y=ส่ง , N=ไม่ได้ส่ง
                                        '*Tao 60-09-07 เปลี่ยน Tender Type ตามที่เลือก
                                        Select Case UCase(wFunctionKB.tW_WalletName)
                                        Case UCase(aVB_WalletType(0).TenderType)
                                            tVB_TenderCode = aVB_WalletType(0).TenderCode
                                            tVB_TenderType = aVB_WalletType(0).TenderType
                                            tVB_TenderName = aVB_WalletType(0).TenderSale
                                        Case UCase(aVB_WalletType(1).TenderType)
                                            tVB_TenderCode = aVB_WalletType(1).TenderCode
                                            tVB_TenderType = aVB_WalletType(1).TenderType
                                            tVB_TenderName = aVB_WalletType(1).TenderSale
                                        Case UCase(aVB_WalletType(2).TenderType)
                                            tVB_TenderCode = aVB_WalletType(2).TenderCode
                                            tVB_TenderType = aVB_WalletType(2).TenderType
                                            tVB_TenderName = aVB_WalletType(2).TenderSale
                                        Case UCase(aVB_WalletType(3).TenderType)
                                            tVB_TenderCode = aVB_WalletType(3).TenderCode
                                            tVB_TenderType = aVB_WalletType(3).TenderType
                                            tVB_TenderName = aVB_WalletType(3).TenderSale
                                        Case Else
                                            tVB_TenderCode = aVB_WalletType(3).TenderCode
                                            tVB_TenderType = aVB_WalletType(3).TenderType
                                            tVB_TenderName = aVB_WalletType(3).TenderSale
                                        End Select
                                    End If
                                End If
                                Call SP_OBJxSetFocus(ogdFunc)
                                Call W_PRCxAddWalletPay
                            Else
                                'Call SP_MSGnShowing("tms_CN999046", nCS_Warning) 'มีการชำระประเภทอื่นแล้ว
                                Call SP_MSGnShowing("tms_CN999038", nCS_Warning) 'มีการชำระด้วย WalletPay แล้ว
                                Call SP_OBJxSetFocus(ogdFunc)
                            End If
                        Case Else
                                Call SP_MSGnShowing("tms_CN999037", nCS_Warning) 'ไม่อนุญาตให้ชำระด้วย WalletPay
                                Call SP_OBJxSetFocus(ogdFunc)
                        End Select
                        bVB_CNTenderAct = False
                        If cVB_Balance = 0 Then
                            Call SP_DATxRsNothing(orsTemp)
                            Unload Me
                            Exit Sub
                        Else
                            otmForm.Enabled = True
                            otmLock.Enabled = True
                        End If
            End Select
        End With
        If Not otmForm.Enabled Then otmForm.Enabled = True
        If Not otmLock.Enabled Then otmLock.Enabled = True
        Call SP_DATxRsNothing(orsTemp)
        Exit Sub
        
ErrHandle:
        If Not otmForm.Enabled Then otmForm.Enabled = True
        If Not otmLock.Enabled Then otmLock.Enabled = True
        Call SP_CRTxLogPrg("ErrorNumber:" & Err.Number & " ErrorDesc:" & Err.Description & " Form:" & Me.Name & " Control:ocmOk" & " Events:ocmOk_Click")
End Sub

'*KT 54-10-01 V.4.0.1 สามารถกดเมนูตัวเลขจาก Keyboad
'Private Sub ogdFunc_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
'     tW_HotKey = ""
'End Sub

Private Sub W_PRCxAddPay()
On Error GoTo ErrHandle
    Call CallByName(wFunctionKB, "W_KBRcvCash", VbMethod)
'    DoEvents '*KT 53-01-14 DoEvents
    Exit Sub
ErrHandle:
End Sub

Private Sub W_PRCxAddPayCrdCard()
On Error GoTo ErrHandle
    Call CallByName(wFunctionKB, "W_KBRcvCredit", VbMethod)
'    DoEvents '*KT 53-01-14 DoEvents
    Exit Sub
ErrHandle:
End Sub

Private Sub W_PRCxAddWalletPay()
'*Tao 60-08-25 Wallet Payment
On Error GoTo ErrHandle
    Call CallByName(wFunctionKB, "W_KBRcvWalletPay", VbMethod)
'    DoEvents '*KT 53-01-14 DoEvents
    Exit Sub
ErrHandle:
End Sub

Private Sub W_PRCxAddPayGiftVoucher()
On Error GoTo ErrHandle
    Call CallByName(wFunctionKB, "W_KBRcvGiftVoucher", VbMethod)
'    DoEvents '*KT 53-01-14 DoEvents
    Exit Sub
ErrHandle:
End Sub

Private Sub W_PRCxAddPaySpcGiftVoucher()
On Error GoTo ErrHandle
    Call CallByName(wFunctionKB, "W_KBRcvSpcGiftVoucher", VbMethod)
'    DoEvents '*KT 53-01-14 DoEvents
    Exit Sub
ErrHandle:
End Sub

Private Sub W_PRCxAddPayVendorCpn()
On Error GoTo ErrHandle
    Call CallByName(wFunctionKB, "W_KBRcvVendorCpn", VbMethod)
'    DoEvents '*KT 53-01-14 DoEvents
    Exit Sub
ErrHandle:
End Sub

Private Sub W_PRCxAddPayVoucher()
On Error GoTo ErrHandle
    Call CallByName(wFunctionKB, "W_KBRcvVoucher", VbMethod)
'    DoEvents '*KT 53-01-14 DoEvents
    Exit Sub
ErrHandle:
End Sub

Private Sub W_PRCxAddPaySpecialCpn()
On Error GoTo ErrHandle
    Call CallByName(wFunctionKB, "W_KBRcvSpecialCpn", VbMethod)
'    DoEvents '*KT 53-01-14 DoEvents
    Exit Sub
ErrHandle:
End Sub

Private Sub W_PRCxAddPayCrdStore()
On Error GoTo ErrHandle
    Call CallByName(wFunctionKB, "W_KBRcvCrdStore", VbMethod)
'    DoEvents '*KT 53-01-14 DoEvents
    Exit Sub
ErrHandle:
End Sub

Private Sub W_PRCxAddPayStoreCoupon()
On Error GoTo ErrHandle
    Call CallByName(wFunctionKB, "W_KBRcvStoreCpn", VbMethod)
'    DoEvents '*KT 53-01-14 DoEvents
    Exit Sub
ErrHandle:
End Sub

Private Sub W_PRCxAddPayCreditNote()
On Error GoTo ErrHandle
    Call CallByName(wFunctionKB, "W_KBRcvCreditNote", VbMethod)
'    DoEvents '*KT 53-01-14 DoEvents
    Exit Sub
ErrHandle:
End Sub

Private Sub W_PRCxAddPayForeignCur()
On Error GoTo ErrHandle
    Call CallByName(wFunctionKB, "W_KBRcvForeignCur", VbMethod)
'    DoEvents '*KT 53-01-14 DoEvents
    Exit Sub
ErrHandle:
End Sub

Private Sub W_PRCxAddPayCV()
On Error GoTo ErrHandle
    Call CallByName(wFunctionKB, "W_KBRcvCV", VbMethod)
'    DoEvents '*KT 53-01-14 DoEvents
    Exit Sub
ErrHandle:
End Sub

Public Sub W_SetGridHeading()
'----------------------------------------------------------
'   Call:
'   Cmt:   set heading of tender grid
'----------------------------------------------------------
    Dim tColW$, iCol%
    Dim tFed As String
On Error GoTo ErrHandle
    tFed = "FNSrcSeqNo,FTTdmType,FCSrcNet"
'    tColW = "705;5295;1980"      '<= tFed, ส่วน tFed จะมีเท่าไหร่ก็ได้ แต่ห้ามเกินจำนวน cols ของ grid
    If bVB_Res640X480 Then
        tColW = "555;4320;1005"
    Else
        tColW = "705;5295;1980"
    End If
    iCol = 3     'last col to show, if > iCol จะถูก hidden
    With ogdTender
        Call SP_GRDxSetProperty(ogdTender, tFed, tColW, iCol, True, 0)
        .ColHidden(.ColIndex("FNSrcSeqNo")) = False
        .ColHidden(.ColIndex("FTTdmType")) = False
        .ColHidden(.ColIndex("FCSrcNet")) = False
    End With
    Exit Sub
ErrHandle:
    
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

Private Sub ogdFunc_DblClick()
    Call ocmOK_Click
End Sub

Private Sub otmForm_Timer()
'    Dim nMaxRow As Long
'    nMaxRow = ogdFunc.Aggregate(flexSTMax, 0, 0, ogdFunc.rows - 1, 0)
    If IsNumeric(tW_HotKey) Then tW_HotKey = "": otmForm.Enabled = False
End Sub

Private Sub W_PRCxClearRC()
'------------------------------------------
'Call :
'Cmt : ทำการ Clear ข้อมูลการชำระเงินทุกครั้งที่ load หน้าจอนี้ ก่อนการรับชำระใหม่
'--------------------------------------------
    Dim tSql$, tWhr$
On Error GoTo ErrHandle
    tWhr = " WHERE  FTTmnNum=''" & tVB_CNTerminalNum & "'' AND FTShdTransNo=''" & tVB_Running & "''"
    tWhr = tWhr & " AND  FTShdTransType=''" & tVB_Transaction & "'' AND FDShdTransDate=''" & tUT_SaleDate & "''"
    tWhr = tWhr & " AND  FTTdmCode <> ''T006''" 'ยกเว้นกรณีการรับชำระด้วย Vendor Coupon ที่รับชำระที่หน้าจอ Items Emptry    '*Ao 53-10-15
    tSql = "STP_CNxSQLFunctionDel @ptTable = '" & tVB_CNTblRC & "',@ptWhere = '" & tWhr & "'"
    Call SP_SQLvExecute(tSql)
    Exit Sub
ErrHandle:
    
End Sub

Private Sub otmLock_Timer()
    '*KT 51-10-31
    Dim bOnOff As Boolean
    Dim bOnOffCur As Boolean
    Dim orsTemp As ADODB.Recordset
    Dim tSql As String
    Dim tErr$
    Dim tSta$
    Dim tStaOk$
On Error GoTo ErrHandle '*KT 52-01-09 ใส่ on error เพื่อดูว่าเกิด Error Number 91 ซึ่งส่วนใหญ่เกิดจากการลง SQLEXPRESS 2005 ไม่สมบูรณ์เมื่อมาทำงานบน Offline
    '*KT 51-02-27 ทำการ Update Ststusbar หน้าจอหลัก
     
    '*KT 65-08-07 0 Block timer interup process
    Select Case UCase(Screen.ActiveForm.Name)
        Case UCase("wCNSubPay640")
            'none
        Case Else
            otmLock.Enabled = False
            Exit Sub
    End Select
    
    bOnOff = IIf(wMain.olaOnOff.Caption = "On-line", True, False)
    bOnOffCur = oDB.SQLbAvailableServer(oDB.OffServerName, oDB.OnComIP)
    If bOnOffCur Then
        wMain.oimOnOff.Picture = wMain.oilOnOff.ListImages(1).Picture
        wMain.olaOnOff.Caption = "On-line"
        'move file ej offline to path online
        If Not bOnOff And bOnOffCur Then
            Call SaveSetting("AdaPosMall", "POSFront", "SentSale", "1")
            Call SP_MOVxEJToOnline
            Call SP_CHKbOnlineEnable 'PH 2.0.0
            Call SP_TMNxUpdateOpenOnline
        End If
    Else
        wMain.oimOnOff.Picture = wMain.oilOnOff.ListImages(2).Picture
        wMain.olaOnOff.Caption = "Off-line"
    End If
     tSta = GetSetting("AdaPosMall", "POSFront", "RepOK")
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
'    If bOnOffCur And bOnOff = False And oDB.IsReExecute = False Then
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
        Call SP_CRTxLogPrg("ErrorNumber:" & Err.Number & " ErrorDesc:" & Err.Description & " Form:" & Me.Name & " Control:otmLock" & " Events:Timer")
    End If
End Sub

Private Sub W_GETxCountOntopDeposit()
'--------------------------------------------
'Call:
'Cmt: ทำการเก็บข้อมูลการรูดบัตรเครดิตแล้วได้ Ontop ของ Deposit หรือ Add Deposit ก่อนหน้า
'---------------------------------------------
    Dim tSql$
    Dim orsOntop As ADODB.Recordset
On Error GoTo ErrHandle
    tSql = "SELECT " & tVB_CNTblCD & ".FTScdCardID FROM " & tVB_CNTblCD
    tSql = tSql & vbCrLf & "WHERE  FTTmnNum='" & tVB_CNTerminalNum & "' AND FTShdTransNo='" & tVB_Running & "'"
    tSql = tSql & vbCrLf & "AND  FTShdTransType='" & tVB_Transaction & "' AND FDShdTransDate='" & tUT_SaleDate & "'"
    tSql = tSql & vbCrLf & "AND  FNDctNo = " & nEN_Disc15Ontop
    tSql = tSql & vbCrLf & "GROUP BY FTScdCardID"
    
    Call SP_SQLvExecute(tSql, orsOntop)
    If Not (orsOntop.BOF Or orsOntop.EOF) Then
        Do While Not orsOntop.EOF
            nVB_OnTopCount = nVB_OnTopCount + 1 'Count การรูดบัตรเครดิต
            aVB_CreditCrd(UBound(aVB_CreditCrd)) = SP_FEDtChkString(orsOntop, "FTScdCardID")
            ReDim Preserve aVB_CreditCrd(UBound(aVB_CreditCrd) + 1) As String
            orsOntop.MoveNext
        Loop
    End If
    Call SP_DATxRsNothing(orsOntop)
    Exit Sub
ErrHandle:
    Call SP_MSGxShowError(Err.Number, Err.Description)
    Call SP_DATxRsNothing(orsOntop)
End Sub

'55-07-20 Neung TK
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

Private Function W_CHKbTenderEmpDicsFlag() As Boolean
    On Error GoTo HandleError
        If aVB_TenderMnu(nVB_CurTenderSeqNo).StaEmpDisc Then
            W_CHKbTenderEmpDicsFlag = True
        Else
            If cVB_CNDiscEmpAmt > 0 Then
                W_CHKbTenderEmpDicsFlag = False
                 Call SP_MSGnShowing("tms_CN005234", nCS_Warning)
            Else
                W_CHKbTenderEmpDicsFlag = True
            End If
        End If
    Exit Function
HandleError:
    W_CHKbTenderEmpDicsFlag = False
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Function

'*Eaw 56-07-02 CommSheet TK-ISS3000-249-TR.xls ไม่ต้อง Recovery RC
'Private Function W_GETxBalRound()
''--------------------------------------------
''Call:
''Cmt: '*Eaw 56-01-04   เปลี่ยนค่าให้ olaBalance.Caption กรณีที่มีการ Rounding
''---------------------------------------------
'    Dim tWhe$
'    Dim tSql$
'    Dim cMeaningDef As Double
'    Dim orsTemp As ADODB.Recordset
'    Dim ni As Long
'On Error GoTo HandleError
'    tSql = "SELECT FNMenMin,FNMenValue FROM TSysMeaning"
'    Call SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine, , , , , True, True) 'read master
'    If Not (orsTemp.BOF Or orsTemp.EOF) Then
'        cMeaningDef = CDbl(Val(Format(SP_FEDvChkNumeric(orsTemp, "FNMenValue"))) / 100)
'    Else
'        cMeaningDef = 0
'    End If
'    Call SP_DATxRsNothing(orsTemp)
'    If bVB_CNAlwRnd Then
'        Select Case nVB_CNTypeRnd
'            Case eRndUp
''                If SP_FMTcCurPoint(cVB_Balance - cVB_RoundRcv, True, False) < 1 Then
'                If SP_FMTcCurPoint(cVB_Balance - cVB_RoundRcv, True, False) < 0.05 Then  '*Eaw 56-07-01 CommSheet TK-ISS3000-250-TR.xls
'                    bVB_CNPayCrdNoSuccess = True
'                    olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance - cVB_RoundRcv), False, True) 'แสดง Balance ที่มีการบวกเศษด้วย เพราะการจ่ายด้วยเครดิตสามารถจ่ายเป็นเศษได้
'                Else
'                    bVB_CNPayCrdNoSuccess = False
'                End If
'            Case eRndDn
''                If SP_FMTcCurPoint(cVB_Balance - cVB_RoundRcv, True, False) < 1 Then
'                If SP_FMTcCurPoint(cVB_Balance - cVB_RoundRcv, True, False) < 0.05 Then '*Eaw 56-07-01 CommSheet TK-ISS3000-250-TR.xls
'                    bVB_CNPayCrdNoSuccess = True
'                    olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance - cVB_RoundRcv), False, True) 'แสดง Balance ที่มีการบวกเศษด้วย เพราะการจ่ายด้วยเครดิตสามารถจ่ายเป็นเศษได้
'                Else
'                    bVB_CNPayCrdNoSuccess = False
'                End If
'            Case eRndMeaning
'                If SP_FMTcCurPoint(cVB_Balance - cVB_RoundRcv, True, False) <= cMeaningDef Then
'                    bVB_CNPayCrdNoSuccess = True
'                    olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance - cVB_RoundRcv), False, True) 'แสดง Balance ที่มีการบวกเศษด้วย เพราะการจ่ายด้วยเครดิตสามารถจ่ายเป็นเศษได้
'                Else
'                    bVB_CNPayCrdNoSuccess = False
'                End If
'        End Select
'    Else
''        If SP_FMTcCurPoint(cVB_Balance - cVB_RoundRcv, True, False) < 1 Then
'        If SP_FMTcCurPoint(cVB_Balance - cVB_RoundRcv, True, False) < 0.05 Then '*Eaw 56-07-01 CommSheet TK-ISS3000-250-TR.xls
'            bVB_CNPayCrdNoSuccess = True
'           olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance - cVB_RoundRcv), False, True) 'แสดง Balance ที่มีการบวกเศษด้วย เพราะการจ่ายด้วยเครดิตสามารถจ่ายเป็นเศษได้
'        Else
'            bVB_CNPayCrdNoSuccess = False
'        End If
'    End If
'HandleError:
'    Call SP_MSGxShowError(Err.Number, Err.Description)
'End Function
'
'Public Function W_PRCxRecTender()
''---------------------------------------------
''Call:
''Cmt: '*Eaw 56-02-06  กู้ข้อมูลการจ่ายเงินของบิลที่ยังไม่จบ
''--------------------------------------------
'Dim tSql$, tWhr$
'Dim orsTemp As ADODB.Recordset
'Dim orsTempHD As ADODB.Recordset
'Dim nSeqNo As Long
'Dim nRow As Long
'Dim cDctAmt As Double
'Dim cDepositAmt As Double
'Dim cShdRnd As Double
'Dim tUpd$    '*Eaw 56-04-10
'On Error GoTo HandleError
'    '*Eaw 56-04-10  CommSheet TK-ISS3000-129-IT2 00.00.01.xls
'    tUpd = " SET FDDateUpd = '" & UT_tStringDate(Date) & "',"
'    tUpd = tUpd & vbCrLf & " FTTimeUpd = '" & Time$ & "',"
'    tUpd = tUpd & vbCrLf & " FTWhoUpd = '" & tVB_CNUserName & "',"
'    tUpd = tUpd & vbCrLf & " FDDateIns = '" & UT_tStringDate(Date) & "',"
'    tUpd = tUpd & vbCrLf & " FTTimeIns = '" & Time$ & "',"
'    tUpd = tUpd & vbCrLf & " FTWhoIns = '" & tVB_CNUserName & "'"
'    '----------------------------------------------------------------------------------------------------------
'    tWhr = "  WHERE FTTmnNum='" & tVB_CNTerminalNum & "'"
'    tWhr = tWhr & "  AND FTShdTransNo='" & tVB_Running & "'"
'    tWhr = tWhr & "  AND FTShdTransType='" & tVB_Transaction & "'"
'    tWhr = tWhr & "  AND FDShdTransDate ='" & tUT_SaleDate & "'"
'    tSql = "SELECT *  FROM " & tVB_CNTblRCR
'    tSql = tSql & tWhr
'    Call SP_SQLvExecute(tSql, orsTemp)
'    tSql = "SELECT SUM(FCShdDepositAmt) AS  FCShdDepositAmt  , FCShdRnd FROM " & tVB_CNTblHD
'    tSql = tSql & tWhr
'    tSql = tSql & "GROUP BY FCShdRnd"
'    Call SP_SQLvExecute(tSql, orsTempHD)
'    orsTempHD.MoveFirst
'    cDepositAmt = SP_FMTcCurPoint(SP_FEDvChkNumeric(orsTempHD, "FCShdDepositAmt"), False, True)
'    cShdRnd = SP_FMTcCurPoint(SP_FEDvChkNumeric(orsTempHD, "FCShdRnd"), False, True)
'
'        Select Case Val(Format(tVB_TransactionSub))
'        Case tEN_TrnDeposit
'            If Not (orsTemp.EOF And orsTemp.BOF) Then
'                orsTemp.MoveFirst
'                With ogdTender
'                    'Clear grid
'                    .rows = .FixedRows
'                    .rows = 4
'                    .TextMatrix(1, 0) = "1"
'                    .TextMatrix(2, 0) = "2"
'                    .TextMatrix(3, 0) = "3"
'                    nRow = .FixedRows
'                    nSeqNo = 1
'                    cDctAmt = 0
'                    Do While Not orsTemp.EOF
''                            .TextMatrix(nRow, .ColIndex("FNSrcSeqNo")) = 1
'                        .TextMatrix(nRow, .ColIndex("FNSrcSeqNo")) = SP_FEDvChkNumeric(orsTemp, "FNSrcSeqNo")
'                        .TextMatrix(nRow, .ColIndex("FTTdmType")) = SP_FEDtChkString(orsTemp, "FTTdmType")
'                        .TextMatrix(nRow, .ColIndex("FCSrcNet")) = SP_FMTcCurPoint(SP_FEDvChkNumeric(orsTemp, "FCSrcNet"), False, True)
'                        cDctAmt = cDctAmt + SP_FMTcCurPoint(SP_FEDvChkNumeric(orsTemp, "FCSrcNet"), False, True)
'                        nRow = nRow + 1
'                        orsTemp.MoveNext
'                    Loop
'                    olaTotal.Caption = SP_FMTcCurPoint(cUT_SubTotal)
'                    olaDiscTotal.Caption = SP_FMTcCurPoint(cVB_TotalDiscAft)
'                    olaNetSale.Caption = SP_FMTcCurPoint(cUT_SubTotal - cVB_TotalDiscAft)
'                    lblDepositAmt.Caption = SP_FMTcCurPoint(cVB_DepositAmt)
'                    olaAmtTender.Caption = SP_FMTcCurPoint(cDctAmt)
'                    cVB_Balance = SP_FMTcCurPoint(cDepositAmt - cDctAmt)
'                    olaBalance.Caption = SP_FMTcCurPoint(cVB_Balance)
'
'                End With
'                tSql = "INSERT INTO  " & tVB_CNTblRC & " SELECT * FROM " & tVB_CNTblRCR
'                tSql = tSql & tWhr
'                Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
'                tSql = "DROP TABLE " & tVB_CNTblRCR
'                Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
'            End If
'        Case Else
'            If Not (orsTemp.EOF And orsTemp.BOF) Then
'                orsTemp.MoveFirst
'                With ogdTender
'                'Clear grid
'                    .rows = .FixedRows
'                    .rows = 4
'                    .TextMatrix(1, 0) = "1"
'                    .TextMatrix(2, 0) = "2"
'                    .TextMatrix(3, 0) = "3"
'                    nRow = .FixedRows
'                    nSeqNo = 1
'                    cDctAmt = 0
'                Do While Not orsTemp.EOF
'                    .TextMatrix(nRow, .ColIndex("FNSrcSeqNo")) = SP_FEDvChkNumeric(orsTemp, "FNSrcSeqNo")
'                    .TextMatrix(nRow, .ColIndex("FTTdmType")) = SP_FEDtChkString(orsTemp, "FTTdmType")
'                    .TextMatrix(nRow, .ColIndex("FCSrcNet")) = SP_FMTcCurPoint(SP_FEDvChkNumeric(orsTemp, "FCSrcNet"), False, True)
'                    cDctAmt = cDctAmt + SP_FMTcCurPoint(SP_FEDvChkNumeric(orsTemp, "FCSrcNet"), False, True)
'                    nRow = nRow + 1
'                    orsTemp.MoveNext
'                Loop
'                olaTotal.Caption = SP_FMTcCurPoint(cUT_SubTotal)
'                olaDiscTotal.Caption = SP_FMTcCurPoint(cVB_TotalDiscAft)
'                olaNetSale.Caption = SP_FMTcCurPoint(cUT_SubTotal - cVB_TotalDiscAft)
'                lblDepositAmt.Caption = SP_FMTcCurPoint(cVB_DepositAmt)
'                olaAmtTender.Caption = SP_FMTcCurPoint(cDctAmt)
'                cVB_Balance = SP_FMTcCurPoint(cUT_SubTotal - cVB_TotalDiscAft - cDctAmt + cShdRnd)
'                olaBalance.Caption = SP_FMTcCurPoint(cVB_Balance)
''                cUT_SubTotal = cUT_SubTotal - cVB_TotalDiscAft      '*Eaw 56-01-18 Test การเปลี่ยน FCShdGrand
'                Call W_PRCxSubRegHDEx
'                cUT_SubTotal = cUT_SubTotal - cVB_TotalDiscAft      '*Eaw 56-01-18 Test การเปลี่ยน FCShdGrand
'                Call W_UPDxTransHD  '*Eaw 56-01-18 Test การเปลี่ยน FCShdGrand
'                End With
'
'            tSql = "INSERT INTO  " & tVB_CNTblRC & " SELECT * FROM " & tVB_CNTblRCR
'            tSql = tSql & tWhr
'            Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
'            tSql = "DROP TABLE " & tVB_CNTblRCR
'            Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
'
'            '*Eaw 56-04-10  CommSheet TK-ISS3000-129-IT2 00.00.01.xls
'            tSql = "UPDATE " & tVB_CNTblHD & tUpd & tWhr
'            Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
'            tSql = "UPDATE " & tVB_CNTblRC & tUpd & tWhr
'            Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
'            '--------------------------------------------------------------------------------
'        End If
'        End Select
'        Exit Function
'HandleError:
'    Call SP_MSGxShowError(Err.Number, Err.Description)
'End Function
Public Sub W_UPDxTransHD()
'---------------------------------------------
'Call:
'Cmt: '*Eaw 56-02-06  ปรับปรุงข้อมูล Transaction HD
'--------------------------------------------
    Dim tSql$, tValue$, tWhr$
On Error GoTo ErrHandle


    tValue = " SET FCShdTotal =" & SP_FMTcCurPoint(cUT_SubTotal, True, False)
    tValue = tValue & vbCrLf & " , FCShdGrand=" & SP_FMTcCurPoint(cUT_SubTotal, True, False)

    tWhr = " WHERE  FTTmnNum=''" & tVB_CNTerminalNum & "''"
    tWhr = tWhr & vbCrLf & "  AND FTShdTransNo=''" & tVB_Running & "''"
    tWhr = tWhr & vbCrLf & "  AND FTShdTransType=''" & tVB_Transaction & "''"
    tWhr = tWhr & vbCrLf & "  AND FDShdTransDate =''" & tUT_SaleDate & "''"
    tWhr = tWhr & vbCrLf & "  AND FTShdStaDoc = ''2''"

    tSql = "STP_CNxSQLFunctionUpd @ptValue='" & tValue & "',@ptTable='" & tVB_CNTblHD & "',@ptWhere='" & tWhr & "'"
    Call SP_SQLbExecute(tSql)
    Exit Sub
ErrHandle:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Sub

Public Sub W_PRCxSubRegHDEx(Optional ByVal pbSpc As Boolean = False)
'---------------------------------------------
'Call:
'Cmt: '*Eaw 56-02-06  ทำการ Update SubReg,SubPro ดึงค่าจาก FCSdtBluePrice,FCSdtYellowPrice Table DT
'--------------------------------------------

    Dim orsTemp As ADODB.Recordset
    Dim tSql As String
    Dim cSpcAmt As Double
    Dim tWhe$
    Dim cSpc As Double
    Dim cSumYellow As Double
    Dim cSumBlue As Double
    Dim cDiscMallReg  As Double
    Dim cDiscMallPro  As Double
    Dim cDiscontop As Double
    Dim tWhr$
    Dim cTrnDisReg As Double
    Dim cTrnDisPro As Double
On Error GoTo ErrHandle

    '*TW 51-08-08
    'ตั้งสมมุติฐานว่า cUT_SubReg ถูกต้องเพียงแต่ยังไม่บวกส่วนลดที่เกิดจาก Mall Card กับ Emp Card เท่านั้น
    'cVB_CNAddSubTotal เป็นยอดที่หายไป
    Select Case SP_VALnText2Double(tVB_TransactionSub)
        Case tEN_TrnDepositCancel, tEN_TrnAddDepositCancel
            'none ไม่ทำงานไม่นั้น Sub Reg ของบิลจะผิด
        Case Else
            Call oDisc.DISxGetSubRegProForBill(cUT_SubReg, cUT_SubPro)
            
            cDiscMallReg = W_PRCcGetDiscCrd(True) '*TW 51-09-03
            cDiscMallPro = W_PRCcGetDiscCrd(False) '*TW 51-09-03
            
            tWhr = " WHERE  FTTmnNum='" & tVB_CNTerminalNum & "' AND FTShdTransNo='" & tVB_Running & "'"
            tWhr = tWhr & " AND  FTShdTransType='" & tVB_Transaction & "' AND FDShdTransDate='" & tUT_SaleDate & "'"
            tWhr = tWhr & " AND  FNDctNo  IN (" & nEN_Disc15Ontop & ")"
            cDiscontop = SP_FMTcCurPoint(CDbl(Val(Format(SP_SQLtFunction("SUM", "FCScdAmt", tVB_CNTblCD, tWhr)))), True, False)
        
            
            '*KT 51-09-24 ML-SF-563 Sub Total Reg incorrect with Mall Card and Tran Markdown
            '*KT 54-11-30 P.4.0 V.4.0.7 Req.18 Add 5 Types of TransMarkdow
            '------------------------------------------------------------------------------------------------------------------
            tSql = "SELECT SUM(CASE WHEN DT.FTSkuStaMallEmpCard = 'Y' THEN  CD.FCScdAmt  ELSE 0 END ) AS FCTrnDiscReg"
            tSql = tSql & vbCrLf & ",SUM(CASE WHEN DT.FTSkuStaMallEmpCard = 'Y' THEN  0  ELSE CD.FCScdAmt  END ) AS FCTrnDiscPro"
            tSql = tSql & vbCrLf & "FROM " & tVB_CNTblDT & " DT INNER JOIN"
            tSql = tSql & vbCrLf & tVB_CNTblCD & " CD ON DT.FTTmnNum = CD.FTTmnNum AND"
            tSql = tSql & vbCrLf & "DT.FTShdTransNo = CD.FTShdTransNo AND DT.FNSdtSeqNo = CD.FNSdtSeqNo AND"
            tSql = tSql & vbCrLf & "DT.FDShdTransDate = CD.FDShdTransDate"
            tSql = tSql & vbCrLf & "WHERE  DT.FTTmnNum='" & tVB_CNTerminalNum & "' AND DT.FTShdTransNo='" & tVB_Running & "'"
            tSql = tSql & vbCrLf & "AND  DT.FTShdTransType='" & tVB_Transaction & "' AND DT.FDShdTransDate='" & tUT_SaleDate & "'"
'            tSql = tSql & vbCrLf & "AND ISNULL(DT.FTSdtCancelFlag,'N') <> 'Y'"             '*Ao 2013/10/02
            tSql = tSql & vbCrLf & "AND ISNULL(DT.FTSdtCancelFlag,'0') <> '1'"
            tSql = tSql & vbCrLf & "AND CD.FNDctNo IN (13,14,23,24,25,26,27)"
            Call SP_SQLvExecute(tSql, orsTemp)
            If Not (orsTemp.BOF Or orsTemp.EOF) Then
                    cTrnDisReg = SP_FMTcCurPoint(SP_FEDvChkNumeric(orsTemp, "FCTrnDiscReg"), True, False)
                    cTrnDisPro = SP_FMTcCurPoint(SP_FEDvChkNumeric(orsTemp, "FCTrnDiscPro"), True, False)
            Else
                    cTrnDisReg = 0
                    cTrnDisPro = 0
            End If
            '-------------------------------------------------------------------------------------------------------------------
            '***KT 51-11-29 V 1.0.1 Phase 2 ข้อ 15
            'กรณีที่มีการลด Trans Disc และ TransMarkdown ก่อนหน้าไม่มีได้ส่วนลด Mall Card ให้แสดง Sub Reg แบบลดส่วนลด Trans Disc และ TransMarkdown
            'กรณีที่มีการลด Trans Disc และ TransMarkdown ก่อนหน้าได้ส่วนลด Mall Card ให้แสดง Sub Reg แบบไม่ลดส่วนลด Trans Disc และ TransMarkdown
            If cDiscMallReg > 0 Then
                    cUT_SubReg = SP_FMTcCurPoint(cUT_SubReg + cDiscMallReg + cTrnDisReg, True, False)
                    cUT_SubTotal = SP_FMTcCurPoint(cUT_SubReg + cUT_SubPro + cDiscontop - cVB_CNAddSubTotal + cDiscMallPro - cTrnDisReg, True, False)
                    cUT_SubPro = SP_FMTcCurPoint(cUT_SubTotal - cUT_SubReg, True, False)
            Else
                    cUT_SubReg = SP_FMTcCurPoint(cUT_SubReg, True, False)
'                    cUT_SubTotal = SP_FMTcCurPoint(cUT_SubReg + cUT_SubPro + cDiscontop - cVB_CNAddSubTotal, True, False)
                    '*KT 52-02-10 V 1.5.8 ต้องทดสอบทุกๆๆเงื่อนไขที่ได้ส่วนลด
                    cUT_SubTotal = SP_FMTcCurPoint(cUT_SubReg + cUT_SubPro + cDiscontop - cVB_CNAddSubTotal + cDiscMallPro, True, False)
                    cUT_SubPro = SP_FMTcCurPoint(cUT_SubTotal - cUT_SubReg, True, False)
            End If
            '------------------------------------------------------------------------------------------------------------------
            
            tSql = "UPDATE " & tVB_CNTblHD
            tSql = tSql & vbCrLf & "SET FCShdSubReg = " & cUT_SubReg
            tSql = tSql & vbCrLf & ",FCShdSubPro = " & cUT_SubPro
            tSql = tSql & vbCrLf & ",FCShdTotal = " & cUT_SubTotal
            tSql = tSql & vbCrLf & "WHERE FTTmnNum = '" & tVB_CNTerminalNum & "'"
            tSql = tSql & vbCrLf & " AND FTShdTransNo = '" & tVB_Running & "'"
            tSql = tSql & vbCrLf & " AND FDShdTransDate = '" & tUT_SaleDate & "'"
            Call SP_SQLvExecute(tSql)
    End Select
    Call SP_DATxRsNothing(orsTemp)
    Exit Sub
ErrHandle:
    Call SP_MSGxShowError(Err.Number, Err.Description)
    Call SP_DATxRsNothing(orsTemp)
End Sub

Private Function W_PRCcGetDiscCrd(ByVal pbSubReg As Boolean) As Double
'---------------------------------------------
'Call:
'Cmt: '*Eaw 56-01-07
'--------------------------------------------

    Dim tWhr As String
    Dim tSql As String
    Dim orsTemp As ADODB.Recordset
    On Error GoTo Err_Handler
    tSql = "SELECT SUM(FCScdAmt) AS FCScdAmt"
    tSql = tSql & vbCrLf & "FROM " & tVB_CNTblCD & " CD LEFT JOIN " & tVB_CNTblDT & " DT ON CD.FTShdTransNo = DT.FTShdTransNo AND CD.FNSdtSeqNo = DT.FNSdtSeqNo"
    tSql = tSql & vbCrLf & "WHERE  CD.FTTmnNum='" & tVB_CNTerminalNum & "' AND CD.FTShdTransNo='" & tVB_Running & "'"
    tSql = tSql & vbCrLf & "AND  CD.FTShdTransType='" & tVB_Transaction & "' AND CD.FDShdTransDate='" & tUT_SaleDate & "'"
    tSql = tSql & vbCrLf & "AND  CD.FNDctNo  IN (" & nEN_Disc11MallCrd & "," & nEN_Disc12EmpCrd & ")"
    If pbSubReg Then
        tSql = tSql & vbCrLf & "AND  DT.FTSkuStaMallEmpCard = 'Y'"
    Else
        tSql = tSql & vbCrLf & "AND  DT.FTSkuStaMallEmpCard = 'N'"
    End If
    If SP_SQLbExecute(tSql, orsTemp) Then
        If Not (orsTemp.BOF And orsTemp.EOF) Then
            W_PRCcGetDiscCrd = SP_FEDvChkNumeric(orsTemp, "FCScdAmt")
        End If
    End If
    'cDiscMall = SP_FMTcCurPoint(Val(Format(SP_SQLtFunction("SUM", "FCScdAmt", tVB_CNTblCD, tWhr))), True, False)
    Call SP_DATxRsNothing(orsTemp)
    Exit Function
Err_Handler:
    Call SP_MSGxShowError(Err.Number, Err.Description)
    Call SP_DATxRsNothing(orsTemp)
End Function

Private Sub otmEDC_Timer()
'-------------------------------------------------
'
'   '*Ao 2014/04/29  '*Ao 2014/04/29    PRJ10001TK-PH2-External Spec-EDC Layaway Cancel-00.02.04 FSBT.docx
'-------------------------------------------------
On Error GoTo ErrHandle
    otmEDC.Enabled = False
    wCNDiscEmp.Show vbModal
    If bVB_ChkEDC = False Then                           '*Ao 2014/07/04
        Unload wSub
        'เริ่ม Ticket ใหม่
        tVB_CNTblHD = "TPSHD" & tVB_CNTerminalNum
        tVB_CNTblDT = "TPSDT" & tVB_CNTerminalNum
        tVB_CNTblCD = "TPSCD" & tVB_CNTerminalNum
        tVB_CNTblRC = "TPSRC" & tVB_CNTerminalNum
        Call UT_GenSaleNewDoc
        tVB_CNMnuCurrent = "1700"
        wMain.W_FRMxActivate
        Call SP_OBJxSetFocus(wMain.otbPrompt) '*KT 51-09-17 แก้ไขเรื่องการยิงสินค้าขึ้น Zero
        wMain.otbPrompt.Locked = False '*KT 51-09-17 แก้ไขเรื่องการยิงสินค้าขึ้น Zero
        bVB_CNStartBill = True '*KT 51-09-17 แก้ไขเรื่องการยิงสินค้าขึ้น Zero
        bVB_TextPromptAct = False
        Exit Sub
    End If
    Exit Sub
ErrHandle:
    otmEDC.Enabled = False
End Sub
