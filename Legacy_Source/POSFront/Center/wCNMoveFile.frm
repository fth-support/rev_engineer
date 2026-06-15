VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.1#0"; "Codejock.SuiteCtrls.Unicode.v11.1.3.ocx"
Begin VB.Form wCNMoveFile 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transfer Data"
   ClientHeight    =   4410
   ClientLeft      =   4200
   ClientTop       =   2070
   ClientWidth     =   6360
   ControlBox      =   0   'False
   Icon            =   "wCNMoveFile.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "wCNMoveFile.frx":058A
   ScaleHeight     =   4410
   ScaleWidth      =   6360
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   3735
      Index           =   2
      Left            =   0
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   -225
      Width           =   6345
      _Version        =   720897
      _ExtentX        =   11192
      _ExtentY        =   6588
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
      Begin VB.CommandButton ocmBrw 
         Caption         =   "..."
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
         Index           =   1
         Left            =   5535
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   1800
         Width           =   615
      End
      Begin VB.TextBox otbForm 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   360
         TabIndex        =   0
         Top             =   840
         Width           =   5160
      End
      Begin VB.TextBox otbForm 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   360
         TabIndex        =   1
         Top             =   1800
         Width           =   5160
      End
      Begin VB.CommandButton ocmBrw 
         Caption         =   "..."
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
         Index           =   0
         Left            =   5535
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   840
         Width           =   615
      End
      Begin MSComCtl2.Animation Animation1 
         Height          =   375
         Left            =   1200
         TabIndex        =   7
         Top             =   2760
         Visible         =   0   'False
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
         _Version        =   393216
         BackColor       =   12648447
         FullWidth       =   257
         FullHeight      =   25
      End
      Begin VB.Label oblCopy 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Copying......"
         Height          =   195
         Left            =   1200
         TabIndex        =   11
         Top             =   2400
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Source:;klg_Source"
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
         Left            =   360
         TabIndex        =   9
         Tag             =   "2;"
         Top             =   480
         Width           =   2460
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Destination:;klg_Dest"
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
         Index           =   1
         Left            =   360
         TabIndex        =   8
         Tag             =   "2;"
         Top             =   1440
         Width           =   2670
      End
   End
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   1200
      Index           =   1
      Left            =   0
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   3210
      Width           =   6345
      _Version        =   720897
      _ExtentX        =   11192
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
         Left            =   4680
         TabIndex        =   3
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
         Index           =   0
         Left            =   3120
         TabIndex        =   2
         TabStop         =   0   'False
         Tag             =   "2;"
         Top             =   480
         Width           =   1455
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid ogdIn 
      Height          =   1215
      Left            =   0
      TabIndex        =   10
      Top             =   4320
      Visible         =   0   'False
      Width           =   3330
      _cx             =   5874
      _cy             =   2143
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
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
      AllowUserResizing=   2
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"wCNMoveFile.frx":11CC
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
Attribute VB_Name = "wCNMoveFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const tCS_CNBchTF2 = "เลือกสถานที่เก็บแฟ้มข้อมูล ;Select directory path"
Private aW_TmnNum() As String
Private tW_TrnMin As String
Private tW_TrnMax As String
Private bW_CheckImp As Boolean

Private Sub W_PRCxUpdateHD()
'----------------------------------------------------------------------
'Call:
'Cmt:       '*Ao 2014/08/05
'----------------------------------------------------------------------
Dim tSql As String

On Error GoTo ErrHandle
        tSql = "UPDATE  TPSTSalHD"
        tSql = tSql & vbCrLf & " SET FTShdStaCancel ='Y' "
        tSql = tSql & vbCrLf & " WHERE  FTTmnNum IN (SELECT FTShdDepRefTmnNum  FROM  TPSTSalHD"
        tSql = tSql & vbCrLf & "                                            WHERE FTShdStaDoc = '1'"
        tSql = tSql & vbCrLf & "                                            AND FTShdTransType  IN  ('" & Format(tEN_TrnDepositCancel, "00") & "','" & Format(tEN_TrnAddDepositCancel, "00") & "')"
        tSql = tSql & vbCrLf & "                                           )"
        tSql = tSql & vbCrLf & " AND FTShdTransNo IN (SELECT  FTShdDepReTransNo FROM  TPSTSalHD"
        tSql = tSql & vbCrLf & "                                            WHERE FTShdStaDoc = '1'"
        tSql = tSql & vbCrLf & "                                            AND FTShdTransType  IN  ('" & Format(tEN_TrnDepositCancel, "00") & "','" & Format(tEN_TrnAddDepositCancel, "00") & "')"
        tSql = tSql & vbCrLf & "                                           )"
        tSql = tSql & vbCrLf & "AND FDShdTransDate IN (SELECT  FDShdDepReTransDate FROM  TPSTSalHD"
        tSql = tSql & vbCrLf & "                                            WHERE FTShdStaDoc = '1'"
        tSql = tSql & vbCrLf & "                                            AND FTShdTransType  IN  ('" & Format(tEN_TrnDepositCancel, "00") & "','" & Format(tEN_TrnAddDepositCancel, "00") & "')"
        tSql = tSql & vbCrLf & "                                           )"
        tSql = tSql & vbCrLf & "AND FTShdStaDoc = '1'"
        Call SP_SQLvExecute(tSql, , oDB.ocnOnLine) 'PH 2.0.0 connect database server
        Call SP_SQLvExecute(tSql) 'Call ocnAll.Execute(tSql)
    
        '*tEN_TrnIPV/tEN_TrnIPVReturn/tEN_TrnIPVRedeem/tEN_TrnIPVVIPService/tEN_TrnIPVTaxExceptSale/tEN_TrnIPVDutyFreeSale/tEN_TrnIPVOtherReceive/tEN_TrnIPVOtherDisbursement
        tSql = "UPDATE  TPSTSalHD"
        tSql = tSql & vbCrLf & " SET FTShdStaCancel ='Y' "
        tSql = tSql & vbCrLf & " WHERE  FTTmnNum IN (SELECT FTShdDepRefTmnNum FROM   TPSTSalHD"
        tSql = tSql & vbCrLf & "                                            WHERE FTShdStaDoc = '1'"
        tSql = tSql & vbCrLf & "                                            AND FTShdTransType  IN  ('" & Format(tEN_TrnIPV, "00") & "','" & Format(tEN_TrnIPVReturn, "00") & "'"
        tSql = tSql & vbCrLf & "                                                                                                    ,'" & Format(tEN_TrnIPVRedeem, "00") & "','" & Format(tEN_TrnIPVVIPService, "00") & "'"
        tSql = tSql & vbCrLf & "                                                                                                    ,'" & Format(tEN_TrnIPVTaxExceptSale, "00") & "','" & Format(tEN_TrnIPVDutyFreeSale, "00") & "'"
        tSql = tSql & vbCrLf & "                                                                                                    ,'" & Format(tEN_TrnIPVOtherReceive, "00") & "','" & Format(tEN_TrnIPVOtherDisbursement, "00") & "')"
        tSql = tSql & vbCrLf & "                                           )"
        tSql = tSql & vbCrLf & " AND FTShdTransNo IN (SELECT FTShdDepReTransNo FROM   TPSTSalHD"
        tSql = tSql & vbCrLf & "                                            WHERE FTShdStaDoc = '1'"
        tSql = tSql & vbCrLf & "                                            AND FTShdTransType  IN  ('" & Format(tEN_TrnIPV, "00") & "','" & Format(tEN_TrnIPVReturn, "00") & "'"
        tSql = tSql & vbCrLf & "                                                                                                    ,'" & Format(tEN_TrnIPVRedeem, "00") & "','" & Format(tEN_TrnIPVVIPService, "00") & "'"
        tSql = tSql & vbCrLf & "                                                                                                    ,'" & Format(tEN_TrnIPVTaxExceptSale, "00") & "','" & Format(tEN_TrnIPVDutyFreeSale, "00") & "'"
        tSql = tSql & vbCrLf & "                                                                                                    ,'" & Format(tEN_TrnIPVOtherReceive, "00") & "','" & Format(tEN_TrnIPVOtherDisbursement, "00") & "')"
        tSql = tSql & vbCrLf & "                                           )"
        tSql = tSql & vbCrLf & "AND FDShdTransDate = (SELECT FDShdDepReTransDate FROM   TPSTSalHD"
        tSql = tSql & vbCrLf & "                                            WHERE FTShdStaDoc = '1'"
        tSql = tSql & vbCrLf & "                                            AND FTShdTransType  IN  ('" & Format(tEN_TrnIPV, "00") & "','" & Format(tEN_TrnIPVReturn, "00") & "'"
        tSql = tSql & vbCrLf & "                                                                                                    ,'" & Format(tEN_TrnIPVRedeem, "00") & "','" & Format(tEN_TrnIPVVIPService, "00") & "'"
        tSql = tSql & vbCrLf & "                                                                                                    ,'" & Format(tEN_TrnIPVTaxExceptSale, "00") & "','" & Format(tEN_TrnIPVDutyFreeSale, "00") & "'"
        tSql = tSql & vbCrLf & "                                                                                                    ,'" & Format(tEN_TrnIPVOtherReceive, "00") & "','" & Format(tEN_TrnIPVOtherDisbursement, "00") & "')"
        tSql = tSql & vbCrLf & "                                           )"
        tSql = tSql & vbCrLf & "AND FTShdStaDoc = '1'"
        Call SP_SQLvExecute(tSql, , oDB.ocnOnLine) 'PH 2.0.0 connect database server
        Call SP_SQLvExecute(tSql)  'Call ocnAll.Execute(tSql)
                
        tSql = "UPDATE  TPSTSalHD"
        tSql = tSql & vbCrLf & " SET FTShdStaCancel ='Y' "
        tSql = tSql & vbCrLf & " WHERE  FTTmnNum IN (SELECT FTShdDepRefTmnNum FROM  TPSTSalHD"
        tSql = tSql & vbCrLf & "                                            WHERE FTShdStaDoc = '1'"
        tSql = tSql & vbCrLf & "                                            AND FTShdTransType  IN  ('" & Format(tEN_TrnCancelVoucher, "00") & "')"
        tSql = tSql & vbCrLf & "                                           )"
        tSql = tSql & vbCrLf & " AND FTShdTransNo IN (SELECT FTShdDepReTransNo FROM  TPSTSalHD"
        tSql = tSql & vbCrLf & "                                            WHERE FTShdStaDoc = '1'"
        tSql = tSql & vbCrLf & "                                            AND FTShdTransType  IN  ('" & Format(tEN_TrnCancelVoucher, "00") & "')"
        tSql = tSql & vbCrLf & "                                           )"
        tSql = tSql & vbCrLf & "AND FDShdTransDate = (SELECT FDShdDepReTransDate FROM  TPSTSalHD"
        tSql = tSql & vbCrLf & "                                            WHERE FTShdStaDoc = '1'"
        tSql = tSql & vbCrLf & "                                            AND FTShdTransType  IN  ('" & Format(tEN_TrnCancelVoucher, "00") & "')"
        tSql = tSql & vbCrLf & "                                           )"
        tSql = tSql & vbCrLf & "AND FTShdStaDoc = '1'"
        Call SP_SQLvExecute(tSql, , oDB.ocnOnLine) 'PH 2.0.0 connect database server
        Call SP_SQLvExecute(tSql)  'Call ocnAll.Execute(tSql)
        '--------------------------------------------------------------------
        Call SaveSetting("AdaPosMall", "POSFront", "SentSale", "1")     '*Ao 2013/11/22 ให้ส่งข้อมูล
Exit Sub
ErrHandle:

End Sub

Private Sub W_PRCxImportDAT(ByVal ptPthDrive As String, ByVal ptPthDat As String, ByVal ptFileType As String)
'----------------------------------------------------------------------
'Call:
'Cmt:       '*Ao 2014/08/05
'----------------------------------------------------------------------
Dim bChk As Boolean, bResults As Boolean
Dim tResults As String, tFolderName As String, tPthIn As String, tPthDateIn As String, tFileName As String
Dim tImport  As String

Dim oFldRootIn As Folder
Dim oSubFolder As Folder
Dim oFileIn As File

Dim oFSO As New FileSystemObject
On Error GoTo ErrHandle
    
    bChk = True
    Set oTransPrn.ActiveConnection = ocnAll
    nVB_SQLVersion = W_CHKxSQLVersion
    
    tImport = "Imported"
    If Not SP_FLDbManFolder(ptPthDrive & "\" & tImport, nEN_Fle1Exist) Then
        Call SP_FLDbManFolder(ptPthDrive & "\" & tImport, nEN_Fld3Create)
    End If
    
    Set oFldRootIn = oFSO.GetFolder(ptPthDrive)
    For Each oSubFolder In oFldRootIn.SubFolders
        tFolderName = oSubFolder.Name
        If Len(tFolderName) > 6 And (UCase(tFolderName) <> UCase(tImport)) Then
            For Each oFileIn In oSubFolder.Files
                oFileIn.Attributes = Normal                                                                                     'set file ไม่ให้ read only
                tPthIn = oFileIn.Path
                tPthDateIn = oFileIn.DateCreated
                tFileName = oFileIn.Name
                If UCase(Right(tPthIn, Len(ptFileType))) = UCase(ptFileType) Then
                    bW_CheckImp = True                  '*********************
                    oblCopy.Caption = "Import File:" & tFileName
                    DoEvents
                    If bChk Then
                        bChk = False
                        bResults = oTransPrn.Import(tPthIn, oDB.OffServerName, oDB.OffUsrName, oDB.OffUsrPwd, oDB.OffDbName)
                    Else
                        bResults = oTransPrn.Import(tPthIn)
                    End If
                    If bResults Then
                        tResults = "End"
                    Else
                        tResults = "Fail"
                    End If
                    Call SP_CRTxLogFile("Import File:" & tFileName & " [" & tResults & "]")
                    
                    If Not SP_FLDbManFolder(ptPthDrive & "\" & tImport & "\" & tFolderName, nEN_Fle1Exist) Then
                        Call SP_FLDbManFolder(ptPthDrive & "\" & tImport & "\" & tFolderName, nEN_Fld3Create)
                    End If
                    oblCopy.Caption = "Copying......" & tFileName
                    DoEvents
                    Call W_PROxCopyFile(tPthIn, ptPthDrive & "\" & tImport & "\" & tFolderName)
                    If SP_FLEbManFile(ptPthDrive & "\" & tImport & "\" & tFolderName & "\" & tFileName, nEN_Fle1Exist) Then
                        Call SP_FLEbManFile(tPthIn, nEN_Fle2Delete)
                    Else
                    
                    End If
                End If
            Next
            Call SP_FLDbManFolder(ptPthDrive & "\" & tFolderName, nEN_Fld2Delete)
        End If
    Next
    
    Set oFSO = Nothing
    oblCopy.Caption = "Copying......"
    DoEvents
    If bW_CheckImp Then
        Call W_PRCxUpdateHD
    End If
Exit Sub
ErrHandle:

End Sub

Private Sub W_FNDxCopyDatToSale(ByVal ptPthDrive As String, ByVal ptPthDat As String, ByVal ptFileType As String)
Dim oFldRootIn As Folder
Dim oSubFolder As Folder
Dim oFileIn As File
Dim tFolderName As String, tPthIn As String, tPthDateIn As String, tFileName As String
Dim nInFiles As Integer

Dim oFSO As New FileSystemObject

On Error GoTo ErrHandle
    nInFiles = 0
    
    Set oFldRootIn = oFSO.GetFolder(ptPthDrive)
    For Each oSubFolder In oFldRootIn.SubFolders
        tFolderName = oSubFolder.Name
        If Len(tFolderName) > 6 Then
            If Not SP_FLDbManFolder(ptPthDat & "\" & tFolderName, nEN_Fle1Exist) Then
                Call SP_FLDbManFolder(ptPthDat & "\" & tFolderName, nEN_Fld3Create)
            End If
            For Each oFileIn In oSubFolder.Files
                oFileIn.Attributes = Normal                                                                                     'set file ไม่ให้ read only
                tPthIn = oFileIn.Path
                tPthDateIn = oFileIn.DateCreated
                tFileName = oFileIn.Name
                If UCase(Right(tPthIn, Len(ptFileType))) = UCase(ptFileType) Then
                    nInFiles = nInFiles + 1
                    If nInFiles >= ogdIn.rows Then ogdIn.rows = ogdIn.rows + 1
                    ogdIn.TextMatrix(nInFiles, 0) = tPthIn
                    ogdIn.TextMatrix(nInFiles, 1) = tFileName
                    ogdIn.TextMatrix(nInFiles, 2) = tFolderName & "\" & tFileName
                    Call W_PROxCopyFile(tPthIn, ptPthDat & "\" & tFolderName)                                                             'Copy To Local
                    oblCopy.Caption = "Copying......" & tFileName
                End If
            Next
        End If
    Next
        
    Set oFSO = Nothing
    
Exit Sub
ErrHandle:
     Set oFSO = Nothing
End Sub

Private Sub W_FNDxCopyDatToDri(ByVal ptPthDat As String, ByVal ptPthDri As String, ByVal ptFileType As String)
'-------------------------------------------------------------------------------------------------
'Call:
'Cmt: '*Ao 2014/07/31    Copy DAT To Dri
'-------------------------------------------------------------------------------------------------
Dim tSql As String, tTransDate As String, tLogCode As String, tFolderDAT As String, tTmnNum As String, tTransNo As String
Dim tTrDate  As String, tFileName As String, tTemp As String
Dim tPathEJ  As String, tPathEJBK As String
Dim tDepID  As String, tStoreID As String, tTransDateRE As String, tDate As String, tFileEj As String
Dim dTransDate  As Date
            
Dim orsTemp As ADODB.Recordset

Dim oFSO As New FileSystemObject

On Error GoTo ErrHandle
    
    tSql = "SELECT FTShdDepID,FTStmCode,FTTmnNum,FTShdTransNo,FDShdTransDate,FTLogCode,FTCstCardCode3,FTShdTransType"
    tSql = tSql & vbCrLf & "FROM TPSTSalHD"
    tSql = tSql & vbCrLf & "WHERE ISNULL(FTStaSentOnOff,'0') <> '1' "
    tSql = tSql & vbCrLf & "ORDER BY FDShdTransDate,FTTmnNum,FTShdTransNo"
    If SP_SQLvExecute(tSql, orsTemp) = 0 Then
        If Not (orsTemp.BOF And orsTemp.EOF) Then
            orsTemp.MoveFirst
            Do While Not orsTemp.EOF
                tTransDate = SP_DTEtStrSTD(SP_FEDdChkDate(orsTemp, "FDShdTransDate"))               'yyyy/MM/dd
                tLogCode = SP_FEDtChkString(orsTemp, "FTLogCode")
                tTmnNum = SP_FEDtChkString(orsTemp, "FTTmnNum")
                tTransNo = SP_FEDtChkString(orsTemp, "FTShdTransNo")
                    
                tTrDate = Replace(tTransDate, "/", "-")                                                                                              'yyyy-MM-dd
                tFolderDAT = tTrDate & "_" & tLogCode                                                                                            '2014-07-21_14070006600017
                
                If SP_FLDbManFolder(ptPthDat & "\" & tFolderDAT, nEN_Fle1Exist) Then
                    If Not SP_FLDbManFolder(ptPthDri & "\" & tFolderDAT, nEN_Fle1Exist) Then
                        Call SP_FLDbManFolder(ptPthDri & "\" & tFolderDAT, nEN_Fld3Create)
                    End If
    
                    tFileName = tFolderDAT & "-" & tTmnNum & "-" & tTransNo & ".dat"             '\2014-07-21_14070006600017\2014-07-21_14070006600017-00066-02948.dat
                    tTemp = ptPthDat & "\" & tFolderDAT & "\" & tFileName
                    If SP_FLEbManFile(tTemp, nEN_Fle1Exist) Then
                        bW_CheckImp = True              '***********************************************
                        Call W_PROxCopyFile(tTemp, ptPthDri & "\" & tFolderDAT)                                                             'Copy To Dri
                        oblCopy.Caption = "Copying......" & tFileName
                        
                        Call W_PRCxUpdateStaOnOff(tTmnNum, tTransNo)
                        Call SP_CRTxLogFile("Copy DAT To Device:" & tFileName & " [Success]")
                    Else
                        Call SP_CRTxLogFile("Copy DAT To Device:" & tFileName & " [Fail]")
                    End If
                End If
                orsTemp.MoveNext
            Loop
            
            tPathEJ = oFSO.GetParentFolderName(App.Path) & "\POSEJ"
            tPathEJBK = tPathEJ & "\Backup"
            If SP_FLDbManFolder(tPathEJ, nEN_Fle1Exist) Then
                If SP_FLDbManFolder(tPathEJBK, nEN_Fle1Exist) Then
                    orsTemp.MoveFirst
                    Do While Not orsTemp.EOF
                        tDepID = SP_FEDtChkString(orsTemp, "FTShdDepID")
                        tStoreID = SP_FEDtChkString(orsTemp, "FTStmCode")
                        dTransDate = SP_FEDdChkDate(orsTemp, "FDShdTransDate")
                        tTransDateRE = SP_FEDtChkString(orsTemp, "FTCstCardCode3")
                        tTmnNum = SP_FEDtChkString(orsTemp, "FTTmnNum")
                        tTransNo = SP_FEDtChkString(orsTemp, "FTShdTransNo")
                        tDate = Format(dTransDate, "yyMMdd")
                        tFileName = Format(tDepID, "00") & Format(tTmnNum, "00000") & Format(tTransNo, "00000") & tStoreID & ".EJ"
                        tFileEj = tPathEJ & "\" & tDate & "\" & tFileName
                        
                        If SP_FLEbManFile(tFileEj, nEN_Fle1Exist) Then
                            If Not SP_FLDbManFolder(ptPthDri & "\" & tDate, nEN_Fle1Exist) Then
                                Call SP_FLDbManFolder(ptPthDri & "\" & tDate, nEN_Fld3Create)
                            End If
                            bW_CheckImp = True              '***********************************************
                            Call W_PROxCopyFile(tFileEj, ptPthDri & "\" & tDate)                                                             'Copy To Dri
                            oblCopy.Caption = "Copying......" & tFileName
                            Call SP_CRTxLogFile("Copy EJ To Device:" & tFileName & " [Success]")
                            If SP_FLEbManFile(ptPthDri & "\" & tDate & "\" & tFileName, nEN_Fle1Exist) Then         '*Ao 2014/08/08    V.1.132
                                If Not SP_FLDbManFolder(tPathEJBK & "\" & tDate, nEN_Fle1Exist) Then
                                    Call SP_FLDbManFolder(tPathEJBK & "\" & tDate, nEN_Fld3Create)
                                End If
                                Call W_PROxCopyFile(tFileEj, tPathEJBK & "\" & tDate)                                                             'Copy To Back
                                
                                Call SP_FLEbManFile(tFileEj, nEN_Fle2Delete)
                            End If
                            
                        Else
                            
                            tFileEj = tPathEJBK & "\" & tDate & "\" & tFileName
                            If SP_FLEbManFile(tFileEj, nEN_Fle1Exist) Then
                                If Not SP_FLDbManFolder(ptPthDri & "\" & tDate, nEN_Fle1Exist) Then
                                    Call SP_FLDbManFolder(ptPthDri & "\" & tDate, nEN_Fld3Create)
                                End If
                                bW_CheckImp = True              '***********************************************
                                Call W_PROxCopyFile(tFileEj, ptPthDri & "\" & tDate)                                                             'Copy To Dri
                                oblCopy.Caption = "Copying......" & tFileName
                                Call SP_CRTxLogFile("Copy EJ To Device:" & tFileName & " [Success]")
                            Else
                                Call SP_CRTxLogFile("Copy EJ To Device:" & tFileName & " [Fail]")
                            End If
                            
                        End If
                        orsTemp.MoveNext
                    Loop
                End If
            End If
'            Call SP_MSGnShowing("All Exported data will be log into files;All Exported data will be log into files", nCS_Information)
        Else
'            Call SP_MSGnShowing("Data had already been exported to Device;Data had already been exported to Device", nCS_Information)
        End If
    End If
    
    Set oFSO = Nothing
Exit Sub
ErrHandle:
    Set oFSO = Nothing
End Sub

Private Sub W_FNDxCopyEJToSV(ByVal ptPthDri As String, ByVal ptFileType As String)
'-------------------------------------------------------------------------------------------------
'Call:
'Cmt: '*Ao 2014/07/29    Copy EJ Dri To Server
'-------------------------------------------------------------------------------------------------
Dim oFSO As New FileSystemObject
Dim oFldRootIn As Folder
Dim oSubFolder As Folder
Dim oFileIn As File
Dim tYYMMDD As String, tPthSV As String, tPthEJ As String, tFileName As String, tTmp As String, tBackup As String
Dim tImport As String

On Error GoTo ErrHandle

'    tTmp = oFSO.GetParentFolderName(App.Path) & "\POSEJ\"                  '*Ao 2014/08/05      V 4.2.131
'    If Not SP_FLDbManFolder(tTmp, nEN_Fle1Exist) Then
'        Call SP_FLDbManFolder(tTmp, nEN_Fld3Create)
'    End If
    tImport = "Imported"
    If Not SP_FLDbManFolder(ptPthDri & "\" & tImport, nEN_Fle1Exist) Then               '*Ao 2014/08/05      V 4.2.131
        Call SP_FLDbManFolder(ptPthDri & "\" & tImport, nEN_Fld3Create)
    End If
    
    Set oFldRootIn = oFSO.GetFolder(ptPthDri)
    
    For Each oSubFolder In oFldRootIn.SubFolders
        tYYMMDD = oSubFolder.Name
        If Len(tYYMMDD) <= 6 And (UCase(tYYMMDD) <> UCase(tImport)) Then
'            tTmp = oFSO.GetParentFolderName(App.Path) & "\POSEJ\" & tYYMMDD                                         'Local                 '*Ao 2014/08/05      V 4.2.131
'            If Not SP_FLDbManFolder(tTmp, nEN_Fle1Exist) Then
'                Call SP_FLDbManFolder(tTmp, nEN_Fld3Create)
'            End If
'            tBackup = oFSO.GetParentFolderName(App.Path) & "\POSEJ\Backup\" & tYYMMDD                  'Local Backup
'            If Not SP_FLDbManFolder(tBackup, nEN_Fle1Exist) Then
'                Call SP_FLDbManFolder(tBackup, nEN_Fld3Create)
'            End If
             If Not SP_FLDbManFolder(ptPthDri & "\" & tImport & "\" & tYYMMDD, nEN_Fle1Exist) Then                                                '*Ao 2014/08/05      V 4.2.131
                Call SP_FLDbManFolder(ptPthDri & "\" & tImport & "\" & tYYMMDD, nEN_Fld3Create)
            End If
            tPthSV = tVB_CNOnDbSrcEJ & "\" & tYYMMDD
            If Not SP_FLDbManFolder(tPthSV, nEN_Fle1Exist) Then
                Call SP_FLDbManFolder(tPthSV, nEN_Fld3Create)
            End If
            For Each oFileIn In oSubFolder.Files
                oFileIn.Attributes = Normal                                                                                     'set file ไม่ให้ read only
                tPthEJ = oFileIn.Path
                tFileName = oFileIn.Name
                If UCase(Right(tPthEJ, Len(ptFileType))) = UCase(ptFileType) Then           '.EJ
'                    If SP_FLEbManFile(tBackup & "\" & tFileName, nEN_Fle1Exist) Then
'                        oFSO.DeleteFile tBackup & "\" & tFileName                                                    'ลบข้อมูลที่ Back up
'                    End If
                    bW_CheckImp = True                  '*********************
                    oblCopy.Caption = "Copying......" & tFileName
                    DoEvents
                    If SP_FLEbManFile(tPthSV & "\" & tFileName, nEN_Fle1Exist) Then
                        oFSO.DeleteFile tPthSV & "\" & tFileName                                                    'ลบข้อมูลที่ Server
                    End If
                    Call W_PROxCopyFile(tPthEJ, tPthSV)
                    If SP_FLEbManFile(tPthSV & "\" & tFileName, nEN_Fle1Exist) Then
                        Call W_PROxCopyFile(tPthEJ, ptPthDri & "\" & tImport & "\" & tYYMMDD)
                        Call SP_CRTxLogFile("Copy EJ From Device:" & tFileName & " [Success]")
                    Else
                        Call SP_CRTxLogFile("Copy EJ From Device:" & tFileName & " [Fail]")
                    End If
                    
                End If
            Next
            Call SP_FLDbManFolder(ptPthDri & "\" & tYYMMDD, nEN_Fld2Delete)
        End If
    Next
 
    Set oFSO = Nothing
    Exit Sub
ErrHandle:
    Set oFSO = Nothing
End Sub

Private Sub W_PROxCopyFile(ByVal ptPthEJ As String, ByVal ptBk As String)
Dim oFSO As New FileSystemObject

On Error GoTo ErrHandle

    oFSO.CopyFile ptPthEJ, ptBk & "\"
    Set oFSO = Nothing
Exit Sub
ErrHandle:
    Set oFSO = Nothing
End Sub

Private Sub W_FNDxCopyEJ(ByVal ptPthEJ As String, ByVal ptPthEJBk As String, ByVal ptBk As String, ByVal ptFileType As String)
'-------------------------------------------------------------------------------------------------
'Call:
'Cmt: '*Ao 2013/11/28    Copy EJ To Dri
'-------------------------------------------------------------------------------------------------
Dim oFSO As New FileSystemObject
Dim oFldRootIn As Folder
Dim oFileIn As File
Dim tPthEJ As String, tFileName As String
Dim tBKTemp As String

On Error GoTo ErrHandle

    Set oFldRootIn = oFSO.GetFolder(ptPthEJ)

    For Each oFileIn In oFldRootIn.Files
        tPthEJ = oFileIn.Path
        tFileName = oFileIn.Name
        If UCase(Right(tPthEJ, Len(ptFileType))) = UCase(ptFileType) Then           'POSEJ\YYMMDD
            If Not (SP_FLEbManFile(ptBk & "\" & tFileName, nEN_Fle1Exist)) Then
                Call W_PROxCopyFile(tPthEJ, ptBk)
                oblCopy.Caption = "Copying......" & tFileName
            End If
        End If
    Next oFileIn
    
    If SP_FLDbManFolder(ptPthEJBk, nEN_Fle1Exist) Then
        Set oFldRootIn = oFSO.GetFolder(ptPthEJBk)
        
        For Each oFileIn In oFldRootIn.Files
            tPthEJ = oFileIn.Path
            tFileName = oFileIn.Name
            If UCase(Right(tPthEJ, Len(ptFileType))) = UCase(ptFileType) Then           'POSEJ\Backup\YYMMDD
                If Not (SP_FLEbManFile(ptBk & "\" & tFileName, nEN_Fle1Exist)) Then
                    Call W_PROxCopyFile(tPthEJ, ptBk)
                    oblCopy.Caption = "Copying......" & tFileName
                End If
            End If
        Next oFileIn
    End If
    
    Set oFSO = Nothing
    Exit Sub
ErrHandle:

End Sub

Private Sub W_PRCxUpdateStaOnOff(ByVal ptTmnNum As String, ByVal ptTransNo As String)
'----------------------------------------------------------
'Call :
'Cmt : '*Eaw 56-11-14 หลังสร้างข้อมูล .DAT แล้วจะต้อง Update FTStaSendOnOff='1' ให้กับข้อมูลใน Local ด้วย
'CommSheet TK-ISS3000-425(Wait confirm)
'----------------------------------------------------------
    Dim tSql As String
On Error GoTo ErrHandle
    '*Update Table Sale
    tSql = "UPDATE TPSTSalHD"
    tSql = tSql & " SET FTStaSentOnOff = '1' "
    tSql = tSql & " WHERE FTTmnNum = '" & ptTmnNum & "'"
    tSql = tSql & " AND FTShdTransNo = '" & ptTransNo & "'"
    Call SP_SQLbExecute(tSql)
    tSql = "UPDATE TPSTSalDT "
    tSql = tSql & " SET FTStaSentOnOff = '1' "
    tSql = tSql & " WHERE FTTmnNum = '" & ptTmnNum & "'"
    tSql = tSql & " AND FTShdTransNo = '" & ptTransNo & "'"
    Call SP_SQLbExecute(tSql)
    tSql = "UPDATE TPSTSalCD "
    tSql = tSql & " SET FTStaSentOnOff = '1' "
    tSql = tSql & " WHERE FTTmnNum = '" & ptTmnNum & "'"
    tSql = tSql & " AND FTShdTransNo = '" & ptTransNo & "'"
    Call SP_SQLbExecute(tSql)
    tSql = "UPDATE TPSTSalRC "
    tSql = tSql & " SET FTStaSentOnOff = '1' "
    tSql = tSql & " WHERE FTTmnNum = '" & ptTmnNum & "'"
    tSql = tSql & " AND FTShdTransNo = '" & ptTransNo & "'"
    Call SP_SQLbExecute(tSql)
    
    '*Update Table Deposit
    tSql = "UPDATE TPSTDepositHD"
    tSql = tSql & " SET FTStaSentOnOff = '1' "
    tSql = tSql & " WHERE FTTmnNum = '" & ptTmnNum & "'"
    tSql = tSql & " AND FTShdTransNo = '" & ptTransNo & "'"
    Call SP_SQLbExecute(tSql)
    tSql = "UPDATE TPSTDepositDT "
    tSql = tSql & " SET FTStaSentOnOff = '1' "
    tSql = tSql & " WHERE FTTmnNum = '" & ptTmnNum & "'"
    tSql = tSql & " AND FTShdTransNo = '" & ptTransNo & "'"
    Call SP_SQLbExecute(tSql)
    tSql = "UPDATE TPSTDepositCD "
    tSql = tSql & " SET FTStaSentOnOff = '1' "
    tSql = tSql & " WHERE FTTmnNum = '" & ptTmnNum & "'"
    tSql = tSql & " AND FTShdTransNo = '" & ptTransNo & "'"
    Call SP_SQLbExecute(tSql)
    tSql = "UPDATE TPSTDepositRC "
    tSql = tSql & " SET FTStaSentOnOff = '1' "
    tSql = tSql & " WHERE FTTmnNum = '" & ptTmnNum & "'"
    tSql = tSql & " AND FTShdTransNo = '" & ptTransNo & "'"
    Call SP_SQLbExecute(tSql)
    
    '*Update Table CV
    tSql = "UPDATE TPSTCV"
    tSql = tSql & " SET FTStaSentOnOff = '1' "
    tSql = tSql & " WHERE FTPosNo = '" & ptTmnNum & "'"
    tSql = tSql & " AND FTTransNo = '" & ptTransNo & "'"
    Call SP_SQLbExecute(tSql)
    
    '*Update Table Sale Point
    tSql = "UPDATE TPSTSalePoint"
    tSql = tSql & " SET FTStaSentOnOff = '1' "
    tSql = tSql & " WHERE FTTmnNum = '" & ptTmnNum & "'"
    tSql = tSql & " AND FTShdTransNo = '" & ptTransNo & "'"
    Call SP_SQLbExecute(tSql)
    
    '*Update Table Voucher
    tSql = "UPDATE TPSTVoucherHD"
    tSql = tSql & " SET FTStaSentOnOff = '1' "
    tSql = tSql & " WHERE FTTmnNum = '" & ptTmnNum & "'"
    tSql = tSql & " AND FTShdTransNo = '" & ptTransNo & "'"
    Call SP_SQLbExecute(tSql)
    tSql = "UPDATE TPSTVoucherRC "
    tSql = tSql & " SET FTStaSentOnOff = '1' "
    tSql = tSql & " WHERE FTTmnNum = '" & ptTmnNum & "'"
    tSql = tSql & " AND FTShdTransNo = '" & ptTransNo & "'"
    Call SP_SQLbExecute(tSql)
    
    Exit Sub
ErrHandle:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If bVB_CNSplash Then KeyCode = 0 '*KT 51-09-17 ป้องการ Key กรณีที่ Splash ขึ้น
'    If KeyCode = 13 Then SendKeys "{TAB}" '*KT 53-01-19 ทำให้ขั้นตอนการ Setfocus ตอน Load ผิดพลาด ถ้ามีการใส่หน้าจอ Login มา
End Sub

Private Sub Form_Load()
    Call W_DisableUSB(3) '*TON 58-12-01
    Dim nRet As Long '*TON 58-12-15
    Dim tRunFile As String
     Dim sSave As String, Ret As Long
    sSave = Space(255)
    Ret = GetSystemDirectory(sSave, 255)
    sSave = Left$(sSave, Ret)
    tRunFile = sSave & "\devcon.exe enable  USBSTOR*"
    'MsgBox (tRunFile)
    'nRet = Shell(tRunFile, vbMinimizedFocus)
    nRet = Shell(tRunFile, vbHide)
    SleepEx 1000, 0 '**
    '=================
    Call SP_FRMxChildCenter(Me)
    Call SP_FRMxGetCaptionFrmDB2Tag(Me)
    Call SP_FRMxSetCapForm(Me)
    Select Case nVB_RTD
        Case 1
'            otbForm(0).Text = tVB_CNPthBnkTr & "\" & Replace(tUT_SaleDate, "/", "-") & "_" & tUT_LoginCode
            otbForm(0).Text = tVB_CNPthBnkTr                            '*Ao 2014/07/31     Copy Dat,EJ  ทุกอัน ที่ ISNULL(FTStaSentOnOff,'0') <> '1'
            otbForm(1).Text = tVB_CNPthDrive
        Case 2
            otbForm(0).Text = tVB_CNPthDrive
            otbForm(1).Text = tVB_CNPthBnkTr
        Case 3
            otbForm(0).Text = tVB_CNPthDrive
            otbForm(1).Text = tVB_CNPthBnkTr
    End Select
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call W_DisableUSB(4) '*TON 58-12-01
   Dim nRet As Long '*TON 58-12-15
    Dim tRunFile As String
    Dim sSave As String, Ret As Long
    sSave = Space(255)
    Ret = GetSystemDirectory(sSave, 255)
    sSave = Left$(sSave, Ret)
    tRunFile = sSave & "\devcon.exe disable  USBSTOR*"
    nRet = Shell(tRunFile, vbHide)
    SleepEx 1000, 0 '**
    Call SP_OBJxSetNothing(Me)
End Sub

Private Sub ocmBrw_Click(Index As Integer)
   Dim nPos As Integer
   Dim tFileName As String
   nPos = IIf(bVB_CNShwThai, 1, 2)
   tFileName = SP_FLDtBrowser(Me, cUT.UT_STRtToken(tCS_CNBchTF2, ";", nPos))
   If Trim(tFileName) <> "" Then otbForm(Index).Text = tFileName
   
End Sub

Private Sub ocmCancel_Click()
    Unload Me
End Sub

Public Function W_CHKxSQLVersion() As Integer
'-----------------------------------------------------------------------------------
'Call:W_CHKxSQLVersion()
'Cmt: *Bump 56-02-07 ตรวจสอบ Version ของ MS SQL Server
'-----------------------------------------------------------------------------------
    W_CHKxSQLVersion = 0

    Dim orsTemp As ADODB.Recordset
    Dim tSql As String
    Dim tSQLVersion As String
    Dim nVersion As String
    On Error GoTo ErrHandle
    
    W_CHKxSQLVersion = 9 'Default SQL2005
    
     tSql = "SELECT SERVERPROPERTY('productversion') as FTMSSQLVersion"

    Set orsTemp = ocnAll.Execute(tSql)

    If Not (orsTemp.BOF And orsTemp.EOF) Then
        tSQLVersion = SP_FEDtChkString(orsTemp, "FTMSSQLVersion")
        'SQL2005 9.00.XX eg. 9.00.1399.06
        'SQL2008 10.0.XX
        'SQL2012 11.0.XX
        tSQLVersion = Mid(tSQLVersion, 1, 4)
        nVersion = CInt(tSQLVersion) '9,10,11
        W_CHKxSQLVersion = nVersion
    Else
        W_CHKxSQLVersion = 9
    End If
ErrHandle:
    Set orsTemp = Nothing
End Function

Private Sub ocmOK_Click(Index As Integer)
    Dim tAlowUSB, tForm1, tForm0 As String
    'tAlowUSB = ReadRegistry(HKEY_CURRENT_USER, "SOFTWARE\VB and VBA Program Settings\AdaPosMall\POSFront\", "Auto-Disable-USB")
    tAlowUSB = GetSetting("AdaPosMall", "POSFront", "Auto-Disable-USB") '*TON 58-12-01
    tForm0 = SP_GETxDriveType(Left(otbForm(0).Text, 2))
    tForm1 = SP_GETxDriveType(Left(otbForm(1).Text, 2))

    Select Case nVB_RTD
        Case 1 'Copy Transaction To Device
            'If (tForm0 <> "Removable" And tForm0 <> "" And tForm1 <> "Removable" And tForm1 <> "") Or tAlowUSB = "0" Then '*TON 58-12-01
                Call SP_CRTxLogPrg(Format(Now, "yyyy-MM-dd hh:mm:ss") & "   Copy Transaction(" & tForm0 & ") To Device(" & tForm1 & ")")
                If MsgBox("Do you want save to device ? ", vbQuestion + vbYesNo, "Transaction") = vbYes Then
                    Call SP_GETbTmnData
                    If Not bVB_CNOpenTmnOff Then
                        If Not SP_FLDbManFolder(otbForm(0).Text, nEN_Fld1Exist) Then
                            Call SP_MSGnShowing("tms_CN005076", nCS_Warning)
    '                        otbForm(0).SetFocus
                            Call SP_OBJxSetFocus(otbForm(0))
                            Exit Sub
                        End If
                        If Not SP_FLDbManFolder(otbForm(1).Text, nEN_Fld1Exist) Then
                            Call SP_MSGnShowing("tms_CN005076", nCS_Warning)
    '                        otbForm(1).SetFocus
                            Call SP_OBJxSetFocus(otbForm(1))
                            Exit Sub
                        End If
                        bW_CheckImp = False             '***********************************************
                         Call W_STAxMoveFile(True)
                         Call W_MOVxSale2Drive
                         SleepEx 1500, 0
                         Call W_STAxMoveFile(False)
    '                     Call SP_MSGnShowing("tms_CN005081", nCS_Warning)
    '                     Call SP_MSGnShowing("All Exported data will be log into files;All Exported data will be log into files", nCS_Information)
                        If bW_CheckImp Then
                            Call SP_MSGnShowing("All Exported data will be log into files;All Exported data will be log into files", nCS_Information)
                        Else
                            Call SP_MSGnShowing("Data had already been exported to Device;Data had already been exported to Device", nCS_Information)
                        End If
                    Else
                        Call SP_MSGnShowing("tms_CN005077", nCS_Warning)
                    End If
                End If
'            Else
'                     Call SP_CRTxLogPrg(Format(Now, "yyyy-MM-dd hh:mm:ss") & "  USB port is disabled  - Copy Transaction(" & tForm0 & ") To Device(" & tForm1 & ")")
'                    MsgBox "Not Allow Copy Transaction To Device." '*TON 58-12-01
'            End If
        Case 2 'Copy Transaction From Device
            'If tForm0 <> "Removable" And tForm0 <> "" And tForm1 <> "Removable" And tForm1 <> "" Or tAlowUSB = "0" Then    '*TON 58-12-01
                Call SP_CRTxLogPrg(Format(Now, "yyyy-MM-dd hh:mm:ss") & "  Copy Transaction(" & tForm0 & ") From Device(" & tForm1 & ")")
                If MsgBox("Do you want copy from device ? ", vbQuestion + vbYesNo, "Transaction") = vbYes Then
                    
                    Call SP_GETbTmnData
                    If Not bVB_CNOpenTmnOff Then
                        If Not SP_FLDbManFolder(otbForm(0).Text, nEN_Fld1Exist) Then
                            Call SP_MSGnShowing("tms_CN005076", nCS_Warning)
    '                        otbForm(0).SetFocus
                            Call SP_OBJxSetFocus(otbForm(0))
                            Exit Sub
                        End If
                        If Not SP_FLDbManFolder(otbForm(1).Text, nEN_Fld1Exist) Then
                            Call SP_MSGnShowing("tms_CN005076", nCS_Warning)
    '                        otbForm(1).SetFocus
                            Call SP_OBJxSetFocus(otbForm(1))
                            Exit Sub
                        End If
                        
    '                    tVB_Transaction = aVB_TransMode(7).ID
    '                    tVB_TransactionSub = tVB_Transaction
    '                    tVB_TransName = aVB_TransMode(7).Name
    '                    tVB_TransNameABB = aVB_NameAbb(tEN_AbbCopyFromDev - 1).UsrLangPri & ";" & aVB_NameAbb(tEN_AbbCopyFromDev - 1).UsrLangSec
    '                    Call SP_INSxTransHD(tVB_CNTblHD)
                        bW_CheckImp = False             '***********************************************
                        Call W_STAxMoveFile(True)
                        
                        Call W_MOVxDrive2Sale
                        
    '                     nVB_TotalTmnNumCopy = IIf(LBound(aW_TmnNum) = 0, 0, LBound(aW_TmnNum) - 1)
    '                     nVB_TotalTmnNumCopy = tVB_CNTerminalNum
    '                     tVB_TrnNumCopyMin = tW_TrnMin
    '                     tVB_TrnNumCopyMax = tW_TrnMax
                        
    '                    Call wFunctionKB.W_SetToComplete
    '                    Call SP_UPDxSetCompleteTerminalMtn
    '                    Call UT_ClearGenNew
                        
                        SleepEx 1500, 0
                        Call W_STAxMoveFile(False)
    '                    Call SP_MSGnShowing("tms_CN005081", nCS_Warning)
                        If bW_CheckImp Then
                            Call SP_MSGnShowing("All imported data will be log into files;All imported data will be log into files", nCS_Information)
                        Else
                            Call SP_MSGnShowing("No data found in device;No data found in device", nCS_Information)
                        End If
                        Call SaveSetting("AdaPosMall", "POSFront", "SentSale", "1")     '*Ao 2013/11/22 ให้ส่งข้อมูล
                    Else
                        Call SP_MSGnShowing("tms_CN005077", nCS_Warning)
                    End If
                End If
'            Else
'                    Call SP_CRTxLogPrg(Format(Now, "yyyy-MM-dd hh:mm:ss") & "  USB port is disabled  - Copy Transaction(" & tForm0 & ") From Device(" & tForm1 & ")")
'                    MsgBox "Not Allow Copy Transaction From Device."
'            End If '*TON 58-12-01
        Case 3 'Update Master
            Call SP_CRTxLogPrg(Format(Now, "yyyy-MM-dd hh:mm:ss") & "  Update Master(" & tForm0 & ") From (" & tForm1 & ")")
            If MsgBox("Do you want update Master ? ", vbQuestion + vbYesNo, "Transaction") = vbYes Then
                If Not SP_FLDbManFolder(otbForm(0).Text, nEN_Fld1Exist) Then
                    Call SP_MSGnShowing("tms_CN005076", nCS_Warning)
'                    otbForm(0).SetFocus
                    Call SP_OBJxSetFocus(otbForm(0))
                    Exit Sub
                End If
                If Not SP_FLDbManFolder(otbForm(1).Text, nEN_Fld1Exist) Then
                    Call SP_MSGnShowing("tms_CN0    05076", nCS_Warning)
'                    otbForm(1).SetFocus
                    Call SP_OBJxSetFocus(otbForm(1))
                    Exit Sub
                End If
                
'               tVB_Transaction = aVB_TransMode(7).ID
'               tVB_TransactionSub = tVB_Transaction
'               tVB_TransName = aVB_TransMode(7).Name
'               tVB_TransNameABB = aVB_NameAbb(tEN_AbbCopyMaster - 1).UsrLangPri & ";" & aVB_NameAbb(tEN_AbbCopyMaster - 1).UsrLangSec
'               Call SP_INSxTransHD(tVB_CNTblHD)


                Call W_STAxMoveFile(True)
                Call W_MOVxDrive2Master
                
'               Call wFunctionKB.W_SetToComplete
'               Call SP_UPDxSetCompleteTerminalMtn
'               Call UT_ClearGenNew

                Call W_STAxMoveFile(False)
                Call SP_MSGnShowing("tms_CN005081", nCS_Warning)
            End If
    End Select
    Unload Me
End Sub

Private Sub W_MOVxSale2Drive()
'-------------------------------------------------------------------------------------------------
'  Call:
'  Cmt: ทำการ Move file master from transaction to trumpdrive
'-------------------------------------------------------------------------------------------------
Dim tPthDrive$
Dim tPthSale$
Dim tPathEJ As String, tPathEJBK As String
Dim tBKTemp  As String

    Dim oFSO As New FileSystemObject
    tPthDrive = otbForm(1).Text
    tPthSale = otbForm(0).Text
'    Call W_FNDxNumAndDate(tPthSale, tPthDrive, ".DAT")                 '*Ao 2014/07/31  CommSheet TK-ISS3000-460-00 00 03 V 4.1.129
    'clear all in "Import"
    'Call SP_CLRxClearImp(tPthSale, ".DAT")
    '---------------------------------------------------------------------------------------------------
    '*Ao 2014/02/03    Copy EJ
'    tPathEJ = oFSO.GetParentFolderName(App.Path) & "\POSEJ"
'    tPathEJBK = tPathEJ & "\Backup\" & Format(CDate(tUT_SaleDate), "yyMMdd")                    '*Ao 2014/07/31  CommSheet TK-ISS3000-460-00 00 03 V 4.1.129
'    If SP_FLDbManFolder(tPathEJ, nEN_Fle1Exist) Then
'        tPathEJ = tPathEJ & "\" & Format(CDate(tUT_SaleDate), "yyMMdd")
'        If SP_FLDbManFolder(tPathEJ, nEN_Fle1Exist) Then
'            If SP_FLDbManFolder(tPthDrive & "\" & Format(CDate(tUT_SaleDate), "yyMMdd"), nEN_Fld3Create) Then
'                tBKTemp = tPthDrive & "\" & Format(CDate(tUT_SaleDate), "yyMMdd")
'            Else
'                MsgBox "Not Create Folder " & tPthDrive & "\" & Format(CDate(tUT_SaleDate), "yyMMdd")
'                Exit Sub
'            End If
'            Call W_FNDxCopyEJ(tPathEJ, tPathEJBK, tBKTemp, ".EJ")
'        Else
'
'        End If
'    Else
'
'    End If
    Call W_FNDxCopyDatToDri(tPthSale, tPthDrive, ".DAT")
    
'    If Not SP_FLDbManFolder(tPathEJ, nEN_Fle1Exist) Then
'        tPathEJ = oFSO.GetParentFolderName(App.Path)
'        tPathEJBK = tPathEJ & "\POSEJ\Backup\" & Format(CDate(tUT_SaleDate), "yyMMdd")
'        tPathEJ = tPathEJ & "\POSEJ\" & Format(CDate(tUT_SaleDate), "yyMMdd")
'        If SP_FLDbManFolder(tPathEJ, nEN_Fle1Exist) Then
'            Call W_FNDxCopyEJ(tPathEJ, tPathEJBK, tPthDrive, ".EJ")
'        Else
'            MsgBox "tPathEJ = " & tPathEJ
'        End If
'    Else
'        tPathEJBK = tPathEJ & "\Backup\" & Format(CDate(tUT_SaleDate), "yyMMdd")
'        tPathEJ = tPathEJ & "\" & Format(CDate(tUT_SaleDate), "yyMMdd")
'        If SP_FLDbManFolder(tPathEJ, nEN_Fle1Exist) Then
''            MsgBox "tPathEJ = " & tPathEJ
'            Call W_FNDxCopyEJ(tPathEJ, tPathEJBK, tPthDrive, ".EJ")
'        Else
'            MsgBox "tPathEJ = " & tPathEJ
'        End If
'    End If
    '---------------------------------------------------------------------------------------------------
End Sub

Private Sub W_MOVxDrive2Sale()
'-------------------------------------------------------------------------------------------------
'  Call:
'  Cmt: ทำการ Move file master from trumpdrive to transaction
'-------------------------------------------------------------------------------------------------
    Dim tPthDrive$
    Dim tPthSale$
    Dim oFSO As New FileSystemObject
    tPthDrive = otbForm(0).Text
    tPthSale = otbForm(1).Text
'    Call W_FNDxNumAndDate(tPthDrive, tPthSale, ".DAT")              '*Ao 2014/03/25              ไม่ UPDATE FTStasendOnOff
    
'    Call W_FNDxCopyDatToSale(tPthDrive, tPthSale, ".DAT")           '*Ao 2014/07/31   CommSheet TK-ISS3000-460-00 00 03   V4.1.129
'    Call W_PRCxImpTransSale
    bW_CheckImp = False
    Call W_PRCxImportDAT(tPthDrive, tPthSale, ".DAT")            '*Ao 2014/07/31   CommSheet TK-ISS3000-460-00 00 03   V4.1.131
    
    Call SaveSetting("AdaPosMall", "POSFront", "SentSale", "1")     '*Ao 2013/11/22 ให้ส่งข้อมูล
    
    Call W_FNDxCopyEJToSV(tPthDrive, ".EJ")                     '*Ao 2014/07/29 TK-ISS3000-460 Copy Transaction from Device
    
    Call SaveSetting("AdaPosMall", "POSFront", "SentSale", "1")     '*Ao 2013/11/22 ให้ส่งข้อมูล
        
End Sub

Private Sub W_MOVxDrive2Master()
'-------------------------------------------------------------------------------------------------
'  Call:
'  Cmt: ทำการ Move file master from trumpdrive to master
'-------------------------------------------------------------------------------------------------
    Dim tPthDrive$
    Dim tPthSale$
    Dim oFSO As New FileSystemObject
    Dim tSql$
    Dim orsMst As ADODB.Recordset
    Dim tTbl As String
    Dim tFileImp As String
    Dim oBcp As New cABulkCopy
    Dim tApp As String
    Dim oFile As File
On Error GoTo ErrHandle
    tPthDrive = otbForm(0).Text
    tPthSale = otbForm(1).Text
    tSql = "SELECT Name  FROM Sys.Sysobjects WHERE xtype = 'U' ORDER BY Name"
    If SP_SQLvExecute(tSql, orsMst) = 0 Then
        If Not (orsMst.BOF And orsMst.EOF) Then
            Do While Not orsMst.EOF
                tTbl = SP_FEDtChkString(orsMst, "Name")
                tFileImp = tPthSale & "\" & tTbl & ".TXT"
                If SP_FLEbManFile(tFileImp, nEN_Fle1Exist) Then
                    Set oFile = oFSO.GetFile(tFileImp)
                    oFile.Attributes = Normal
                    oFSO.DeleteFile (tFileImp)
                End If
                orsMst.MoveNext
            Loop
            Call W_FNDxNumAndDate(tPthDrive, tPthSale, ".TXT")
            'clear all in "Import"
'            Call SP_CLRxClearImp(tPthDrive,".TXT")
            If Right$(tPthSale, 1) = "\" Then tPthSale = UCase$(Mid$(tPthSale, 1, Len(tPthSale) - 1))
            If Right$(App.Path, 1) = "\" Then
                tApp = UCase$(Mid$(App.Path, 1, Len(App.Path) - 1))
            Else
                tApp = UCase$(App.Path)
            End If
            If tPthSale = tApp Then   'Path export ต้องไม่เท่ากับ Application path
                Call SP_MSGnShowing("tms_CN014028", tCS_CNMsgErr)
                Exit Sub
            End If
            Call oBcp.W_DEFxConnect(tVB_CNOffDbSrc, tVB_CNOffDbSQLAuthUsrC, tVB_CNOffDbSQLAuthUsrP)
            If Not (orsMst.BOF And orsMst.EOF) Then orsMst.MoveFirst
            Do While Not orsMst.EOF
                tTbl = SP_FEDtChkString(orsMst, "Name")
                tFileImp = tPthSale & "\" & tTbl & ".txt"
                If (SP_CHKbTable(tTbl)) And SP_FLEbManFile(tFileImp, nEN_Fle1Exist) Then   'พบไฟล์ และ มีตารางในฐานข้อมูล
                    Me.oblCopy.Caption = "Copying......" & tTbl
                    DoEvents
                    Call oBcp.W_GETnImportTbl(tVB_CNOffDbCurFile, tTbl, tFileImp)
                    If SP_GETbFindFile(tFileImp) <> 0 Then  'Import File สำเร็จ
                        Call Kill(tFileImp)  'ลบไฟล์ Import
                    End If
                End If
                orsMst.MoveNext
            Loop
        End If
    End If
    Call SP_DATxRsNothing(orsMst)
    Exit Sub
ErrHandle:
    Call SP_DATxRsNothing(orsMst)
End Sub

Private Sub W_FNDxNumAndDate(ByVal ptPthIn As String, ByVal ptPthBk As String, ByVal ptFileType$)
'-------------------------------------------------------------------------------------------------
'Call:
'Cmt: ทำการ Move File ตามนามสกุลที่ส่งไปยัง Path ที่ระบุไว้
'-------------------------------------------------------------------------------------------------
    Dim oFSO As New FileSystemObject
    Dim tPthIn$, tPthBk$, tPthLastIn$, tPthDateIn$
    Dim oFldRootIn As Folder
    Dim oFileIn As File
    Dim nInFiles%, nBkFiles%, nLoop%
    Dim tFileName$
    Dim tTmnNum$
    Dim tTrnNo$
    Dim nI As Long
    Dim bPass As Boolean

    Set oFldRootIn = oFSO.GetFolder(ptPthIn)
    nInFiles = 0
    For Each oFileIn In oFldRootIn.Files
        tPthIn = oFileIn.Path
        tPthDateIn = oFileIn.DateCreated
        tFileName = oFileIn.Name
        If UCase(Right(tPthIn, 4)) = UCase(ptFileType) Then
            nInFiles = nInFiles + 1
            If nInFiles >= ogdIn.rows Then ogdIn.rows = ogdIn.rows + 1
            ogdIn.TextMatrix(nInFiles, 0) = tPthIn
            ogdIn.TextMatrix(nInFiles, 1) = tPthDateIn
            ogdIn.TextMatrix(nInFiles, 2) = tFileName
        End If
    Next oFileIn
    
    ReDim aW_TmnNum(0)
    tW_TrnMin = ""
    tW_TrnMax = ""
    For nLoop = 1 To nInFiles
            ogdIn.ColSort(1) = flexSortGenericDescending
            tPthLastIn = ogdIn.TextMatrix(nLoop, 0)
            tFileName = ogdIn.TextMatrix(nLoop, 2)
            oFSO.CopyFile tPthLastIn, ptPthBk & "\"
            nBkFiles = nBkFiles + 1
            oblCopy.Caption = "Copying......" & tFileName
            DoEvents
'            tTmnNum = Mid(tFileName, 26, 5)
'            tTrnNo = Mid(tFileName, 32, 5)
            
            tTmnNum = Mid(tFileName, 27, 5)
            tTrnNo = Mid(tFileName, 33, 5)
'            Call W_PRCxUpdateStaOnOff(tTmnNum, tTrnNo) '*Eaw 56-11-14 เรียกให้ Update FTStasendOnOff CommSheet TK-ISS3000-425(Wait confirm)
'            Call W_PRCxUpdateStaOnOff(tTmnNum, tTrnNo) '*Bump 57-03-04 เรียกให้ Update FTStasendOnOff CommSheet TK-ISS3000-425( FSBT Confirm 04/03/2014)
            Select Case nVB_RTD
                Case 1                                          'Copy Transaction To Device
                    Call W_PRCxUpdateStaOnOff(tTmnNum, tTrnNo)
                Case Else   'N/A    'Ao 2014/03/25
                    
            End Select
            If nLoop = 1 Then tW_TrnMin = tTrnNo
            bPass = False
            If UBound(aW_TmnNum) < LBound(aW_TmnNum) Then
                For nI = UBound(aW_TmnNum) To LBound(aW_TmnNum) - 1
                    If aW_TmnNum(nI) = tTmnNum Then bPass = True: Exit For
                Next nI
            End If
            If Not bPass Then
                aW_TmnNum(UBound(aW_TmnNum)) = tTmnNum
                ReDim Preserve aW_TmnNum(UBound(aW_TmnNum) + 1)
            End If
            
    Next nLoop
    tW_TrnMax = tTrnNo
    Set oFSO = Nothing
End Sub

Private Sub otbForm_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    '*KT 52-06-05 ในหน้าจอ Tender ถ้ากดปุ่ม Price Check จะปรากฏ Screen Find Customer แล้วถ้ารูดบัตร โปรแกรมจะ Hang PH 2.0 ML-SF-012
    '*KT 52-06-05 PH 2.0 ML-SF-013 ในหน้าจอ Discount --> Mall Card --> Price check --> Enter เครื่อง POS จะHang โดยต้อง Restart อย่างเดียว
    If Shift = 0 Then
        Select Case KeyCode
            Case vbKeyF4
                ocmBrw_Click (Index)
        End Select
    End If
End Sub

Private Sub W_PRCxImpTransSale()
'--------------------------------------------------------------
'Call:
'Cmt:
'-------------------------------------------------------------
    Dim nI As Long
    Dim tSql$
    Set oTransPrn.ActiveConnection = ocnAll
    nVB_SQLVersion = W_CHKxSQLVersion '*Bump 56-03-06 เพิ่มการเช็ค Version ของ MS SQL  CommSheet TK-ISS3000-036.xls
    For nI = 1 To ogdIn.rows - 1
            If ogdIn.TextMatrix(nI, 0) <> "" Then
'            oblCopy.Caption = "Import file:" & ogdIn.TextMatrix(nI, 2)                 '*Ao 2014/07/31   CommSheet TK-ISS3000-460-00 00 03   V4.1.129
            oblCopy.Caption = "Import file:" & ogdIn.TextMatrix(nI, 1)
            DoEvents
            If nI = 1 Then
                 '*KT 51-12-15 กำหนดตำแหน่ง Path ผิด
                 'Call oTransPrn.Import(otbForm(1).Text & "\" & ogdIn.TextMatrix(nI, 2), oDB.OnServerName, oDB.OnUsrName, oDB.OnUsrPwd, oDB.OnDbName)
                 Call oTransPrn.Import(otbForm(1).Text & "\" & ogdIn.TextMatrix(nI, 2), oDB.OffServerName, oDB.OffUsrName, oDB.OffUsrPwd, oDB.OffDbName) '*Bump 57-03-06 460 TK-ISS3000-460  Copy Transaction from Device
            Else
                 Call oTransPrn.Import(otbForm(1).Text & "\" & ogdIn.TextMatrix(nI, 2)) '*KT 51-12-15 กำหนดตำแหน่ง Path ผิด เพราะ ogdIn.TextMatrix(nI, 0) เก็บ Path ของ TrumDrive
            End If
        End If
    Next nI
    oblCopy.Caption = "Copying....."
    DoEvents
    '*KT 51-12-17 V 1.0.3 Phase 2 ทำการ Update FTShdStaCancel กรรีที่มีการทำ IPV ,Cancel
    '-------------------------------------------------------------------
    '    '*DEPOSIT
'        tSql = "UPDATE  TPSTSalHD"
'        tSql = tSql & vbCrLf & " SET FTShdStaCancel ='Y' "
'        tSql = tSql & vbCrLf & " WHERE  FTTmnNum IN (SELECT FTShdTmnDepositCancel  FROM  TPSTSalHD"
'        tSql = tSql & vbCrLf & "                                            WHERE FTShdStaDoc = '1'"
'        tSql = tSql & vbCrLf & "                                            AND FTShdTransType  IN  ('" & Format(tEN_TrnDepositCancel, "00") & "','" & Format(tEN_TrnAddDepositCancel, "00") & "')"
'        tSql = tSql & vbCrLf & "                                           )"
'        tSql = tSql & vbCrLf & " AND FTShdTransNo IN (SELECT  FTShdTrnDepositCancel FROM  TPSTSalHD"
'        tSql = tSql & vbCrLf & "                                            WHERE FTShdStaDoc = '1'"
'        tSql = tSql & vbCrLf & "                                            AND FTShdTransType  IN  ('" & Format(tEN_TrnDepositCancel, "00") & "','" & Format(tEN_TrnAddDepositCancel, "00") & "')"
'        tSql = tSql & vbCrLf & "                                           )"
'        tSql = tSql & vbCrLf & "AND FDShdTransDate = (SELECT  FDShdDepReTransDate FROM  TPSTSalHD"
'        tSql = tSql & vbCrLf & "                                            WHERE FTShdStaDoc = '1'"
'        tSql = tSql & vbCrLf & "                                            AND FTShdTransType  IN  ('" & Format(tEN_TrnDepositCancel, "00") & "','" & Format(tEN_TrnAddDepositCancel, "00") & "')"
'        tSql = tSql & vbCrLf & "                                           )"
'        tSql = tSql & vbCrLf & "AND FTShdStaDoc = '1'"
        tSql = "UPDATE  TPSTSalHD"
        tSql = tSql & vbCrLf & " SET FTShdStaCancel ='Y' "
        tSql = tSql & vbCrLf & " WHERE  FTTmnNum IN (SELECT FTShdDepRefTmnNum  FROM  TPSTSalHD"
        tSql = tSql & vbCrLf & "                                            WHERE FTShdStaDoc = '1'"
        tSql = tSql & vbCrLf & "                                            AND FTShdTransType  IN  ('" & Format(tEN_TrnDepositCancel, "00") & "','" & Format(tEN_TrnAddDepositCancel, "00") & "')"
        tSql = tSql & vbCrLf & "                                           )"
        tSql = tSql & vbCrLf & " AND FTShdTransNo IN (SELECT  FTShdDepReTransNo FROM  TPSTSalHD"
        tSql = tSql & vbCrLf & "                                            WHERE FTShdStaDoc = '1'"
        tSql = tSql & vbCrLf & "                                            AND FTShdTransType  IN  ('" & Format(tEN_TrnDepositCancel, "00") & "','" & Format(tEN_TrnAddDepositCancel, "00") & "')"
        tSql = tSql & vbCrLf & "                                           )"
        tSql = tSql & vbCrLf & "AND FDShdTransDate IN (SELECT  FDShdDepReTransDate FROM  TPSTSalHD"
        tSql = tSql & vbCrLf & "                                            WHERE FTShdStaDoc = '1'"
        tSql = tSql & vbCrLf & "                                            AND FTShdTransType  IN  ('" & Format(tEN_TrnDepositCancel, "00") & "','" & Format(tEN_TrnAddDepositCancel, "00") & "')"
        tSql = tSql & vbCrLf & "                                           )"
        tSql = tSql & vbCrLf & "AND FTShdStaDoc = '1'"
        Call SP_SQLvExecute(tSql, , oDB.ocnOnLine) 'PH 2.0.0 connect database server
        Call SP_SQLvExecute(tSql) 'Call ocnAll.Execute(tSql)
    
        '*tEN_TrnIPV/tEN_TrnIPVReturn/tEN_TrnIPVRedeem/tEN_TrnIPVVIPService/tEN_TrnIPVTaxExceptSale/tEN_TrnIPVDutyFreeSale/tEN_TrnIPVOtherReceive/tEN_TrnIPVOtherDisbursement
        tSql = "UPDATE  TPSTSalHD"
        tSql = tSql & vbCrLf & " SET FTShdStaCancel ='Y' "
        tSql = tSql & vbCrLf & " WHERE  FTTmnNum IN (SELECT FTShdDepRefTmnNum FROM   TPSTSalHD"
        tSql = tSql & vbCrLf & "                                            WHERE FTShdStaDoc = '1'"
        tSql = tSql & vbCrLf & "                                            AND FTShdTransType  IN  ('" & Format(tEN_TrnIPV, "00") & "','" & Format(tEN_TrnIPVReturn, "00") & "'"
        tSql = tSql & vbCrLf & "                                                                                                    ,'" & Format(tEN_TrnIPVRedeem, "00") & "','" & Format(tEN_TrnIPVVIPService, "00") & "'"
        tSql = tSql & vbCrLf & "                                                                                                    ,'" & Format(tEN_TrnIPVTaxExceptSale, "00") & "','" & Format(tEN_TrnIPVDutyFreeSale, "00") & "'"
        tSql = tSql & vbCrLf & "                                                                                                    ,'" & Format(tEN_TrnIPVOtherReceive, "00") & "','" & Format(tEN_TrnIPVOtherDisbursement, "00") & "')"
        tSql = tSql & vbCrLf & "                                           )"
        tSql = tSql & vbCrLf & " AND FTShdTransNo IN (SELECT FTShdDepReTransNo FROM   TPSTSalHD"
        tSql = tSql & vbCrLf & "                                            WHERE FTShdStaDoc = '1'"
        tSql = tSql & vbCrLf & "                                            AND FTShdTransType  IN  ('" & Format(tEN_TrnIPV, "00") & "','" & Format(tEN_TrnIPVReturn, "00") & "'"
        tSql = tSql & vbCrLf & "                                                                                                    ,'" & Format(tEN_TrnIPVRedeem, "00") & "','" & Format(tEN_TrnIPVVIPService, "00") & "'"
        tSql = tSql & vbCrLf & "                                                                                                    ,'" & Format(tEN_TrnIPVTaxExceptSale, "00") & "','" & Format(tEN_TrnIPVDutyFreeSale, "00") & "'"
        tSql = tSql & vbCrLf & "                                                                                                    ,'" & Format(tEN_TrnIPVOtherReceive, "00") & "','" & Format(tEN_TrnIPVOtherDisbursement, "00") & "')"
        tSql = tSql & vbCrLf & "                                           )"
        tSql = tSql & vbCrLf & "AND FDShdTransDate = (SELECT FDShdDepReTransDate FROM   TPSTSalHD"
        tSql = tSql & vbCrLf & "                                            WHERE FTShdStaDoc = '1'"
        tSql = tSql & vbCrLf & "                                            AND FTShdTransType  IN  ('" & Format(tEN_TrnIPV, "00") & "','" & Format(tEN_TrnIPVReturn, "00") & "'"
        tSql = tSql & vbCrLf & "                                                                                                    ,'" & Format(tEN_TrnIPVRedeem, "00") & "','" & Format(tEN_TrnIPVVIPService, "00") & "'"
        tSql = tSql & vbCrLf & "                                                                                                    ,'" & Format(tEN_TrnIPVTaxExceptSale, "00") & "','" & Format(tEN_TrnIPVDutyFreeSale, "00") & "'"
        tSql = tSql & vbCrLf & "                                                                                                    ,'" & Format(tEN_TrnIPVOtherReceive, "00") & "','" & Format(tEN_TrnIPVOtherDisbursement, "00") & "')"
        tSql = tSql & vbCrLf & "                                           )"
        tSql = tSql & vbCrLf & "AND FTShdStaDoc = '1'"
        Call SP_SQLvExecute(tSql, , oDB.ocnOnLine) 'PH 2.0.0 connect database server
        Call SP_SQLvExecute(tSql)  'Call ocnAll.Execute(tSql)
                
        tSql = "UPDATE  TPSTSalHD"
        tSql = tSql & vbCrLf & " SET FTShdStaCancel ='Y' "
        tSql = tSql & vbCrLf & " WHERE  FTTmnNum IN (SELECT FTShdDepRefTmnNum FROM  TPSTSalHD"
        tSql = tSql & vbCrLf & "                                            WHERE FTShdStaDoc = '1'"
        tSql = tSql & vbCrLf & "                                            AND FTShdTransType  IN  ('" & Format(tEN_TrnCancelVoucher, "00") & "')"
        tSql = tSql & vbCrLf & "                                           )"
        tSql = tSql & vbCrLf & " AND FTShdTransNo IN (SELECT FTShdDepReTransNo FROM  TPSTSalHD"
        tSql = tSql & vbCrLf & "                                            WHERE FTShdStaDoc = '1'"
        tSql = tSql & vbCrLf & "                                            AND FTShdTransType  IN  ('" & Format(tEN_TrnCancelVoucher, "00") & "')"
        tSql = tSql & vbCrLf & "                                           )"
        tSql = tSql & vbCrLf & "AND FDShdTransDate = (SELECT FDShdDepReTransDate FROM  TPSTSalHD"
        tSql = tSql & vbCrLf & "                                            WHERE FTShdStaDoc = '1'"
        tSql = tSql & vbCrLf & "                                            AND FTShdTransType  IN  ('" & Format(tEN_TrnCancelVoucher, "00") & "')"
        tSql = tSql & vbCrLf & "                                           )"
        tSql = tSql & vbCrLf & "AND FTShdStaDoc = '1'"
        Call SP_SQLvExecute(tSql, , oDB.ocnOnLine) 'PH 2.0.0 connect database server
        Call SP_SQLvExecute(tSql)  'Call ocnAll.Execute(tSql)
        '--------------------------------------------------------------------
        Call SaveSetting("AdaPosMall", "POSFront", "SentSale", "1")     '*Ao 2013/11/22 ให้ส่งข้อมูล
End Sub

Private Sub W_STAxMoveFile(pbEnable As Boolean)
'----------------------------------------------------------
'   Call:
'   Cmt:    show from splash
'----------------------------------------------------------
On Error Resume Next
    If pbEnable Then
        Animation1.Visible = True
        Call Animation1.Open(App.Path & "\FILECOPY_16.AVI")
        Call Animation1.Play
        oblCopy.Visible = True
'        DoEvents        '*KT 53-01-14
    Else
        Call Animation1.Stop
        Call Animation1.Close
        Animation1.Visible = False
        oblCopy.Visible = False
'        DoEvents        '*KT 53-01-14
    End If
End Sub

