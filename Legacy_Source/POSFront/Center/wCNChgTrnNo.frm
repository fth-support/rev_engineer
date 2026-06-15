VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.1#0"; "Codejock.SuiteCtrls.v11.1.3.ocx"
Begin VB.Form wCNChgTrnNo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Change Transaction Number;klg_Title"
   ClientHeight    =   4425
   ClientLeft      =   5925
   ClientTop       =   2340
   ClientWidth     =   5880
   Icon            =   "wCNChgTrnNo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   5880
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   1200
      Index           =   0
      Left            =   0
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   3480
      Width           =   5865
      _Version        =   720897
      _ExtentX        =   10345
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
         Left            =   2640
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
         Left            =   4200
         TabIndex        =   2
         TabStop         =   0   'False
         Tag             =   "2;"
         Top             =   240
         Width           =   1455
      End
   End
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   3750
      Index           =   2
      Left            =   0
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   10
      Width           =   5865
      _Version        =   720897
      _ExtentX        =   10345
      _ExtentY        =   6615
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
      PaintManager.Position=   2
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
         Left            =   2640
         MaxLength       =   5
         TabIndex        =   0
         Top             =   960
         Width           =   2535
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Current Tran No. :;klg_CurTrn"
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
         Left            =   240
         TabIndex        =   7
         Tag             =   "2;"
         Top             =   360
         Width           =   2265
      End
      Begin VB.Label olaTrnNo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "XXXXXXXXX"
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
         Left            =   2640
         TabIndex        =   6
         Tag             =   "2;"
         Top             =   360
         Width           =   1485
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
         Left            =   240
         TabIndex        =   5
         Tag             =   "2;"
         Top             =   960
         Width           =   2220
      End
   End
End
Attribute VB_Name = "wCNChgTrnNo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Call SP_FRMxChildCenter(Me)
    Call SP_FRMxGetCaptionFrmDB2Tag(Me)
    Call SP_FRMxSetCapForm(Me)
    olaTrnNo.Caption = W_GETtCurTrnNoBySale
    otbTransNo.Text = olaTrnNo.Caption
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrHandle
    Call SP_OBJxSetNothing(Me)
ErrHandle:
    
End Sub

Private Sub ocmCancel_Click()
    Unload Me
End Sub

Private Sub ocmOK_Click()
    Dim tTrnNo$
    Dim tSql$
    Dim nEff As Long
    
    tTrnNo = Trim(otbTransNo.Text)
    tTrnNo = Format(Val(Format(tTrnNo)), "00000")
    otbTransNo.Text = tTrnNo
'    DoEvents        '*KT 53-01-14
    If SP_MSGnShowing("tms_CN005196", nCS_OkCancel) = vbOK Then
        If Val(Format(tTrnNo)) >= Val(Format(olaTrnNo.Caption)) Then
            'Update ที่ ini Config
            tSql = "UPDATE"
            tSql = tSql & "  TConfigTerminal"
            tSql = tSql & " SET FTCfgValue = '" & tTrnNo & "'"
            tSql = tSql & " WHERE FTCfgCode='LRunning'"
            nEff = 0
            Call ocnCfg.Execute(tSql, nEff)
            If nEff > 0 Then
                'Update Terminal Online/OfflineCall ocmCancel_Click
                tSql = "UPDATE "
                tSql = tSql & vbCrLf & "TCNMTerminalMtn"
                tSql = tSql & vbCrLf & "SET FTTmnLSaleNo = '" & tTrnNo & "'"
                tSql = tSql & vbCrLf & "WHERE   FTTmnNum = '" & tVB_CNTerminalNum & "'"
                nEff = 0
                If SP_SQLvExecute(tSql, , , , , nEff) = 0 Then
                    If nEff > 0 Then
                        Call SP_SQLvExecute(tSql, , oDB.ocnOnLine) 'PH 2.0.0 connect database server
                        Call ocmCancel_Click
                    Else
                        MsgBox "Update not success.", vbCritical
                    End If
                Else
                    MsgBox "Update not success.", vbCritical
                End If
            Else
                MsgBox "Update not success.", vbCritical
            End If
        Else
            Call SP_MSGnShowing("tms_CN005197", nCS_Warning)
        End If
    Else
        Call ocmCancel_Click
    End If
End Sub

Private Function W_GETtCurTrnNoBySale() As String
    Dim orsTemp As ADODB.Recordset
    Dim tSql$
    Dim tTmp$
    Dim tTmp2$
On Error GoTo ErrHandle
    'Table Temp
    If SP_TBLbChkExist("TPSHD" & tVB_CNTerminalNum, oDB.ocnOnLine) Then
        tSql = "SELECT  MAX(FTShdTransNo)  AS FTTrnNo FROM  TPSHD" & tVB_CNTerminalNum
        tSql = tSql & vbCrLf & "WHERE FTShdStaDoc = '1'"
        If SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine) = 0 Then
            If Not (orsTemp.BOF Or orsTemp.EOF) Then
                tTmp = Format(Val(Format(SP_FEDtChkString(orsTemp, "FTTrnNo"))), "00000")
            Else
                tTmp = "00000"
            End If
        Else
             tTmp = "00000"
        End If
    Else
        tTmp = "00000"
    End If
    'Table Sale
    tSql = "SELECT  MAX(FTShdTransNo)  AS FTTrnNo FROM  TPSTSalHD"
    tSql = tSql & vbCrLf & "WHERE FTTmnNum = '" & tVB_CNTerminalNum & "'"
    tSql = tSql & vbCrLf & "AND FTShdStaDoc = '1'"
    If SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine) = 0 Then
        If Not (orsTemp.BOF Or orsTemp.EOF) Then
            tTmp2 = Format(Val(Format(SP_FEDtChkString(orsTemp, "FTTrnNo"))), "00000")
        Else
            tTmp2 = "00000"
        End If
    Else
        tTmp2 = "00000"
    End If
    If Val(Format(tTmp)) > Val(Format(tTmp2)) Then
        W_GETtCurTrnNoBySale = tTmp
    Else
        W_GETtCurTrnNoBySale = tTmp2
    End If
    Call SP_DATxRsNothing(orsTemp)
    Exit Function
ErrHandle:
    W_GETtCurTrnNoBySale = "00000"
    Call SP_DATxRsNothing(orsTemp)
End Function

Private Sub otbTransNo_GotFocus()
    Call SP_OTBxSelText(otbTransNo)
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
        Case 6, 8          'backspace
            'none
        Case 46             'ทศนิยม
            KeyAscii = 0
        Case 13
            Call ocmOK_Click
        Case Else
            If Not IsNumeric(Chr(KeyAscii)) Then
                KeyAscii = 0
            End If
    End Select
End Sub
