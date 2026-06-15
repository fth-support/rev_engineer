VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.1#0"; "Codejock.SuiteCtrls.v11.1.3.ocx"
Begin VB.Form wCNReturnCredit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Credit Card"
   ClientHeight    =   3960
   ClientLeft      =   3045
   ClientTop       =   8190
   ClientWidth     =   8460
   Icon            =   "wCNReturnCredit.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   8460
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   3495
      Index           =   2
      Left            =   0
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   -240
      Width           =   8445
      _Version        =   720897
      _ExtentX        =   14896
      _ExtentY        =   6165
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
      Begin VB.TextBox otbAmount 
         Alignment       =   1  'Right Justify
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
         Index           =   4
         Left            =   6480
         TabIndex        =   9
         Top             =   2760
         Width           =   1815
      End
      Begin VB.TextBox otbForm 
         Alignment       =   1  'Right Justify
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
         Index           =   4
         Left            =   2040
         MaxLength       =   20
         TabIndex        =   8
         Top             =   2760
         Width           =   3615
      End
      Begin VB.TextBox otbAmount 
         Alignment       =   1  'Right Justify
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
         Index           =   3
         Left            =   6480
         TabIndex        =   7
         Top             =   2169
         Width           =   1815
      End
      Begin VB.TextBox otbForm 
         Alignment       =   1  'Right Justify
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
         Index           =   3
         Left            =   2040
         MaxLength       =   20
         TabIndex        =   6
         Top             =   2169
         Width           =   3615
      End
      Begin VB.TextBox otbForm 
         Alignment       =   1  'Right Justify
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
         Left            =   2040
         MaxLength       =   20
         TabIndex        =   0
         Top             =   405
         Width           =   3615
      End
      Begin VB.TextBox otbAmount 
         Alignment       =   1  'Right Justify
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
         Left            =   6480
         TabIndex        =   1
         Top             =   405
         Width           =   1815
      End
      Begin VB.TextBox otbForm 
         Alignment       =   1  'Right Justify
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
         Left            =   2040
         MaxLength       =   20
         TabIndex        =   2
         Top             =   993
         Width           =   3615
      End
      Begin VB.TextBox otbAmount 
         Alignment       =   1  'Right Justify
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
         Left            =   6480
         TabIndex        =   3
         Top             =   960
         Width           =   1815
      End
      Begin VB.TextBox otbForm 
         Alignment       =   1  'Right Justify
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
         Index           =   2
         Left            =   2040
         MaxLength       =   20
         TabIndex        =   4
         Top             =   1581
         Width           =   3615
      End
      Begin VB.TextBox otbAmount 
         Alignment       =   1  'Right Justify
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
         Index           =   2
         Left            =   6480
         TabIndex        =   5
         Top             =   1581
         Width           =   1815
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amt :"
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
         Index           =   9
         Left            =   5760
         TabIndex        =   23
         Top             =   2835
         Width           =   690
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amt :"
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
         Index           =   8
         Left            =   5760
         TabIndex        =   22
         Top             =   2244
         Width           =   690
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Credit Card 5 :"
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
         Index           =   7
         Left            =   120
         TabIndex        =   21
         Top             =   2835
         Width           =   1845
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Credit Card 4 :"
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
         Index           =   6
         Left            =   120
         TabIndex        =   20
         Top             =   2250
         Width           =   1845
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Credit Card 1 :"
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
         Left            =   120
         TabIndex        =   19
         Top             =   480
         Width           =   1845
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Credit Card 2 :"
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
         Left            =   120
         TabIndex        =   18
         Top             =   1068
         Width           =   1845
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Credit Card 3 :"
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
         Index           =   2
         Left            =   120
         TabIndex        =   17
         Top             =   1656
         Width           =   1845
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amt :"
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
         Index           =   3
         Left            =   5760
         TabIndex        =   16
         Top             =   480
         Width           =   690
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amt :"
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
         Index           =   4
         Left            =   5760
         TabIndex        =   15
         Top             =   1068
         Width           =   690
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amt :"
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
         Index           =   5
         Left            =   5760
         TabIndex        =   14
         Top             =   1656
         Width           =   690
      End
   End
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   960
      Index           =   1
      Left            =   0
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   3000
      Width           =   8445
      _Version        =   720897
      _ExtentX        =   14896
      _ExtentY        =   1693
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
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
         Height          =   495
         Left            =   3480
         TabIndex        =   10
         Tag             =   "2;"
         Top             =   360
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
         Height          =   495
         Left            =   6840
         TabIndex        =   12
         Tag             =   "2;"
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton ocmCancelAll 
         Caption         =   "&Cancel All"
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
         Left            =   5160
         TabIndex        =   11
         Tag             =   "2;"
         Top             =   360
         Width           =   1455
      End
   End
End
Attribute VB_Name = "wCNReturnCredit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public nW_SeqNo As Integer
Private bW_FirstLoad As Boolean
Public cW_ChkPrice As Double
'Public nW_DiscType As Integer
Private Sub W_DEFxFindOnTop()
Dim orsTemp As ADODB.Recordset
Dim nIndex As Integer, nLoop As Integer
Dim tSql As String
Dim nEff As Long
Dim bChkDup As Boolean

On Error GoTo ErrHandle
    tSql = "SELECT DISTINCT FTScdCardID FROM " & tVB_CNTblCD
    tSql = tSql & vbCrLf & " WHERE FTTmnNum = '" & tVB_CNTerminalNum & "'"
    tSql = tSql & vbCrLf & " AND FTShdTransNo = '" & tVB_Running & "'"
    tSql = tSql & vbCrLf & " AND FDShdTransDate = '" & tUT_SaleDate & "'"
    tSql = tSql & vbCrLf & " AND FNDctNo = " & nEN_Disc15Ontop
'    tSql = tSql & vbCrLf & " AND NOT (FNSdtSeqNo = " & nW_SeqNo & ")"
    If SP_SQLbExecute(tSql, orsTemp, , nEff) Then
        If Not (orsTemp.BOF And orsTemp.EOF) Then
            orsTemp.MoveFirst
            Do While Not orsTemp.EOF
                For nIndex = 0 To otbForm.UBound
                    If (Trim(otbForm(nIndex).Text) = "") And (bChkDup = False) Then
                        For nLoop = 0 To otbForm.UBound
                            If otbForm(nLoop).Text = SP_FEDtChkString(orsTemp, "FTScdCardID") Then
                                bChkDup = False
                                Exit For
                            Else
                                bChkDup = True
                            End If
                        Next nLoop
                        If bChkDup Then
                            otbForm(nIndex).Text = SP_FEDtChkString(orsTemp, "FTScdCardID")
'                            otbForm(nIndex).Enabled = False
                            otbForm(nIndex).Locked = True
                            otbForm(nIndex).BackColor = nVB_CNClrBLock
                            otbForm(nIndex).TabStop = False
                            otbAmount(nIndex).Text = SP_FMTcCurPoint(0, False, True)
                            bChkDup = True
                        End If
                    End If
                Next nIndex
                bChkDup = False
                orsTemp.MoveNext
            Loop
       End If   'SP_SQLbExecute
    End If
Exit Sub
ErrHandle:
    Call SP_MSGShowError
End Sub
Private Sub W_DEFxDelMallCardAll()
Dim orsTemp As ADODB.Recordset
Dim tSql As String
Dim nEff As Long
On Error GoTo ErrHandle
    tSql = "DELETE FROM  " & tVB_CNTblCD
    tSql = tSql & vbCrLf & " WHERE FTTmnNum = '" & tVB_CNTerminalNum & "'"
    tSql = tSql & vbCrLf & " AND FTShdTransNo = '" & tVB_Running & "'"
    tSql = tSql & vbCrLf & " AND FDShdTransDate = '" & tUT_SaleDate & "'"
    tSql = tSql & vbCrLf & " AND FNDctNo = " & nEN_Disc15Ontop
    If SP_SQLbExecute(tSql, orsTemp, , nEff) Then
        bW_MallANDTopALL = True
'        MsgBox "DELETE ONTOP"
    End If
Exit Sub
ErrHandle:
    Call SP_MSGShowError
End Sub
Private Sub Form_Activate()
    bW_FirstLoad = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SP_OBJxSetNothing(Me)
End Sub

Private Sub ocmOk_Click()
Dim nCount As Integer, nLoop As Integer
'    nCount = 0
    ReDim aVB_CNOnTop(0)
    For nLoop = 0 To otbForm.UBound
        If Not (Trim(otbForm(nLoop).Text) = "") And Val(otbAmount(nLoop).Text) > 0 Then
'            ReDim Preserve aVB_CNOnTop(nCount)
            aVB_CNOnTop(nLoop).IDCard = otbForm(nLoop).Text
            aVB_CNOnTop(nLoop).Amt = Val(otbAmount(nLoop).Text)
        Else
'            ReDim Preserve aVB_CNOnTop(nCount)
            aVB_CNOnTop(nLoop).IDCard = ""
            aVB_CNOnTop(nLoop).Amt = 0
        End If
        ReDim Preserve aVB_CNOnTop(nLoop + 1)
    Next nLoop
    bVB_CNReturnOk = True
    Unload Me
End Sub

Private Sub Form_Load()
Dim orsTemp As ADODB.Recordset
Dim nIndex As Integer
Dim tSql As String
Dim nEff As Long
Dim bChkDup As Boolean
Dim tCardID As String
On Error GoTo ErrHandle
    Call SP_FRMxChildCenter(Me)
    
    ReDim aVB_CNOnTop(0)
    
    tSql = "SELECT FTScdCardID,FCScdAmt FROM " & tVB_CNTblCD
    tSql = tSql & vbCrLf & " WHERE FTTmnNum = '" & tVB_CNTerminalNum & "'"
    tSql = tSql & vbCrLf & " AND FTShdTransNo = '" & tVB_Running & "'"
    tSql = tSql & vbCrLf & " AND FDShdTransDate = '" & tUT_SaleDate & "'"
    tSql = tSql & vbCrLf & " AND FNSdtSeqNo = " & nW_SeqNo
    tSql = tSql & vbCrLf & " AND FNDctNo = " & nEN_Disc15Ontop
    If SP_SQLbExecute(tSql, orsTemp, , nEff) Then
        nIndex = 0
        If Not (orsTemp.BOF And orsTemp.EOF) Then
            orsTemp.MoveFirst
            Do While Not orsTemp.EOF
                otbForm(nIndex).Text = SP_FEDtChkString(orsTemp, "FTScdCardID")
                tCardID = tCardID & "," & Trim(otbForm(nIndex).Text)
                otbAmount(nIndex).Text = SP_FMTcCurPoint(SP_FEDvChkNumeric(orsTemp, "FCScdAmt"), False, True)
                
                aVB_CNOnTop(nIndex).IDCard = otbForm(nIndex).Text
                aVB_CNOnTop(nIndex).Amt = CDbl(Val(Format(otbAmount(nIndex).Text)))  'otbAmount(nLoop).Text
                ReDim Preserve aVB_CNOnTop(nIndex + 1)
                
                If Not (Trim(otbForm(nIndex).Text) = "") Then
'                    otbForm(nIndex).Enabled = False
'                    otbAmount(nIndex).Enabled = False
                    otbForm(nIndex).Locked = True
                    otbForm(nIndex).BackColor = nVB_CNClrBLock
                    otbForm(nIndex).TabStop = False
                End If
                nIndex = nIndex + 1
                orsTemp.MoveNext
            Loop
            Call W_DEFxFindOnTop
        Else
            If Not (cW_ChkPrice = 0) Then
                For nIndex = 0 To UBound(aVB_CNOnTop)
                    If Not (aVB_CNOnTop(nIndex).IDCard = "") Then
                        otbForm(nIndex).Text = aVB_CNOnTop(nIndex).IDCard
                        otbAmount(nIndex).Text = SP_FMTcCurPoint(aVB_CNOnTop(nIndex).Amt, False, True)
                    End If
                    If Not (Trim(otbForm(nIndex).Text) = "") Then
'                        otbForm(nIndex).Enabled = False
'                        otbAmount(nIndex).Enabled = False
                         otbForm(nIndex).Locked = True
                         otbForm(nIndex).BackColor = nVB_CNClrBLock
                         otbForm(nIndex).TabStop = False
                    End If
                Next nIndex
                Call W_DEFxFindOnTop
            Else
                Call W_DEFxFindOnTop
            End If
        End If
    End If
    Exit Sub
ErrHandle:
    Call SP_MSGShowError
End Sub
Private Sub ocmCancel_Click()
'Dim nIndex As Integer

'    For nIndex = 0 To otbForm.UBound
'        ReDim Preserve aVB_CNOnTop(nIndex)
'        If Val(otbAmount(nIndex).Text) = 0 Then
'            aVB_CNOnTop(nIndex).IDCard = ""
'            aVB_CNOnTop(nIndex).Amt = 0
'        End If
'    Next nIndex

    bVB_CNReturnOk = False
    Unload Me
End Sub
Private Sub ocmCancelAll_Click()
Dim nIndex As Integer
'    tSql = "DELECT FROM " & tVB_CNTblCD
'    tSql = tSql & vbCrLf & " WHERE FNSdtSeqNo = " & nW_SeqNo
'    tSql = tSql & vbCrLf & " AND FNDctNo = " & nW_DiscType
    For nIndex = 0 To otbForm.UBound
        otbForm(nIndex).Text = ""
        otbAmount(nIndex).Text = ""
'        otbForm(nIndex).Enabled = True
'        otbAmount(nIndex).Enabled = True
          otbForm(nIndex).Locked = False
          otbForm(nIndex).BackColor = nVB_CNClrBUnLock
          otbForm(nIndex).TabStop = True
    Next nIndex
    ReDim aVB_CNOnTop(0)
    For nIndex = 0 To otbForm.UBound
        aVB_CNOnTop(nIndex).IDCard = ""
        aVB_CNOnTop(nIndex).Amt = 0
        ReDim Preserve aVB_CNOnTop(nIndex + 1)
    Next nIndex
    Call SP_OBJxSetFocus(otbForm(0))
    Call W_DEFxDelMallCardAll
End Sub

Private Sub otbAmount_Change(Index As Integer)
    If bW_FirstLoad = False Then
        otbAmount(Index).Text = SP_DEFtChangeNum(Trim(otbAmount(Index).Text))
        otbAmount(Index).SelStart = Len(otbAmount(Index).Text)
    End If
End Sub

Private Sub otbAmount_GotFocus(Index As Integer)
    otbAmount(Index).SelStart = 0
    otbAmount(Index).SelLength = Len(otbAmount(Index).Text)
End Sub

Private Sub otbAmount_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
     Select Case KeyCode
        Case vbKeyDelete         'delete
            otbAmount(Index).Text = ""
    End Select
End Sub

Private Sub otbAmount_KeyPress(Index As Integer, KeyAscii As Integer)
    Call UT_ChkKeyAscii(KeyAscii)
    Select Case KeyAscii
        Case 8, 6           'backspace, delete
            'do nothing
        Case 46             '·È¹ÔÂÁ
            KeyAscii = 0
        Case 13
'            Call ocmOk_Click
            SendKeys "{TAB}"
        Case Else
            If Not IsNumeric(Chr(KeyAscii)) Then
                KeyAscii = 0
            End If
    End Select

End Sub
Private Sub otbForm_KeyPress(Index As Integer, KeyAscii As Integer)
    Call UT_ChkKeyAscii(KeyAscii)
    Select Case KeyAscii
        Case 8, 6           'backspace, delete
            'do nothing
        Case 46             '·È¹ÔÂÁ
            KeyAscii = 0
        Case 13
'            Call ocmOk_Click
            SendKeys "{TAB}"
        Case Else
            If Not IsNumeric(Chr(KeyAscii)) Then
                KeyAscii = 0
            End If
    End Select

End Sub
