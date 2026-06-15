VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.1#0"; "Codejock.SuiteCtrls.v11.1.3.ocx"
Begin VB.Form wCNReturnEmp 
   Caption         =   "Employee Card"
   ClientHeight    =   3015
   ClientLeft      =   5505
   ClientTop       =   9255
   ClientWidth     =   8220
   ControlBox      =   0   'False
   Icon            =   "wCNReturnEmp.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   8220
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   2535
      Index           =   2
      Left            =   0
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   -240
      Width           =   8205
      _Version        =   720897
      _ExtentX        =   14473
      _ExtentY        =   4471
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
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
         Left            =   1800
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
         Left            =   6240
         TabIndex        =   1
         Top             =   405
         Width           =   1815
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Emp Card :"
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
         TabIndex        =   8
         Top             =   480
         Width           =   1440
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
         Left            =   5520
         TabIndex        =   7
         Top             =   480
         Width           =   690
      End
   End
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   960
      Index           =   1
      Left            =   0
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2040
      Width           =   8205
      _Version        =   720897
      _ExtentX        =   14473
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
         Left            =   3120
         TabIndex        =   2
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
         Left            =   6480
         TabIndex        =   4
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
         Left            =   4800
         TabIndex        =   3
         Tag             =   "2;"
         Top             =   360
         Width           =   1455
      End
   End
End
Attribute VB_Name = "wCNReturnEmp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public nW_SeqNo As Integer
Private bW_FirstLoad As Boolean
Public cW_ChkPrice As Double

Private Sub W_DEFxFindEmpCard()
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
    tSql = tSql & vbCrLf & " AND FNDctNo = " & nEN_Disc12EmpCrd
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
       End If
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
    tSql = tSql & vbCrLf & " AND FNDctNo = " & nEN_Disc12EmpCrd
    If SP_SQLbExecute(tSql, orsTemp, , nEff) Then
'        bW_MallANDTopALL = True
    End If
    Exit Sub
ErrHandle:
    Call SP_MSGShowError
End Sub

Private Sub Form_Activate()
    bW_FirstLoad = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrHandle
    Call SP_OBJxSetNothing(Me)
    Exit Sub
ErrHandle:
    
End Sub

Private Sub ocmOk_Click()
    Dim nCount As Integer, nLoop As Integer
    ReDim aVB_CNEmpCard(0)
    For nLoop = 0 To otbForm.UBound
        If Not (Trim(otbForm(nLoop).Text) = "") And Val(otbAmount(nLoop).Text) > 0 Then
            aVB_CNEmpCard(nLoop).IDCard = otbForm(nLoop).Text
            aVB_CNEmpCard(nLoop).Amt = CDbl(Val(Format(otbAmount(nLoop).Text)))
        Else
            aVB_CNEmpCard(nLoop).IDCard = ""
            aVB_CNEmpCard(nLoop).Amt = 0
        End If
        ReDim Preserve aVB_CNEmpCard(nLoop + 1)
    Next nLoop
    bVB_CNReturnOk = True
    Unload Me
End Sub

Private Sub Form_Load()
    Dim orsTemp As ADODB.Recordset
    Dim nIndex As Integer, nLoop As Integer
    Dim tSql As String
    Dim nEff As Long
    Dim bChkDup As Boolean
    Dim tCardID As String

On Error GoTo ErrHandle
    Call SP_FRMxChildCenter(Me)
    
    bW_FirstLoad = True
    
    ReDim aVB_CNEmpCard(0)
    
    tSql = "SELECT FTScdCardID,FCScdAmt FROM " & tVB_CNTblCD
    tSql = tSql & vbCrLf & " WHERE FTTmnNum = '" & tVB_CNTerminalNum & "'"
    tSql = tSql & vbCrLf & " AND FTShdTransNo = '" & tVB_Running & "'"
    tSql = tSql & vbCrLf & " AND FDShdTransDate = '" & tUT_SaleDate & "'"
    tSql = tSql & vbCrLf & " AND FNSdtSeqNo = " & nW_SeqNo
    tSql = tSql & vbCrLf & " AND FNDctNo = " & nEN_Disc12EmpCrd
    If SP_SQLbExecute(tSql, orsTemp, , nEff) Then
        nIndex = 0
        If Not (orsTemp.BOF And orsTemp.EOF) Then
            orsTemp.MoveFirst
            Do While Not orsTemp.EOF
                otbForm(nIndex).Text = SP_FEDtChkString(orsTemp, "FTScdCardID")
                
                tCardID = tCardID & "," & Trim(otbForm(nIndex).Text)
                otbAmount(nIndex).Text = SP_FMTcCurPoint(SP_FEDvChkNumeric(orsTemp, "FCScdAmt"), False, True)
                
                aVB_CNEmpCard(nIndex).IDCard = otbForm(nIndex).Text
                aVB_CNEmpCard(nIndex).Amt = CDbl(Val(Format(otbAmount(nIndex).Text)))  'otbAmount(nLoop).Text
                ReDim Preserve aVB_CNEmpCard(nIndex + 1)
                
                If Not (Trim(otbForm(nIndex).Text) = "") Then
'                    otbForm(nIndex).Enabled = False
'                    otbAmount(nIndex).Enabled = False
                    otbForm(nIndex).Locked = True
                    otbForm(nIndex).BackColor = nVB_CNClrBLock
                    otbForm(nIndex).TabStop = False
'                    otbAmount(nIndex).Locked = True
'                    otbAmount(nIndex).BackColor = nVB_CNClrBLock
'                    otbAmount(nIndex).TabStop = False
                End If
                nIndex = nIndex + 1
                orsTemp.MoveNext
            Loop
            Call W_DEFxFindEmpCard
        Else
            If Not (cW_ChkPrice = 0) Then      'เท่ากับ 0 ยังไม่มีค่า
                For nIndex = 0 To UBound(aVB_CNEmpCard) - 1
                     If Not (aVB_CNEmpCard(nIndex).IDCard = "") Then
                        otbForm(nIndex).Text = aVB_CNEmpCard(nIndex).IDCard
                        otbAmount(nIndex).Text = SP_FMTcCurPoint(aVB_CNEmpCard(nIndex).Amt, False, True)
                     End If
                    If Not (Trim(otbForm(nIndex).Text) = "") Then
'                        otbForm(nIndex).Enabled = False
'                        otbAmount(nIndex).Enabled = False
                        otbForm(nIndex).Locked = True
                        otbForm(nIndex).BackColor = nVB_CNClrBLock
                        otbForm(nIndex).TabStop = False
'                        otbAmount(nIndex).Locked = True
'                        otbAmount(nIndex).BackColor = nVB_CNClrBLock
'                        otbAmount(nIndex).TabStop = False
                    End If
                Next nIndex
                Call W_DEFxFindEmpCard
            Else
                Call W_DEFxFindEmpCard
            End If              'If Not (cW_ChkPrice = 0) Then
        End If
    End If
    Exit Sub
ErrHandle:
    Call SP_MSGShowError
End Sub

Private Sub ocmCancel_Click()
Dim nIndex As Integer

'    ReDim aVB_CNEmpCard(0)
'    For nIndex = 0 To otbForm.UBound
'        If Val(Format(otbAmount(nIndex).Text)) = 0 Then
'            aVB_CNEmpCard(nIndex).IDCard = ""
'            aVB_CNEmpCard(nIndex).Amt = 0
'        End If
'        ReDim Preserve aVB_CNEmpCard(nIndex + 1)
'    Next nIndex

    bVB_CNReturnOk = False
    Unload Me
End Sub

Private Sub ocmCancelAll_Click()
Dim nIndex As Integer
    
    For nIndex = 0 To otbForm.UBound
        otbForm(nIndex).Text = ""
        otbAmount(nIndex).Text = ""
'        otbForm(nIndex).Enabled = True
'        otbAmount(nIndex).Enabled = True
        otbForm(nIndex).Locked = False
        otbForm(nIndex).BackColor = nVB_CNClrBUnLock
        otbForm(nIndex).TabStop = True
        otbAmount(nIndex).Locked = False
        otbAmount(nIndex).BackColor = nVB_CNClrBUnLock
        otbAmount(nIndex).TabStop = True
    Next nIndex
    ReDim aVB_CNEmpCard(0)
    For nIndex = 0 To otbForm.UBound
        aVB_CNEmpCard(nIndex).IDCard = ""
        aVB_CNEmpCard(nIndex).Amt = 0
        ReDim Preserve aVB_CNEmpCard(nIndex + 1)
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
    If Not otbAmount(Index).Locked Then
        otbAmount(Index).SelStart = 0
        otbAmount(Index).SelLength = Len(otbAmount(Index).Text)
    End If
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
        Case 46             'ทศนิยม
            KeyAscii = 0
        Case 13
            SendKeys "{TAB}"
'            Call ocmOk_Click
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
        Case 46             'ทศนิยม
            KeyAscii = 0
        Case 13
            SendKeys "{TAB}"
'            Call ocmOk_Click
        Case Else
            If Not IsNumeric(Chr(KeyAscii)) Then
                KeyAscii = 0
            End If
    End Select
End Sub


