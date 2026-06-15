VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.1#0"; "Codejock.SuiteCtrls.v11.1.3.ocx"
Begin VB.Form wCNReturnMall 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mall Card"
   ClientHeight    =   3705
   ClientLeft      =   150
   ClientTop       =   2835
   ClientWidth     =   8940
   Icon            =   "wCNReturnMall.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   8940
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   3255
      Index           =   2
      Left            =   0
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   -240
      Width           =   8925
      _Version        =   720897
      _ExtentX        =   15743
      _ExtentY        =   5741
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
         Index           =   3
         Left            =   6960
         TabIndex        =   7
         Top             =   2520
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
         Left            =   2280
         MaxLength       =   20
         TabIndex        =   6
         Top             =   2520
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
         Left            =   6960
         TabIndex        =   5
         Top             =   1800
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
         Left            =   2280
         MaxLength       =   20
         TabIndex        =   4
         Top             =   1800
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
         Left            =   6960
         TabIndex        =   3
         Top             =   1102
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
         Left            =   2280
         MaxLength       =   20
         TabIndex        =   2
         Top             =   1102
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
         Left            =   6960
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
         Index           =   0
         Left            =   2280
         MaxLength       =   20
         TabIndex        =   0
         Top             =   405
         Width           =   3615
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Point :"
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
         Left            =   6000
         TabIndex        =   20
         Top             =   2595
         Width           =   810
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mall Card(Point) :"
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
         TabIndex        =   19
         Top             =   2520
         Width           =   2190
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
         Left            =   6000
         TabIndex        =   18
         Top             =   1875
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
         Left            =   6000
         TabIndex        =   17
         Top             =   1170
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
         Index           =   3
         Left            =   6000
         TabIndex        =   16
         Top             =   480
         Width           =   690
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mall Card 3 :"
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
         TabIndex        =   15
         Top             =   1875
         Width           =   1605
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mall Card 2 :"
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
         TabIndex        =   14
         Top             =   1170
         Width           =   1605
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mall Card 1 :"
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
         TabIndex        =   13
         Top             =   480
         Width           =   1605
      End
   End
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   960
      Index           =   1
      Left            =   0
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   2760
      Width           =   8925
      _Version        =   720897
      _ExtentX        =   15743
      _ExtentY        =   1693
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
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
         Left            =   5640
         TabIndex        =   9
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
         Left            =   7320
         TabIndex        =   10
         Tag             =   "2;"
         Top             =   360
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
         Height          =   495
         Left            =   3960
         TabIndex        =   8
         Tag             =   "2;"
         Top             =   360
         Width           =   1455
      End
   End
End
Attribute VB_Name = "wCNReturnMall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public nW_SeqNo As Integer
Private bW_FirstLoad As Boolean
Public cW_ChkPrice As Double
Private Sub W_DEFxFindMallCard()
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
    tSql = tSql & vbCrLf & " AND FNDctNo = " & nEN_Disc11MallCrd
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
    tSql = tSql & vbCrLf & " AND FNDctNo = " & nEN_Disc11MallCrd
    If SP_SQLbExecute(tSql, orsTemp, , nEff) Then
        bW_MallANDTopALL = True
'        MsgBox "DELETE MAll Card ALL"
    End If
    Exit Sub
ErrHandle:
    Call SP_MSGShowError
End Sub

Private Sub Form_Activate()
    bW_FirstLoad = False
    Select Case Val(Format(tVB_TransactionSub))
        Case tEN_TrnOtherDisbursement
             If otbForm(3).Locked Then
                If otbAmount(3).Enabled Then otbAmount(3).SetFocus
            Else
                If otbForm(3).Enabled Then otbForm(3).SetFocus
            End If
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SP_OBJxSetNothing(Me)
End Sub

Private Sub ocmOk_Click()
Dim nCount As Integer, nLoop As Integer
'    nCount = 0
    ReDim aVB_CNMallCard(0)
    For nLoop = 0 To otbForm.UBound - 1
        If Not (Trim(otbForm(nLoop).Text) = "") And Val(Format(otbAmount(nLoop).Text)) > 0 Then
            aVB_CNMallCard(nLoop).IDCard = otbForm(nLoop).Text
            aVB_CNMallCard(nLoop).Amt = CDbl(Val(Format(otbAmount(nLoop).Text)))  'otbAmount(nLoop).Text
        Else
            aVB_CNMallCard(nLoop).IDCard = ""
            aVB_CNMallCard(nLoop).Amt = 0
        End If
        ReDim Preserve aVB_CNMallCard(nLoop + 1)
    Next nLoop
    'สำหรับ member point
    'ReDim Preserve aVB_CNMallCard(3)
    aVB_CNMallCard(3).IDCard = otbForm(3).Text
    aVB_CNMallCard(3).Amt = CDbl(Val(Format(otbAmount(3).Text)))
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
    
    '*KT 54-09-13 V.4.0.1 Req3 Phase4.0(Member Point) Member Point on Other Receive หน้าจอ Discount สำหรับ Transaction Other Disbursement ให้ทำรายการคืน Point ได้อย่างเดียว
    Select Case Val(Format(tVB_TransactionSub))
        Case tEN_TrnOtherDisbursement
            otbForm(0).Locked = True
            otbForm(0).BackColor = nVB_CNClrBLock
            otbForm(1).Locked = True
            otbForm(1).BackColor = nVB_CNClrBLock
            otbForm(2).Locked = True
            otbForm(2).BackColor = nVB_CNClrBLock
            otbAmount(0).Locked = True
            otbAmount(0).BackColor = nVB_CNClrBLock
            otbAmount(1).Locked = True
            otbAmount(1).BackColor = nVB_CNClrBLock
            otbAmount(2).Locked = True
            otbAmount(2).BackColor = nVB_CNClrBLock
    End Select
    
    If (Not aVB_CNMallCard) = True Then
        ReDim aVB_CNMallCard(3)
    End If
    
    tSql = "SELECT FTScdCardID,FCScdAmt FROM " & tVB_CNTblCD
    tSql = tSql & vbCrLf & " WHERE FTTmnNum = '" & tVB_CNTerminalNum & "'"
    tSql = tSql & vbCrLf & " AND FTShdTransNo = '" & tVB_Running & "'"
    tSql = tSql & vbCrLf & " AND FDShdTransDate = '" & tUT_SaleDate & "'"
    tSql = tSql & vbCrLf & " AND FNSdtSeqNo = " & nW_SeqNo
    tSql = tSql & vbCrLf & " AND FNDctNo = " & nEN_Disc11MallCrd
    If SP_SQLbExecute(tSql, orsTemp, , nEff) Then
        nIndex = 0
        If Not (orsTemp.BOF And orsTemp.EOF) Then
            orsTemp.MoveFirst
            Do While Not orsTemp.EOF
                otbForm(nIndex).Text = SP_FEDtChkString(orsTemp, "FTScdCardID")
                tCardID = tCardID & "," & Trim(otbForm(nIndex).Text)
                otbAmount(nIndex).Text = SP_FMTcCurPoint(SP_FEDvChkNumeric(orsTemp, "FCScdAmt"), False, True)
                
                aVB_CNMallCard(nIndex).IDCard = otbForm(nIndex).Text
                aVB_CNMallCard(nIndex).Amt = CDbl(Val(Format(otbAmount(nIndex).Text)))  'otbAmount(nLoop).Text
'                ReDim Preserve aVB_CNMallCard(nIndex + 1)
                
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
            Call W_DEFxFindMallCard
        Else
            If Not (cW_ChkPrice = 0) Then      'เท่ากับ 0 ยังไม่มีค่า
                For nIndex = 0 To UBound(aVB_CNMallCard) - 1
                     If Not (aVB_CNMallCard(nIndex).IDCard = "") Then
                        otbForm(nIndex).Text = aVB_CNMallCard(nIndex).IDCard
                        otbAmount(nIndex).Text = SP_FMTcCurPoint(aVB_CNMallCard(nIndex).Amt, False, True)
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
                Call W_DEFxFindMallCard
            Else
                Call W_DEFxFindMallCard
            End If              'If Not (cW_ChkPrice = 0) Then
        End If
    End If
    
    
    'Member Pint
    If (Not aVB_CNMallCard) = True Then
        otbForm(3).Text = ""
        otbAmount(3).Text = ""
        otbForm(3).Enabled = True
        otbAmount(3).Enabled = True
    Else
        If UBound(aVB_CNMallCard) = 3 Then
            If aVB_CNMallCard(3).IDCard <> "" Then
                otbForm(3).Text = aVB_CNMallCard(3).IDCard
                otbAmount(3).Text = aVB_CNMallCard(3).Amt
                If aVB_CNMallCard(3).IDCard <> "" Then
'                    otbForm(3).Enabled = False
'                    otbAmount(3).Enabled = False
                    otbForm(3).Locked = True
                    otbForm(3).BackColor = nVB_CNClrBLock
                    otbForm(3).TabStop = False
                End If
            End If
        End If
    End If
    

    Exit Sub
ErrHandle:
    Call SP_MSGShowError
End Sub

Private Sub ocmCancel_Click()
Dim nIndex As Integer

'    ReDim aVB_CNMallCard(0)
'    For nIndex = 0 To otbForm.UBound
'        If Val(Format(otbAmount(nIndex).Text)) = 0 Then
'            aVB_CNMallCard(nIndex).IDCard = ""
'            aVB_CNMallCard(nIndex).Amt = 0
'        End If
'        ReDim Preserve aVB_CNMallCard(nIndex + 1)
'    Next nIndex

    bVB_CNReturnOk = False
    Unload Me
End Sub

Private Sub ocmCancelAll_Click()
Dim nIndex As Integer
On Error Resume Next
    ' *KT 54-09-13 V.4.0.1 Req3 Phase4.0(Member Point) Member Point on Other Receive หน้าจอ Discount สำหรับ Transaction Other Disbursement ให้ทำรายการคืน Point ได้อย่างเดียว
    Select Case Val(Format(tVB_TransactionSub))
        Case tEN_TrnOtherDisbursement
            otbForm(3).Locked = False
            otbForm(3).BackColor = nVB_CNClrBUnLock
            otbAmount(3).Locked = False
            otbAmount(3).BackColor = nVB_CNClrBUnLock
            otbForm(3).Text = ""
            otbAmount(3).Text = ""
        Case Else
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
    End Select
    
    ReDim aVB_CNMallCard(0)
    For nIndex = 0 To otbForm.UBound
        aVB_CNMallCard(nIndex).IDCard = ""
        aVB_CNMallCard(nIndex).Amt = 0
        ReDim Preserve aVB_CNMallCard(nIndex + 1)
    Next nIndex
    
    Select Case Val(Format(tVB_TransactionSub))
        Case tEN_TrnOtherDisbursement
            Call SP_OBJxSetFocus(otbForm(3))
        Case Else
            Call SP_OBJxSetFocus(otbForm(0))
    End Select
    
    Call W_DEFxDelMallCardAll
End Sub

Private Sub otbAmount_Change(Index As Integer)
    If bW_FirstLoad = False Then
        If otbAmount(Index).Text <> "" Then
            If Index <> 3 Then
                otbAmount(Index).Text = SP_DEFtChangeNum(Trim(otbAmount(Index).Text))
                otbAmount(Index).SelStart = Len(otbAmount(Index).Text)
            End If
        End If
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
