VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.1#0"; "Codejock.SuiteCtrls.v11.1.3.ocx"
Object = "{9D6F2CC0-07CE-11CF-A800-0020AF9EBB73}#1.0#0"; "Scanner.ocx"
Begin VB.Form wCNTMkdNo 
   Caption         =   "Transaction Markdown;klg_Title"
   ClientHeight    =   2505
   ClientLeft      =   7095
   ClientTop       =   6600
   ClientWidth     =   4380
   ControlBox      =   0   'False
   Icon            =   "wCNTMkdNo.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2505
   ScaleWidth      =   4380
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   1840
      Index           =   2
      Left            =   0
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   -240
      Width           =   4365
      _Version        =   720897
      _ExtentX        =   7699
      _ExtentY        =   3246
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
      Begin VB.TextBox otbNumber 
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
         Left            =   120
         MaxLength       =   20
         TabIndex        =   0
         Text            =   "01234567890123456789"
         Top             =   960
         Width           =   4095
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Number:;klg_No"
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
         TabIndex        =   4
         Tag             =   "2;"
         Top             =   480
         Width           =   2085
      End
   End
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   1170
      Index           =   1
      Left            =   0
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1350
      Width           =   4365
      _Version        =   720897
      _ExtentX        =   7699
      _ExtentY        =   2064
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
      Begin SCANNERLib.Scanner oScanner 
         Left            =   240
         Top             =   240
         _Version        =   65536
         _ExtentX        =   741
         _ExtentY        =   741
         _StockProps     =   0
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
   End
End
Attribute VB_Name = "wCNTMkdNo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
On Error GoTo ErrHandle
    Call SP_FRMxChildCenter(Me)
    Call SP_FRMxGetCaptionFrmDB2Tag(Me)
    Call SP_FRMxSetCapForm(Me)
    Call W_FRMxSetInitial
'    otbID.Text = ""
    If bVB_OPOSScannerEnable Then
        With oScanner
            .Open tVB_Scanner
            .Claim 0 'TeamPos2000
            'Enable the device.
            .DeviceEnabled = True
            'Enable the event.
            .DataEventEnabled = True
            .AutoDisable = True
            .ClearInput
        End With
    End If
    Call SP_FRMxSetTimeCount(True, nW_TimerLock)       'จับเวลา Lock ใหม่  '*KT 52-08-17 Locked Keyboard
    If otbNumber.Text <> "" Then
        Call SP_OTBxSelText(otbNumber)
    End If
    Exit Sub
ErrHandle:
    bVB_SalemanIDChg = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrHandle
    With oScanner
        If bVB_OPOSScannerEnable Then
            If .DeviceEnabled = True Then
                .DeviceEnabled = False
                .Release
                .Close
            End If
        End If
    End With
    Call SP_OBJxSetNothing(Me)
ErrHandle:
End Sub

Private Sub ocmCancel_Click()
    tVB_CNTranMDownNo = ""
    Unload Me
End Sub

Private Sub ocmOk_Click()
    tVB_CNTranMDownNo = Trim(otbNumber.Text)
    Unload Me
End Sub

Private Sub oScanner_DataEvent(ByVal Status As Long)
On Error GoTo ErrHandle
     '*KT 52-08-17 Locked Keyboard
        With oScanner
            otbNumber.Text = .ScanData
'            DoEvents '*Ao 53-01-14   DoEvents
            If otbNumber.Text <> "" Then
                If Asc(Right(otbNumber.Text, 1)) = 13 Then
                    otbNumber.Text = Mid(otbNumber.Text, 1, Len(otbNumber.Text) - 1)
                End If
            End If
            If bVB_OPOSScannerEnable Then
                    If .DeviceEnabled = False Then
                         .Open tVB_Scanner
                        .Claim 0 'TeamPos2000
                         'Enable the device.
                         .DeviceEnabled = True
                         'Enable the event.
                         .DataEventEnabled = True
                         .AutoDisable = True
                         .ClearInput
                    Else
                        .ClearInput
                    End If
            End If
        End With
        Call SP_OBJxSetFocus(ocmOk)
    Exit Sub
ErrHandle:
    MsgBox "Scaner error" & vbCrLf & "Runtime : " & Err.Number & vbCrLf & Err.Description, vbCritical
End Sub

Private Sub W_FRMxSetInitial()
    otbNumber.Text = ""
     Select Case tVB_CNTrnDiscType
        Case "1893"
            Me.Caption = Me.Caption & "1"
        Case "1892"
            Me.Caption = Me.Caption & "2"
        Case "1891"
            Me.Caption = Me.Caption & "3"
        Case "1890"
            Me.Caption = Me.Caption & "4"
        Case "1889"
            Me.Caption = Me.Caption & "5"
    End Select
End Sub


Private Sub otbNumber_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case Shift
        Case 0
            Select Case KeyCode
                Case vbKeyDelete         'delete
                    otbNumber.Text = ""
                Case 13
                    
                    Call ocmOk_Click
                    KeyCode = 0
            End Select
    End Select
End Sub

Private Sub otbNumber_KeyPress(KeyAscii As Integer)
    Call UT_ChkKeyAscii(KeyAscii)
End Sub
