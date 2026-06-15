VERSION 5.00
Begin VB.Form wCNSplash 
   BorderStyle     =   0  'None
   Caption         =   "Please wait..."
   ClientHeight    =   2745
   ClientLeft      =   8625
   ClientTop       =   3780
   ClientWidth     =   5805
   Icon            =   "wCNSplash.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "wCNSplash.frx":6852
   ScaleHeight     =   2745
   ScaleWidth      =   5805
   ShowInTaskbar   =   0   'False
   Begin VB.Timer otmClosed 
      Enabled         =   0   'False
      Interval        =   400
      Left            =   4920
      Top             =   0
   End
   Begin VB.Timer otmSplash 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   5400
      Top             =   0
   End
   Begin VB.PictureBox opgSplash 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   1560
      ScaleHeight     =   285
      ScaleWidth      =   2730
      TabIndex        =   0
      Top             =   1560
      Width           =   2730
   End
   Begin VB.Label olaForm 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "System is busy now,Please wait a moment. We will try request again. "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Tag             =   $"wCNSplash.frx":F32B
      Top             =   2040
      Visible         =   0   'False
      Width           =   5535
   End
End
Attribute VB_Name = "wCNSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private m_cProgress As New Collection
Public tW_Caption$
Public nW_Tag As Long
Public nW_Time As Long

Private Sub Form_Deactivate()
    otmSplash.Enabled = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = 0
End Sub

Private Sub Form_Load()
    Dim cProgress As cProgressBar
On Error Resume Next
   DoEvents

   Call SP_FRMxChildCenter(Me)
   
   ' Progress bar in picProgress(0) with XP Style:
    Set cProgress = New cProgressBar
    cProgress.DrawObject = opgSplash
    cProgress.Min = 0
    cProgress.Max = 100
    cProgress.Tag = IIf(nW_Tag <= 0, 5, nW_Tag)
    If AP_APIbIsWindowsXP Then
        cProgress.XpStyle = True
    Else
        cProgress.BarColor = &HC0FFC0
    End If
    cProgress.ShowText = True
    m_cProgress.Add cProgress, "opgSplash"
   
   If tW_Caption = "" Then
        olaForm.Visible = False
   Else
        olaForm.Caption = tW_Caption
        cProgress.Text = olaForm.Caption
        olaForm.Visible = True
   End If
   otmSplash.Interval = IIf(nW_Time <= 0, 20, nW_Time)
   otmSplash.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrHandle
    otmClosed.Enabled = False
    otmSplash.Enabled = False
    Call SP_OBJxSetNothing(Me)
    Set wCNSplash = Nothing
ErrHandle:
    
End Sub

Private Sub otmClosed_Timer()
'    Unload wCNSplash
End Sub

Private Sub otmSplash_Timer()
Dim cProgress As cProgressBar
    DoEvents
   'If Not m_cProgress Is Nothing Then
    '  For Each cProgress In m_cProgress
      Set cProgress = m_cProgress.Item("opgSplash")
         With cProgress
            .Value = .Value + .Tag
            If .ShowText Then .Text = CLng(.Percent) & "%"
            If .Value >= .Max Then
                Unload Me
                Exit Sub
               .Tag = -1 * Abs(.Tag)
            Else
                If .Value < 1 Then .Tag = Abs(.Tag)
            End If
         End With
     ' Next
   'End If

End Sub
