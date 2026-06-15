VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form wCNMsgBox 
   ClientHeight    =   1020
   ClientLeft      =   6210
   ClientTop       =   3705
   ClientWidth     =   4185
   Icon            =   "wCNMsgBox.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1020
   ScaleWidth      =   4185
   Begin MSComctlLib.ImageList oilForm 
      Left            =   120
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "wCNMsgBox.frx":058A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "wCNMsgBox.frx":0BBB
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "wCNMsgBox.frx":11C8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image oimForm 
      Height          =   495
      Left            =   200
      Top             =   240
      Width           =   615
   End
   Begin VB.Label olaForm 
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   360
      Width           =   3015
   End
End
Attribute VB_Name = "wCNMsgBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*KT 53-04-09 PH3-ML-SF-112 Allow Scanner when alert message

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim tFileMedia$
On Error GoTo ErrHandle '*KT 67-01-04 runtime error  "stop statement encounter."
    tFileMedia = "C:\WINDOWS\Media\Windows XP Critical Stop.wav"            '*Ao 53-03-15 File เสียง
    Select Case KeyCode
        Case 27
            Unload Me
        Case 13
            '*Ao 53-03-15
            'คือให้ PC Speaker ส่งเสียงที่ความถี่ 3000Hz (3KHz) ด้วยระยะเวลา 0.3 วินาที (300mS)
            If SP_FLEbManFile(tFileMedia, nEN_Fle1Exist) Then 'ตรวจสอบว่ามี file มีหรือไม่
                Call PlaySound(tFileMedia, 0&, SND_ASYNC)
            Else
                Beep 5000, 400
            End If
        Case Else
            KeyCode = 0
    End Select
    Exit Sub
ErrHandle:
    
End Sub

Private Sub Form_Load()
    Call SP_FRMxChildCenter(Me)
    Call W_FRMxInitial
End Sub

Private Sub W_FRMxInitial()
    Me.Caption = tVB_CNMsgAppTitle
    olaForm.Caption = tVB_CNMsgFrm
    Select Case tVB_CNMsgStyle
        Case "16"
            oimForm.Picture = oilForm.ListImages(1).Picture
        Case "64"
            oimForm.Picture = oilForm.ListImages(2).Picture
        Case Else
            oimForm.Picture = oilForm.ListImages(3).Picture
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrHandle
     tVB_CNMsgAppTitle = ""
     tVB_CNMsgFrm = ""
     tVB_CNMsgStyle = ""
     Call SP_OBJxSetNothing(Me)
     Exit Sub
ErrHandle:
    
End Sub

