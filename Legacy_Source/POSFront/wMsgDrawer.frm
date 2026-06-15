VERSION 5.00
Begin VB.Form wMsgDrawer 
   BackColor       =   &H80000010&
   BorderStyle     =   0  'None
   Caption         =   "Drawer"
   ClientHeight    =   3480
   ClientLeft      =   5475
   ClientTop       =   2190
   ClientWidth     =   6720
   Icon            =   "wMsgDrawer.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   6720
   ShowInTaskbar   =   0   'False
   Begin VB.Timer otmDrawer 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   120
      Top             =   120
   End
   Begin VB.Shape oshForm 
      BorderWidth     =   8
      Height          =   3495
      Left            =   0
      Top             =   0
      Width           =   6735
   End
   Begin VB.Label olaForm 
      Alignment       =   2  'Center
      BackColor       =   &H80000010&
      Caption         =   "Close the Cash Drawer"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   30
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Index           =   1
      Left            =   0
      TabIndex        =   1
      Top             =   1920
      Width           =   6615
   End
   Begin VB.Label olaForm 
      Alignment       =   2  'Center
      BackColor       =   &H80000010&
      Caption         =   "Cash Drawer"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   39
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   6615
   End
End
Attribute VB_Name = "wMsgDrawer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
On Error GoTo ErrHandle
    Call SP_FRMxChildCenter(Me)
    otmDrawer.Enabled = True
    Exit Sub
ErrHandle:

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrHandle
    Call SP_OBJxSetNothing(Me)
    Exit Sub
ErrHandle:

End Sub

Private Sub otmDrawer_Timer()
On Error GoTo ErrHandle
    If Not oOpos.DrawerOpen Then
        otmDrawer.Enabled = False
        Unload Me
    End If
    Exit Sub
ErrHandle:
    otmDrawer.Enabled = False
    Unload Me
End Sub
