VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.1#0"; "Codejock.SuiteCtrls.v11.1.3.ocx"
Begin VB.Form wOver 
   BorderStyle     =   0  'None
   Caption         =   "Over/Short;klg_OverShort"
   ClientHeight    =   1995
   ClientLeft      =   5085
   ClientTop       =   5160
   ClientWidth     =   5040
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1995
   ScaleWidth      =   5040
   ShowInTaskbar   =   0   'False
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   2295
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   5055
      _Version        =   720897
      _ExtentX        =   8916
      _ExtentY        =   4048
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
      PaintManager.Position=   2
      Begin VB.CommandButton ocmAccept 
         Caption         =   "Accept;klg_OverAccept"
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
         TabIndex        =   2
         Tag             =   "2;"
         Top             =   1320
         Width           =   1455
      End
      Begin VB.CommandButton ocmTotal 
         Caption         =   "Total;klg_OverTotal"
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
         Left            =   1920
         TabIndex        =   1
         Tag             =   "2;"
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label olaForm 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "- $11.77"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   360
         TabIndex        =   4
         Top             =   720
         Width           =   4335
      End
      Begin VB.Label olaForm 
         BackStyle       =   0  'Transparent
         Caption         =   "Amount Input is Short."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   3
         Top             =   240
         Width           =   4335
      End
   End
End
Attribute VB_Name = "wOver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public tW_Msg As String
Public cW_Amount As String

Private Sub Form_Load()
    Call SP_FRMxChildCenter(Me)
    Call SP_FRMxGetCaptionFrmDB2Tag(Me)
    Call SP_FRMxSetCapForm(Me)
    olaForm(0).Caption = tW_Msg
    olaForm(1).Caption = CStr(cW_Amount)
End Sub

Private Sub ocmAccept_Click()
    bVB_OverShort = True
    Unload Me
End Sub

Private Sub ocmTotal_Click()
    bVB_OverShort = False
    Unload Me
End Sub
