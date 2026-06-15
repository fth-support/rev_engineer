VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.1#0"; "Codejock.SuiteCtrls.v11.1.3.ocx"
Begin VB.Form wCNChgOptDate 
   Caption         =   "Operation Date;klg_Title"
   ClientHeight    =   2355
   ClientLeft      =   4395
   ClientTop       =   4020
   ClientWidth     =   4710
   ControlBox      =   0   'False
   Icon            =   "wCNChgOptDate.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2355
   ScaleWidth      =   4710
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   1080
      Index           =   2
      Left            =   0
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1560
      Width           =   4680
      _Version        =   720897
      _ExtentX        =   8255
      _ExtentY        =   1905
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
         Left            =   1560
         TabIndex        =   1
         Tag             =   "2;"
         Top             =   120
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
         Left            =   3120
         TabIndex        =   2
         Tag             =   "2;"
         Top             =   120
         Width           =   1455
      End
   End
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   1800
      Index           =   0
      Left            =   0
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   -240
      Width           =   4680
      _Version        =   720897
      _ExtentX        =   8255
      _ExtentY        =   3175
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
      Begin MSComCtl2.DTPicker odpOptDate 
         Height          =   615
         Left            =   960
         TabIndex        =   0
         Top             =   960
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   1085
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   58785793
         CurrentDate     =   40000
      End
      Begin VB.Label olaForm 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Confirm Operation Date;klg_Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   14.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   5
         Tag             =   "2;"
         Top             =   480
         Width           =   4695
      End
   End
End
Attribute VB_Name = "wCNChgOptDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)             '*Ao 53-01-14
'    If KeyCode = 13 Then SendKeys "{Tab}"
'End Sub

Private Sub Form_Load()
    Call SP_FRMxChildCenter(Me)
    Call SP_FRMxGetCaptionFrmDB2Tag(Me)
    Call SP_FRMxSetCapForm(Me)
    odpOptDate.Value = Format(Date, "dd/MM/yyyy")
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrHandle
    Call SP_OBJxSetNothing(Me)
ErrHandle:

End Sub

Private Sub ocmCancel_Click()
    bVB_CNChgDptDate = False
    Unload Me
End Sub

Private Sub ocmOK_Click()
    bVB_CNChgDptDate = True
    tUT_SaleDate = Format(odpOptDate.Value, "YYYY/MM/DD")
    Unload Me
End Sub


Private Sub odpOptDate_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 13
            If IsDate(odpOptDate.Value) Then
                SendKeys "{Tab}"
            End If
    End Select
End Sub
