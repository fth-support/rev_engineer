VERSION 5.00
Begin VB.Form wCNMsgSlip 
   BackColor       =   &H80000010&
   BorderStyle     =   0  'None
   Caption         =   "wCNMsgSlip"
   ClientHeight    =   5160
   ClientLeft      =   6075
   ClientTop       =   2955
   ClientWidth     =   7215
   LinkTopic       =   "Form1"
   ScaleHeight     =   5160
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   Begin VB.Timer otmDrawer 
      Interval        =   500
      Left            =   120
      Top             =   80
   End
   Begin VB.Shape oshForm 
      BorderWidth     =   8
      Height          =   5175
      Left            =   0
      Top             =   0
      Width           =   7185
   End
   Begin VB.Label olaForm 
      Alignment       =   2  'Center
      BackColor       =   &H80000010&
      Caption         =   "Out of paper"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   30
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   2880
      Width           =   6855
   End
   Begin VB.Label olaForm 
      Alignment       =   2  'Center
      BackColor       =   &H80000010&
      Caption         =   "Mesage Error"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   39
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   6975
   End
End
Attribute VB_Name = "wCNMsgSlip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private tW_SlipPrinter As String
Private tW_Turnoff As String
Private tW_Error  As String

Private Sub Form_Load()
Dim nLevel As Long

On Error GoTo ErrHandle
    Call SP_FRMxChildCenter(Me)
    tW_SlipPrinter = "Slip printer is power off."
'    tW_Turnoff = "Turn off"
    tW_Turnoff = "Please power on the Slip Printer." '*Eaw 56-07-23 CommSheet TK-ISS3000-274 TR FAPL-053 00.00.02.xls
    tW_Error = "Slip printer error"
    olaForm(0).Caption = tW_SlipPrinter
'    nStatus = oOposSlip.nGetStatusUpdateEvent()
    With oOposSlip
        nLevel = .nCheckHealth(OPOS_CH_INTERNAL)
        If nLevel = OPOS_E_ILLEGAL Then
            olaForm(1).Caption = tW_Turnoff
        Else
            olaForm(1).Caption = tW_Error
        End If
    
    End With
    Exit Sub
ErrHandle:

End Sub

Private Sub Form_Unload(Cancel As Integer)

On Error GoTo ErrHandle
    otmDrawer.Enabled = False
   '*Eaw 56-07-23 CommSheet TK-ISS3000-274 TR FAPL-053 00.00.02.xls
    oOposSlip.ClosePrint
    oOposSlip.OpenPrint tVB_CashSlipPrinter, nVB_MaxCharSlip, nVB_CashCodePage 'Open Printer OPOS
    oOposSlip.SlpLineChars = nVB_MaxCharSlip

    Call SP_OBJxSetNothing(Me)
    Exit Sub
ErrHandle:

End Sub

Private Sub otmDrawer_Timer()
Dim nLevel As Long

On Error GoTo ErrHandle
    
    olaForm(0).Caption = tW_SlipPrinter
'    nStatus = oOposSlip.nGetStatusUpdateEvent()
'    olaForm(1).Caption = SP_GETtPrinterOPOSStatus(nStatus)
    
    With oOposSlip
        nLevel = .nCheckHealth(OPOS_CH_INTERNAL)
        If nLevel = 0 Then
            Unload Me
        Else
            If nLevel = OPOS_E_ILLEGAL Then
                olaForm(1).Caption = tW_Turnoff
            Else
                olaForm(1).Caption = tW_Error
            End If
        End If
    
    End With
    Exit Sub
ErrHandle:
    Unload Me
End Sub
