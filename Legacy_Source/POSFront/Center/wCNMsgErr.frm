VERSION 5.00
Begin VB.Form wCNMsgErr 
   BackColor       =   &H80000010&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3570
   ClientLeft      =   1515
   ClientTop       =   6750
   ClientWidth     =   6810
   Icon            =   "wCNMsgErr.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3570
   ScaleWidth      =   6810
   ShowInTaskbar   =   0   'False
   Begin VB.Timer otmDrawer 
      Interval        =   500
      Left            =   120
      Top             =   120
   End
   Begin VB.Shape oshForm 
      BorderWidth     =   8
      Height          =   3495
      Left            =   40
      Top             =   40
      Width           =   6735
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
      Height          =   1215
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   6615
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
      Height          =   1455
      Index           =   1
      Left            =   0
      TabIndex        =   0
      Top             =   1920
      Width           =   6615
   End
End
Attribute VB_Name = "wCNMsgErr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Call SP_FRMxChildCenter(Me)
    With oOpos
        Select Case .ResultCode
            Case OPOS_E_EXTENDED
                Select Case .ResultCodeExtended
                    Case OPOS_EPTR_COVER_OPEN
                        If .PowerState = 2004 Then
                            olaForm(0) = "Printer"
                            olaForm(1) = "Turn off"
                        ElseIf .oPrinterState = 4 Then
                            olaForm(0) = "Printer"
                            olaForm(1) = "Turn off"
                        Else
                            If .CoverOpen Then
                                olaForm(0) = "Printer"
                                olaForm(1) = "Printer cover is open"
                            ElseIf .ErrorLevel = 2 Then
                                olaForm(0) = "Printer"
                                olaForm(1) = "Out of paper"
                            Else
                                Unload Me
                            End If
                        End If
                    Case OPOS_EPTR_REC_EMPTY
                        If .PowerState = 2004 Then
                            olaForm(0) = "Printer"
                            olaForm(1) = "Turn off"
                        ElseIf .oPrinterState = 4 Then
                            olaForm(0) = "Printer"
                            olaForm(1) = "Turn off"
                        Else
                            If .CoverOpen Then
                                olaForm(0) = "Printer"
                                olaForm(1) = "Printer cover is open"
                            ElseIf .ErrorLevel = 2 Then
                                olaForm(0) = "Printer"
                                olaForm(1) = "Out of paper"
                            Else
                                Unload Me
                            End If
                        End If
                End Select
            Case OPOS_SUCCESS
                If .PowerState = 2004 Then
                    If .CoverOpen Then
                        olaForm(0) = "Printer"
                        olaForm(1) = "Printer cover is open"
                    ElseIf .ErrorLevel = 2 Then
                        olaForm(0) = "Printer"
                        olaForm(1) = "Out of paper"
                    Else
                        olaForm(0) = "Printer"
                        olaForm(1) = "Turn off"
                    End If

                ElseIf .oPrinterState = 4 Then
                    olaForm(0) = "Printer"
                    olaForm(1) = "Turn off"
                Else
                    If .CoverOpen Then
                        olaForm(0) = "Printer"
                        olaForm(1) = "Printer cover is open"
                    ElseIf .ErrorLevel = 2 Then
                        olaForm(0) = "Printer"
                        olaForm(1) = "Out of paper"
                    End If
                End If
        End Select
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim nI  As Long
    Dim LF$
    
On Error GoTo ErrHandle
    LF = Chr(10)
    With oOpos
        '*KT 51-09-2008 กรณีที่เปิด Printer ให้ทำการตัดกระดาษแล้วพิมพ์ ใหม่
        '-----------------------------------------------------------------
        '*Ao 2013/01/24   Print FRONTTECH
'         If bVB_ImageLogo Then
'            .TransactionPrint PTR_S_RECEIPT, PTR_TP_TRANSACTION 'end print by OPOS
'             .Printnormal PTR_S_RECEIPT, "" & Chr(10)
'              Call oPrn.PrintLogo
'            .TransactionPrint PTR_S_RECEIPT, PTR_TP_NORMAL 'end print by OPOS
'            .CutPaper 90
'            .Printnormal PTR_S_RECEIPT, Chr(16) & Chr(20) '*KT 51-12-22 Phase 2 Clear Buffer
'         Else
'            '*KT 51-10-14 V 3.0.135/V 3.0.200 กรณีที่ไม่มีการ Set Logo
'            '--------------------------------------------------------------
'            .TransactionPrint PTR_S_RECEIPT, PTR_TP_TRANSACTION 'end print by OPOS
'            .Printnormal PTR_S_RECEIPT, "" & Chr(10)
'             For nI = 1 To 8
'                .Printnormal PTR_S_RECEIPT, "" & LF
'             Next nI
'            .TransactionPrint PTR_S_RECEIPT, PTR_TP_NORMAL 'end print by OPOS
'            .CutPaper 90
'            .Printnormal PTR_S_RECEIPT, Chr(16) & Chr(20) '*KT 51-12-22 Phase 2 Clear Buffer
'            '---------------------------------------------------------------
'         End If
        '-----------------------------------------------------------------
    End With
    otmDrawer.Enabled = False
    Call SP_OBJxSetNothing(Me)
ErrHandle:
    
End Sub

Private Sub otmDrawer_Timer()
    With oOpos
        Select Case .ResultCode
            Case OPOS_E_EXTENDED
                Select Case .ResultCodeExtended
                    Case OPOS_EPTR_COVER_OPEN
                        If .PowerState = 2004 Then
                            If .CoverOpen Then
                                olaForm(0) = "Printer"
                                olaForm(1) = "Printer cover is open"
                            ElseIf .ErrorLevel = 2 Then
                                olaForm(0) = "Printer"
                                olaForm(1) = "Out of paper"
'                            ElseIf .ErrorLevel = 1 Then
'                                olaForm(0) = "Printer"
'                                olaForm(1) = "Turn off"
                            Else
                                Unload Me
                            End If
                        ElseIf .oPrinterState = 4 Then
                            olaForm(0) = "Printer"
                            olaForm(1) = "Turn off"
                        Else
                            If .CoverOpen Then
                                olaForm(0) = "Printer"
                                olaForm(1) = "Printer cover is open"
                            ElseIf .ErrorLevel = 2 Then
                                olaForm(0) = "Printer"
                                olaForm(1) = "Out of paper"
                            Else
                                Unload Me
                            End If
                        End If
                    Case OPOS_EPTR_REC_EMPTY
                        If .PowerState = 2004 Then
                            If .CoverOpen Then
                                olaForm(0) = "Printer"
                                olaForm(1) = "Printer cover is open"
                            ElseIf .ErrorLevel = 2 Then
                                olaForm(0) = "Printer"
                                olaForm(1) = "Out of paper"
                            Else
                                Unload Me
                            End If
                        ElseIf .oPrinterState = 4 Then
                            olaForm(0) = "Printer"
                            olaForm(1) = "Turn off"
                        Else
                            If .CoverOpen Then
                                olaForm(0) = "Printer"
                                olaForm(1) = "Printer cover is open"
                            ElseIf .ErrorLevel = 2 Then
                                olaForm(0) = "Printer"
                                olaForm(1) = "Out of paper"
                            Else
                                Unload Me
                            End If
                        End If
                End Select
            Case OPOS_SUCCESS
                If .PowerState = 2004 Then
                    If .CoverOpen Then
                        olaForm(0) = "Printer"
                        olaForm(1) = "Printer cover is open"
                    ElseIf .ErrorLevel = 2 Then
                        olaForm(0) = "Printer"
                        olaForm(1) = "Out of paper"
                    Else
                        olaForm(0) = "Printer"
                        olaForm(1) = "Turn off"
                    End If
                ElseIf .oPrinterState = 4 Then
                    olaForm(0) = "Printer"
                    olaForm(1) = "Turn off"
                Else
                    If .CoverOpen Then
                        olaForm(0) = "Printer"
                        olaForm(1) = "Printer cover is open"
                    ElseIf .ErrorLevel = 2 Then
                        olaForm(0) = "Printer"
                        olaForm(1) = "Out of paper"
                    Else
                        Unload Me
                    End If
                End If
        End Select
    End With
End Sub
