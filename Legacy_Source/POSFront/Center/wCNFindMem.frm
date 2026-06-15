VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.1#0"; "Codejock.SuiteCtrls.v11.1.3.ocx"
Begin VB.Form wCNFindMem 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Find Customer by Tel;klg_Title"
   ClientHeight    =   4575
   ClientLeft      =   5610
   ClientTop       =   4005
   ClientWidth     =   7155
   ControlBox      =   0   'False
   Icon            =   "wCNFindMem.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   7155
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   1110
      Index           =   2
      Left            =   0
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   3720
      Width           =   7140
      _Version        =   720897
      _ExtentX        =   12594
      _ExtentY        =   1958
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
      PaintManager.Position=   2
      Begin VB.CommandButton ocmOk 
         Caption         =   "&Ok;klg_Ok"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   3600
         TabIndex        =   1
         TabStop         =   0   'False
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
         Height          =   510
         Left            =   5280
         TabIndex        =   2
         TabStop         =   0   'False
         Tag             =   "2;"
         Top             =   120
         Width           =   1455
      End
   End
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   3960
      Index           =   0
      Left            =   0
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   -240
      Width           =   7140
      _Version        =   720897
      _ExtentX        =   12594
      _ExtentY        =   6985
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
      Begin VB.TextBox otbTel 
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
         Left            =   2520
         MaxLength       =   20
         TabIndex        =   0
         Top             =   720
         Width           =   4215
      End
      Begin VB.Label olbLPoint 
         BackColor       =   &H0080C0FF&
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
         Left            =   2520
         TabIndex        =   13
         Top             =   3120
         Width           =   4215
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Last Point:;klg_LPoint"
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
         Left            =   240
         TabIndex        =   12
         Tag             =   "2;"
         Top             =   3180
         Width           =   2670
      End
      Begin VB.Label olaCusID 
         BackColor       =   &H0080C0FF&
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
         Left            =   2520
         TabIndex        =   11
         Top             =   1320
         Width           =   4215
      End
      Begin VB.Label olaLName 
         BackColor       =   &H0080C0FF&
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
         Left            =   2520
         TabIndex        =   10
         Top             =   2520
         Width           =   4215
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Last Name:;klg_LName"
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
         Left            =   240
         TabIndex        =   9
         Tag             =   "2;"
         Top             =   2520
         Width           =   2910
      End
      Begin VB.Label olaFName 
         BackColor       =   &H0080C0FF&
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
         Left            =   2520
         TabIndex        =   8
         Top             =   1920
         Width           =   4215
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "First Name:;klg_FName"
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
         Left            =   240
         TabIndex        =   7
         Tag             =   "2;"
         Top             =   1920
         Width           =   2955
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tel:;klg_Tel"
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
         Left            =   240
         TabIndex        =   6
         Tag             =   "2;"
         Top             =   720
         Width           =   1470
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer ID:;klg_CusID"
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
         Left            =   240
         TabIndex        =   5
         Tag             =   "2;"
         Top             =   1320
         Width           =   3045
      End
   End
End
Attribute VB_Name = "wCNFindMem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub olaName_Click()
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'If KeyCode = 13 Then SendKeys "{Tab}"
    Select Case KeyCode
        Case vbKeyDelete         'delete
            otbTel.Text = ""
    End Select
End Sub

Private Sub Form_Load()
    Call SP_FRMxChildCenter(Me)
    Call SP_FRMxGetCaptionFrmDB2Tag(Me)
    Call SP_FRMxSetCapForm(Me)
    Call W_FRMxSetInitial
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SP_OBJxSetNothing(Me)
End Sub

Private Sub ocmCancel_Click()
    tVB_DiscMallCardIDByTel = ""
    Unload Me
End Sub

Private Sub ocmOK_Click()
    tVB_DiscMallCardIDByTel = Trim(olaCusID.Caption)
    Unload Me
End Sub

Private Sub W_FRMxSetInitial()
    olaCusID.BackColor = nVB_CNClrBLock
    olaFName.BackColor = nVB_CNClrBLock
    olaLName.BackColor = nVB_CNClrBLock
    olbLPoint.BackColor = nVB_CNClrBLock
End Sub

'*KT 53-01-26 กรณีเบอร์มากกว่า 1 คน ทุกครั้งที่เลือกจาก wCNSelectMember จะเข้า otbTel_GotFocus ทุกครั้งแทนที่ ocmOK setfocus
'Private Sub otbTel_GotFocus()
'    Call SP_OTBxSelText(otbTel)
'End Sub

Private Sub otbTel_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDelete         'delete
            otbTel.Text = ""
'        Case 13 '*KT 53-01-26 ไม่ใช้งานเพราะ otbTel_KeyPress ทำงานแล้ว และW_PRCxFindCusByTel มีการเรียก Call SP_OBJxSetFocus(ocmOk)
'            If Trim(otbTel.Text) <> "" Then
'                SendKeys "{Tab}"
'            End If
    End Select
End Sub

Private Sub otbTel_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 13
            If Trim(otbTel.Text) <> "" Then
                Call W_PRCxFindCusByTel(Trim(otbTel.Text))
            End If
            If olaCusID.Caption = "" Then
                Call SP_OBJxSetFocus(otbTel)
            End If
    End Select
End Sub

Private Sub W_PRCxFindCusByTel(ByVal ptTel As String)
'---------------------------------------------------
'Call:
'Cmt: ต้องการค้นหาเลขที่ Mall Card จากหมายเลขโทรศัพท์ แล้วแสดงข้อมูล
'----------------------------------------------------
    Dim tSql$
    Dim orsTemp As ADODB.Recordset
On Error GoTo ErrHandle
    Call SP_STAxMouse(Me, True)
    If ocnMember Is Nothing Then
        Call SP_MSGnShowing("tms_CN005134", tCS_CNMsgErr)
        olaCusID.Caption = ""
        olaFName.Caption = ""
        olaLName.Caption = ""
        olbLPoint.Caption = ""
         Call SP_DATxRsNothing(orsTemp)
         Call SP_STAxMouse(Me, False)
        Exit Sub
    End If
    tSql = "SELECT Top 1000  FTMcdName,FTMcdSurName,FTMcdCode,FDMcdExpDate,FTMcdSta,FCMcdTotalPoint"
    tSql = tSql & vbCrLf & "FROM TCNMMallCard"
    tSql = tSql & vbCrLf & "WHERE (FTMcdTel = '" & ptTel & "')"
    tSql = tSql & vbCrLf & "OR (FTMcdMobile = '" & ptTel & "')"
    tSql = tSql & vbCrLf & "OR (FTMcdOfficePhone = '" & ptTel & "')"
    If SP_SQLvExecute(tSql, orsTemp, ocnMember, , , , True) = 0 Then
        If Not (orsTemp.BOF Or orsTemp.EOF) Then
            If orsTemp.RecordCount > 1 Then
                Call SP_STAxMouse(Me, False)
                Set wCNSelectMember.orsMember = orsTemp.Clone
                wCNSelectMember.Show vbModal
'                DoEvents        '*KT 53-01-14
                If tVB_FindMember <> "" Then
                    orsTemp.Filter = "FTMcdCode = '" & tVB_FindMember & "'"
                     If Not (orsTemp.BOF And orsTemp.EOF) Then
                        olaCusID.Caption = SP_FEDtChkString(orsTemp, "FTMcdCode")
                        olaFName.Caption = SP_FEDtChkString(orsTemp, "FTMcdName")
                        olaLName.Caption = SP_FEDtChkString(orsTemp, "FTMcdSurName")
                        olbLPoint.Caption = SP_FEDvChkNumeric(orsTemp, "FCMcdTotalPoint")
                        Call SP_OBJxSetFocus(ocmOk)
                    Else
                        olaCusID.Caption = ""
                        olaFName.Caption = ""
                        olaLName.Caption = ""
                        olbLPoint.Caption = ""
                         Call SP_OBJxSetFocus(otbTel)
                         Call SP_OTBxSelText(otbTel)
                    End If
                Else
                    olaCusID.Caption = ""
                    olaFName.Caption = ""
                    olaLName.Caption = ""
                    olbLPoint.Caption = ""
                     Call SP_OBJxSetFocus(otbTel)
                     Call SP_OTBxSelText(otbTel)
                End If
            Else
                olaCusID.Caption = SP_FEDtChkString(orsTemp, "FTMcdCode")
                olaFName.Caption = SP_FEDtChkString(orsTemp, "FTMcdName")
                olaLName.Caption = SP_FEDtChkString(orsTemp, "FTMcdSurName")
                olbLPoint.Caption = SP_FEDvChkNumeric(orsTemp, "FCMcdTotalPoint")
                Call SP_OBJxSetFocus(ocmOk)
            End If
        Else
            olaCusID.Caption = ""
            olaFName.Caption = ""
            olaLName.Caption = ""
            olbLPoint.Caption = ""
            Call SP_MSGnShowing("tms_CN005108", tCS_CNMsgErr)
            Call SP_OBJxSetFocus(otbTel)
            Call SP_OTBxSelText(otbTel)
        End If
    Else
        olaCusID.Caption = ""
        olaFName.Caption = ""
        olaLName.Caption = ""
        olbLPoint.Caption = ""
        Call SP_MSGnShowing("tms_CN005108", tCS_CNMsgErr)
        Call SP_OBJxSetFocus(otbTel)
        Call SP_OTBxSelText(otbTel)
    End If
    Call SP_DATxRsNothing(orsTemp)
    Call SP_STAxMouse(Me, False)
    Exit Sub
ErrHandle:
    Call SP_MSGxShowError(Err.Number, Err.Description)
    Call SP_DATxRsNothing(orsTemp)
    Call SP_STAxMouse(Me, False)
End Sub
