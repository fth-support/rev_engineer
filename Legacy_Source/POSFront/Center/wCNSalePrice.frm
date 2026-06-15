VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.1#0"; "Codejock.SuiteCtrls.Unicode.v11.1.3.ocx"
Begin VB.Form wCNSalePrice 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Prompt Sale;klg_Title"
   ClientHeight    =   3225
   ClientLeft      =   7530
   ClientTop       =   7815
   ClientWidth     =   5790
   ControlBox      =   0   'False
   Icon            =   "wCNSalePrice.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   5790
   Tag             =   "2"
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   2535
      Index           =   0
      Left            =   0
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   -240
      Width           =   5775
      _Version        =   720897
      _ExtentX        =   10186
      _ExtentY        =   4471
      _StockProps     =   64
      Appearance      =   10
      Color           =   32
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
         Left            =   2280
         MaxLength       =   11
         TabIndex        =   0
         Text            =   "0.00"
         Top             =   1680
         Width           =   3135
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   255
         Index           =   2
         Left            =   480
         TabIndex        =   10
         Top             =   2160
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label olaItemName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Item1"
         BeginProperty Font 
            Name            =   "AngsanaUPC"
            Size            =   18
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   2280
         TabIndex        =   9
         Tag             =   "SubTotal;ยอดคงเหลือ"
         Top             =   1080
         Width           =   570
      End
      Begin VB.Label olaBarC 
         BackStyle       =   0  'Transparent
         Caption         =   "0123456789"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   2280
         TabIndex        =   8
         Top             =   520
         Width           =   3135
      End
      Begin VB.Label olaForm 
         BackStyle       =   0  'Transparent
         Caption         =   "Sale Price :;klg_Saleprice"
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
         Index           =   2
         Left            =   360
         TabIndex        =   7
         Tag             =   "2;"
         Top             =   1740
         Width           =   2055
      End
      Begin VB.Label olaForm 
         BackStyle       =   0  'Transparent
         Caption         =   "Item Name :;klg_Itemname"
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
         TabIndex        =   6
         Tag             =   "2;"
         Top             =   1140
         Width           =   1695
      End
      Begin VB.Label olaForm 
         BackStyle       =   0  'Transparent
         Caption         =   "Barcode :;klg_Bar"
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
         TabIndex        =   5
         Tag             =   "2;"
         Top             =   540
         Width           =   1935
      End
   End
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   1305
      Index           =   1
      Left            =   0
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1920
      Width           =   5775
      _Version        =   720897
      _ExtentX        =   10186
      _ExtentY        =   2302
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
      Begin VB.CommandButton ocmOk 
         Caption         =   "&Ok(F9);klg_Ok"
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
         Left            =   2400
         TabIndex        =   1
         TabStop         =   0   'False
         Tag             =   "2;"
         Top             =   600
         Width           =   1455
      End
      Begin VB.CommandButton ocmCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel(Esc);klg_Cancel"
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
         Left            =   3960
         TabIndex        =   2
         TabStop         =   0   'False
         Tag             =   "2;"
         Top             =   600
         Width           =   1455
      End
   End
End
Attribute VB_Name = "wCNSalePrice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'หมายเหตุ : หน้าจอนี้สามารถระบุจำนวนเงินได้ทั้งหมด 10 degit เป็น ทศนิยม 2 degit  จำนวนเงินเต็ม 7 degit เครื่องจุด 1 degit
Option Explicit

Private Sub Form_Activate()
    olaBarC.Caption = tUT_SKUBarC
    'olaItemName.Caption = tUT_SKUName
    olaItemName.Caption = Replace(tUT_SKUName, "&", "&&") 'Bump 65-09-27 CommSheet TTK-ISS3000-811 TK ISSUE - The item name is displayed wrongly without the _ sign.xlsx
    'Replace
    Call SP_OBJxSetFocus(otbForm)
    otbForm.SelStart = 0
    otbForm.SelLength = Len(otbForm.Text)
End Sub

Private Sub Form_Load()
    Call SP_FRMxChildCenter(Me)
    Call SP_FRMxGetCaptionFrmDB2Tag(Me)
    Call SP_FRMxSetCapForm(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SP_OBJxSetNothing(Me)
End Sub

Private Sub ocmCancel_Click()
    bVB_PromptSale = False
   
    Unload Me
End Sub

Private Sub ocmOK_Click()
'*KT 53-05-29 แก้ไขการเกิด Dot send โดยการแก้ไขจากตัวตังต้น 3.0.19
On Error GoTo ErrHandle '*KT 67-01-04  runtime error  "stop statement encounter."
    Dim tFileMedia As String
    tFileMedia = "C:\WINDOWS\Media\Windows XP Critical Stop.wav"            '*Ao 53-03-15 File เสียง
'*Eaw 56-06-24 ไม่ต้อง Check banknote ตอนระบุราคา CommSheet TK-ISS3000-201-ST 00 00 04 FSBT NG_eaw.xlsx
'    If Not (SP_CHKbBanknoteAlw(SP_VALcText2Double(otbForm.Text))) Then
'        '*KT 53-04-23 กรณีที่การระบุจำนวนเงินไม่ผ่านให้ขึ้นเสียงเตือน ซึ่งเกิดจากปัญหา Communication Sheet  PH2-ML-SF-121 ให้มีเสียง Beepในกรณีที่ทำการ Scan สินค้า.xls
'        '*Ao 53-03-15
'        'คือให้ PC Speaker ส่งเสียงที่ความถี่ 3000Hz (3KHz) ด้วยระยะเวลา 0.3 วินาที (300mS)
'        If SP_FLEbManFile(tFileMedia, nEN_Fle1Exist) Then 'ตรวจสอบว่ามี file มีหรือไม่
'            Call PlaySound(tFileMedia, 0&, SND_ASYNC)
'        Else
'            Beep 5000, 400
'        End If
'        Call SP_OBJxSetFocus(otbForm)
'        Exit Sub      '*AO51-03-26   Comm.Sheet  ML-SF-100
'    End If
    'PH2-ML-SF- 121 สินค้า Prompt For Price (ใส่ราคาเอง) ให้ตรวจสอบโดยการห้ามใส่ราคาเป็น 0.00
    If Val(Format(otbForm.Text)) <= 0 Then
        Call SP_MSGnShowing("tms_CN005213", nCS_Warning)
        '*KT 53-04-23 กรณีที่การระบุจำนวนเงินไม่ผ่านให้ขึ้นเสียงเตือน ซึ่งเกิดจากปัญหา Communication Sheet  PH2-ML-SF-121 ให้มีเสียง Beepในกรณีที่ทำการ Scan สินค้า.xls
        '*Ao 53-03-15
        'คือให้ PC Speaker ส่งเสียงที่ความถี่ 3000Hz (3KHz) ด้วยระยะเวลา 0.3 วินาที (300mS)
        If SP_FLEbManFile(tFileMedia, nEN_Fle1Exist) Then 'ตรวจสอบว่ามี file มีหรือไม่
            Call PlaySound(tFileMedia, 0&, SND_ASYNC)
        Else
            Beep 5000, 400
        End If
        Call SP_OBJxSetFocus(otbForm)
        Exit Sub
    End If
    bVB_PromptSale = True
    cUT_RegPrice = SP_FMTcCurPoint(SP_FMTtNum2Shw(otbForm.Text), False, False)  'Regular Price
    cUT_SalePrice = cUT_RegPrice 'Sale Price ต้องหา promotion ก่อน
    bUT_PromptSale = True '*KT 52-05-04 แก้ไขกรณีที่ราคาสินค้าเป็น 0 แล้วขึ้น Prompt Sale ให้ระบุราคาถือว่าเป้นสินค้า Price Overide
    '*KT 52-03-31 PH1.8-ML-SF-013 กรณีที่สินค้าได้ราคา member price แต่เป็นสินค้า Prompt Sale นั้นต้องใส่ราคา Prompt Sale ก่อนแล้วคำนวณราคา Member Price ที่ได้
    '------------------------------------------------------------------------------------
    If bVB_CNMemberPrice Then
        If xVB_CNMemberP.ID <> "" And xVB_CNMemberP.MallCardMemberPrice = "Y" Then
            Call wFunctionKB.W_GETxMemberPrice(tUT_SKUCode, tVB_StaMemPriceAlw, cUT_MemberPrice, tVB_CNEventMemPrice, nVB_CNMemPriceMaxQty, tVB_CNMemPriceType, cVB_CNMemPriceAmt, cVB_CNMemPriceAmtDis) '*KT 52-09-07 Member Price + Consignment
        Else
            tVB_StaMemPriceAlw = "N"
            cUT_MemberPrice = 0
            tVB_CNEventMemPrice = ""
            nVB_CNMemPriceMaxQty = 0
            tVB_CNMemPriceType = ""
            cVB_CNMemPriceAmt = 0
        End If
    Else
       tVB_StaMemPriceAlw = "N"
       cUT_MemberPrice = 0
       tVB_CNEventMemPrice = ""
       nVB_CNMemPriceMaxQty = 0
        tVB_CNMemPriceType = ""
        cVB_CNMemPriceAmt = 0
    End If
    '-------------------------------------------------------------------------------------------
    Unload Me
    Exit Sub
ErrHandle:
     Exit Sub
End Sub

Private Sub otbForm_Change()            '*AO51-03-12
    otbForm.Text = SP_DEFtChangeNum(Trim(otbForm.Text))
    otbForm.SelStart = Len(otbForm.Text)
End Sub
Private Sub otbForm_GotFocus()
    Call SP_OBJxSetFocus(otbForm)
End Sub

Private Sub otbForm_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDelete         'delete
            otbForm.Text = ""
    End Select
End Sub

Private Sub otbForm_KeyPress(KeyAscii As Integer)
    Call UT_ChkKeyAscii(KeyAscii)
    Select Case KeyAscii
        Case 8, 6           'backspace, delete
            'do nothing
        Case 46             'ทศนิยม
            KeyAscii = 0
'            If InStr(Trim(otbForm.Text), ".") > 0 Then     '*AO51-03-12    Comm.Sheet SF- 070
'                KeyAscii = 0
'            End If
        Case 13
            Call ocmOK_Click
        Case Else
            If Not IsNumeric(Chr(KeyAscii)) Then
                KeyAscii = 0
            Else
            
            End If
    End Select
End Sub


