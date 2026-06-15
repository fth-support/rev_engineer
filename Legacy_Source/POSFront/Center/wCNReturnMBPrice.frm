VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.1#0"; "Codejock.SuiteCtrls.Unicode.v11.1.3.ocx"
Begin VB.Form wCNReturnMBPrice 
   Caption         =   "Member Price;klg_Title"
   ClientHeight    =   3120
   ClientLeft      =   6165
   ClientTop       =   3825
   ClientWidth     =   8910
   Icon            =   "wCNReturnMBPrice.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3120
   ScaleWidth      =   8910
   Tag             =   "2;"
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   2655
      Index           =   2
      Left            =   0
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   -240
      Width           =   8925
      _Version        =   720897
      _ExtentX        =   15743
      _ExtentY        =   4683
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
      Begin VB.TextBox otbMemberID 
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
         MaxLength       =   20
         TabIndex        =   0
         Top             =   405
         Width           =   3615
      End
      Begin VB.TextBox otbAmount 
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
         Left            =   6960
         TabIndex        =   1
         Top             =   405
         Width           =   1815
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Member Card;klg_MemberID"
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
         Left            =   120
         TabIndex        =   8
         Tag             =   "2;"
         Top             =   480
         Width           =   2100
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amt :;klg_Amt"
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
         Left            =   6000
         TabIndex        =   7
         Tag             =   "2;"
         Top             =   480
         Width           =   960
      End
   End
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   960
      Index           =   1
      Left            =   0
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2160
      Width           =   8925
      _Version        =   720897
      _ExtentX        =   15743
      _ExtentY        =   1693
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
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
         Left            =   3960
         TabIndex        =   2
         Tag             =   "2;"
         Top             =   360
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
         Left            =   7320
         TabIndex        =   4
         Tag             =   "2;"
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton ocmCancelAll 
         Caption         =   "&Cancel All;klg_CancelAll"
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
         Left            =   5640
         TabIndex        =   3
         Tag             =   "2;"
         Top             =   360
         Width           =   1455
      End
   End
End
Attribute VB_Name = "wCNReturnMBPrice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private bW_FoundMemberCard As Boolean
Private tW_IsChkDigitMallCard As String '1: allow check 0 : not allow check other : not allow check
Public nW_SeqNo As Long

Private Sub Form_Load()
    bW_FoundMemberCard = False
    Call SP_FRMxChildCenter(Me)
    Call SP_FRMxGetCaptionFrmDB2Tag(Me)
    Call SP_FRMxSetCapForm(Me)
    Call W_FRMxInitail
    If otbMemberID.Locked Then
        Call SP_OBJxSetFocus(otbAmount)
        Call SP_OTBxSelText(otbAmount)
    Else
        Call SP_OBJxSetFocus(otbMemberID)
    End If
End Sub

Private Sub W_FRMxInitail()
 '----------------------------------------------------
 'Ret:
 'Cmt:
 '----------------------------------------------------
    Call W_FRMxBlank
    Call W_FRMxSetMemberID
End Sub

Private Sub W_FRMxBlank()
  '----------------------------------------------------
 'Ret:
 'Cmt: set หน้า Member Card และ Amt = Blank
 '----------------------------------------------------
    otbMemberID.Text = ""
    otbMemberID.Locked = False
    otbMemberID.BackColor = nVB_CNClrBUnLock
    otbMemberID.TabStop = True
    otbAmount.Text = ""
    otbAmount.Locked = False
    otbAmount.BackColor = nVB_CNClrBUnLock
    otbAmount.TabStop = True
End Sub

Private Sub W_FRMxSetMemberID()
  '----------------------------------------------------
 'Ret:
 'Cmt: ตรวจสอบว่าก่นหน้ามีการให้ราการสินค้าอื่นที่เป็น Member Price หรือไม่ ถ้ามีต้องนำหมายเลข Member ID นั้น
 'เป็นค่า Default ในครั้งนี้
 '----------------------------------------------------
    Dim tSql As String
    Dim orsTemp As ADODB.Recordset
On Error GoTo HandleError
    tSql = "SELECT FNDctNo,FTTmnNum,FTShdTransNo,FDShdTransDate,FNScdSeqNo ,FCScdAmt"
    tSql = tSql & vbCrLf & "FROM " & tVB_CNTblCD
    tSql = tSql & vbCrLf & "WHERE  FTTmnNum='" & tVB_CNTerminalNum & "' AND FTShdTransNo='" & tVB_Running & "'"
    tSql = tSql & vbCrLf & "AND  FTShdTransType='" & tVB_Transaction & "' AND FDShdTransDate='" & tUT_SaleDate & "'"
    tSql = tSql & vbCrLf & "AND FNDctNo IN (" & nEN_Disc21MPrice & "," & nEN_Disc22MDisc & ")"
    tSql = tSql & vbCrLf & "AND FNSdtSeqNo = " & nW_SeqNo
    If SP_SQLvExecute(tSql, orsTemp) = 0 Then
        If Not (orsTemp.BOF And orsTemp.EOF) Then
            If bVB_CNMemberPrice And xVB_CNMemberP.ID <> "" Then
                otbMemberID.Text = xVB_CNMemberP.ID
                otbMemberID.Locked = True
                otbMemberID.BackColor = nVB_CNClrBLock
                otbMemberID.TabStop = False
                Call SP_OBJxSetFocus(otbAmount)
            End If
            otbAmount.Text = SP_FMTcCurPoint(SP_FEDvChkNumeric(orsTemp, "FCScdAmt"), False, True)
            cVB_CNDiscAmt = SP_VALcText2Double(otbAmount.Text)
        Else
            If bVB_CNMemberPrice And xVB_CNMemberP.ID <> "" Then
                otbMemberID.Text = xVB_CNMemberP.ID
                otbMemberID.Locked = True
                otbMemberID.BackColor = nVB_CNClrBLock
                otbMemberID.TabStop = False
                Call SP_OBJxSetFocus(otbAmount)
            End If
            otbAmount.Text = "0.00"
        End If
    Else
        If bVB_CNMemberPrice And xVB_CNMemberP.ID <> "" Then
            otbMemberID.Text = xVB_CNMemberP.ID
            otbMemberID.Locked = True
            otbMemberID.BackColor = nVB_CNClrBLock
            otbMemberID.TabStop = False
            Call SP_OBJxSetFocus(otbAmount)
        End If
        otbAmount.Text = "0.00"
    End If
    Call SP_DATxRsNothing(orsTemp)
    Exit Sub
HandleError:
    Call SP_DATxRsNothing(orsTemp)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SP_OBJxSetNothing(Me)
End Sub

Private Sub ocmCancel_Click()

    If otbMemberID.Locked = False Then
        bVB_CNMemberPrice = False
        xVB_CNMemberP.ID = ""
        xVB_CNMemberP.FName = ""
        xVB_CNMemberP.LName = ""
        xVB_CNMemberP.Tel = ""
        xVB_CNMemberP.Expire = Date
        xVB_CNMemberP.Status = ""
        xVB_CNMemberP.MallCardType = ""
        xVB_CNMemberP.MallCardExpire = Date
        xVB_CNMemberP.MallCardMemberPrice = "N"
        xVB_CNMemberP.MallCardMemberPoint = "N"
        xVB_CNMemberP.CardIDToken = ""
    End If
    bVB_CNReturnOk = False
    Unload Me
End Sub

Private Sub ocmCancelAll_Click()
    Call W_PRCxClaerAllCD
    bW_FoundMemberCard = False
'    bVB_CNMemberPrice = False
'    xVB_CNMemberP.ID = ""
'    xVB_CNMemberP.FName = ""
'    xVB_CNMemberP.LName = ""
'    xVB_CNMemberP.Tel = ""
'    xVB_CNMemberP.Expire = Date
'    xVB_CNMemberP.Status = ""
'    xVB_CNMemberP.MallCardType = ""
'    xVB_CNMemberP.MallCardExpire = Date
'    xVB_CNMemberP.MallCardMemberPrice = "N"
'    xVB_CNMemberP.MallCardMemberPoint = "N"
    Call W_FRMxInitail
    If otbMemberID.Locked Then
        Call SP_OBJxSetFocus(otbAmount)
        Call SP_OTBxSelText(otbAmount)
    Else
        Call SP_OBJxSetFocus(otbMemberID)
    End If
End Sub

Private Sub ocmOk_Click()
    Dim cAmt As Double
    cAmt = Val(Format(otbAmount.Text))
    If otbMemberID.Text = "" Then
        Call SP_MSGnShowing("tms_CN005214", nCS_Warning)
        Call SP_OBJxSetFocus(otbMemberID)
        Call SP_OTBxSelText(otbMemberID)
        Exit Sub
    End If
    If cAmt <= 0 Then
        Call SP_MSGnShowing("tms_CN005064", nCS_Warning)
        Call SP_OBJxSetFocus(otbAmount)
        Call SP_OTBxSelText(otbAmount)
        Exit Sub
    End If
    
    If bVB_CNMemberPrice And xVB_CNMemberP.ID <> "" Then
        bVB_CNMemberPrice = True
    Else
        If Not bW_FoundMemberCard Then
            xVB_CNMemberP.ID = Trim(otbMemberID.Text)
            If Not W_PRCxFindMemberCardByID(xVB_CNMemberP.ID) Then
                bW_FoundMemberCard = False
                otbMemberID.Text = ""
                otbAmount.Text = ""
                Call SP_OTBxSelText(otbMemberID)
                Exit Sub
            Else
                bW_FoundMemberCard = True
            End If
        Else
             xVB_CNMemberP.ID = Trim(otbMemberID.Text)
        End If
        If Not W_PRCxFindMallCardByIDMemberPrice(xVB_CNMemberP.ID) Then
                Call SP_MSGnShowing("tms_CN005200", nCS_Error)
                otbMemberID.Text = ""
                otbAmount.Text = ""
                Call SP_OTBxSelText(otbMemberID)
                Exit Sub
        Else
                '*KT 52-09-30 PH2-ML-SF-149 หมายเลขบัตร Mall Card ไม่ตรงกัน ระหว่าง M-Price กับ รูด Mall Card (Mall8)
                If tW_IsChkDigitMallCard = "1" Then
                    If Not SP_PRCbIsChkDigitMallCard(xVB_CNMemberP.ID) Then
                        Call SP_MSGnShowing("tms_CN005215", nCS_Error)
                        otbMemberID.Text = ""
                        otbAmount.Text = ""
                        Call SP_OTBxSelText(otbMemberID)
                        Exit Sub
                    End If
                End If
        End If
        bVB_CNMemberPrice = True
    End If
    cVB_CNDiscAmt = cAmt
    bVB_CNReturnOk = True
    Unload Me
End Sub

Private Function W_PRCxFindMemberCardByID(ByVal ptCusID$) As Boolean
'--------------------------------------
'Call:
'Cmt: ค้นหา Format ของเลขที่บัตร Mall Card อยู่ในช่วงใด
'          กรรีที่ไม่สามารถติดต่อฐานข้อมูล Member Server ได้ให้อนุญาตให้ทำ Function Member Price ได้ และกรณีที่หมายเลข member ที่รูดหรือระบุไม่มีในฐานข้อมูลให้ถือว่า
'         เป็น New Member สามารถทำ Function Member Price ได้เหมือนกัน
'*KT 52-02-12 New Requirment
'1.  เพิ่มหน้าจอ เพื่อรับบัตร Member กรณีที่ลูกค้าต้องการได้ราคา Member Price ทั้งนี้ ขึ้นอยู่กับ Card Type หาก Member Price ไม่ได้ระบุค่าจะเอาราคาตามปกติของสินค้าขณะนั้น มาแสดง
'*KT 52-05-26 PH 1.8 ML-SF-076 ในกรณีที่ POS ไม่สามารถติดต่อ Member Server ไม่ได้ จะขาย Member Price ไม่ได้
'--------------------------------------
    Dim tSql$
    Dim orsTemp As ADODB.Recordset
    Dim tCardNo As String
    Dim tTokenCardNo As String
    
On Error GoTo ErrHandle

    W_PRCxFindMemberCardByID = False
    '===================================
    '*Bmp 61-02-15 Token
    tTokenCardNo = ""
    If SP_PRCbToken("ReturnMBPrice", ptCusID, tTokenCardNo) Then
        tCardNo = tTokenCardNo
    Else
        tCardNo = ptCusID
    End If
    '===================================
    
    If ocnMember Is Nothing Then
        'Call SP_MSGnShowing("tms_CN005134", tCS_CNMsgErr)
        xVB_CNMemberP.FName = "New Member"
        xVB_CNMemberP.LName = ""
        xVB_CNMemberP.Tel = ""
        xVB_CNMemberP.Expire = Date
        xVB_CNMemberP.Status = ""
        xVB_CNMemberP.CardIDToken = tCardNo
        W_PRCxFindMemberCardByID = True
         Call SP_DATxRsNothing(orsTemp)
        Exit Function
    End If
    
    tSql = "SELECT TOP 1  FTMcdName,FTMcdSurName,FTMcdCode,FTMcdTel,FDMcdExpDate,FTMcdSta  FROM TCNMMallCard"
    tSql = tSql & " WHERE FTMcdCode='" & tCardNo & "'"
    Call SP_SQLvExecute(tSql, orsTemp, ocnMember, , , , True)
    If Not (orsTemp.BOF And orsTemp.EOF) Then
        If CDate(Format(SP_FEDdChkDate(orsTemp, "FDMcdExpDate"), "YYYY/MM/DD")) >= CDate(Format(SP_GETdChkDateDisc, "YYYY/MM/DD")) Then '*KT 52-06-04 ตรวจสอบวันที่หมดอายุตามที่กำหนดใน parameter
            Select Case SP_FEDtChkString(orsTemp, "FTMcdSta")
                Case "A", "1"
                    xVB_CNMemberP.FName = SP_FEDtChkString(orsTemp, "FTMcdName")
                    xVB_CNMemberP.LName = SP_FEDtChkString(orsTemp, "FTMcdSurName")
                    xVB_CNMemberP.Tel = SP_FEDtChkString(orsTemp, "FTMcdTel")
                    xVB_CNMemberP.Expire = SP_FEDdChkDate(orsTemp, "FDMcdExpDate")
                    xVB_CNMemberP.Status = SP_FEDtChkString(orsTemp, "FTMcdSta")
                    xVB_CNMemberP.CardIDToken = tCardNo
                    W_PRCxFindMemberCardByID = True
                Case "E", "2"
                    Call SP_MSGnShowing("tms_CN005201", tCS_CNMsgErr)
                    xVB_CNMemberP.FName = ""
                    xVB_CNMemberP.LName = ""
                    xVB_CNMemberP.Tel = ""
                    xVB_CNMemberP.Expire = Date
                    xVB_CNMemberP.Status = ""
                    xVB_CNMemberP.CardIDToken = ""
                    W_PRCxFindMemberCardByID = False
                Case "L", "3"
                    Call SP_MSGnShowing("tms_CN005202", tCS_CNMsgErr)
                    xVB_CNMemberP.FName = ""
                    xVB_CNMemberP.LName = ""
                    xVB_CNMemberP.Tel = ""
                    xVB_CNMemberP.Expire = Date
                    xVB_CNMemberP.Status = ""
                    xVB_CNMemberP.CardIDToken = ""
                    W_PRCxFindMemberCardByID = False
                Case "H", "4"
                    Call SP_MSGnShowing("tms_CN005203", tCS_CNMsgErr)
                    xVB_CNMemberP.FName = ""
                    xVB_CNMemberP.LName = ""
                    xVB_CNMemberP.Tel = ""
                    xVB_CNMemberP.Expire = Date
                    xVB_CNMemberP.Status = ""
                    xVB_CNMemberP.CardIDToken = ""
                    W_PRCxFindMemberCardByID = False
                Case Else 'นอกจากนี้อนุญาตให้ผ่าน
                    xVB_CNMemberP.FName = SP_FEDtChkString(orsTemp, "FTMcdName")
                    xVB_CNMemberP.LName = SP_FEDtChkString(orsTemp, "FTMcdSurName")
                    xVB_CNMemberP.Tel = SP_FEDtChkString(orsTemp, "FTMcdTel")
                    xVB_CNMemberP.Expire = SP_FEDdChkDate(orsTemp, "FDMcdExpDate")
                    xVB_CNMemberP.Status = SP_FEDtChkString(orsTemp, "FTMcdSta")
                    xVB_CNMemberP.CardIDToken = tCardNo
                    W_PRCxFindMemberCardByID = True
            End Select
        Else
            Call SP_MSGnShowing("tms_CN005201", tCS_CNMsgErr)
            xVB_CNMemberP.FName = ""
            xVB_CNMemberP.LName = ""
            xVB_CNMemberP.Tel = ""
            xVB_CNMemberP.Expire = Date
            xVB_CNMemberP.Status = ""
            xVB_CNMemberP.CardIDToken = ""
            W_PRCxFindMemberCardByID = False
        End If
    Else 'ไม่พบใน Member server ให้ถือว่าเป็นบัตรใหม่
'        Call SP_MSGnShowing("tms_CN005200", tCS_CNMsgErr)
        xVB_CNMemberP.FName = "New Member"
        xVB_CNMemberP.LName = ""
        xVB_CNMemberP.Tel = ""
        xVB_CNMemberP.Expire = Date
        xVB_CNMemberP.Status = ""
        xVB_CNMemberP.CardIDToken = tCardNo
        W_PRCxFindMemberCardByID = True
    End If
    Call SP_DATxRsNothing(orsTemp)
    Exit Function
ErrHandle:
        'Call SP_MSGnShowing("tms_CN005200", tCS_CNMsgErr)
        Call SP_DATxRsNothing(orsTemp)
        xVB_CNMemberP.FName = "New Member"
        xVB_CNMemberP.LName = ""
        xVB_CNMemberP.Tel = ""
        xVB_CNMemberP.Expire = Date
        xVB_CNMemberP.Status = ""
        xVB_CNMemberP.CardIDToken = ""
        W_PRCxFindMemberCardByID = True
End Function

Private Function W_PRCxFindMallCardByIDMemberPrice(ByVal ptCusID$) As Boolean
'--------------------------------------
'Call:
'Cmt: ค้นหา Format ของเลขที่บัตร Mall Card อยู่ในช่วงใด
'*KT 52-02-12 New Requirment
'1.  เพิ่มหน้าจอ เพื่อรับบัตร Member กรณีที่ลูกค้าต้องการได้ราคา Member Price ทั้งนี้ ขึ้นอยู่กับ Card Type หาก Member Price ไม่ได้ระบุค่าจะเอาราคาตามปกติของสินค้าขณะนั้น มาแสดง
'*KT 52-09-30 PH2-ML-SF-149 หมายเลขบัตร Mall Card ไม่ตรงกัน ระหว่าง M-Price กับ รูด Mall Card (Mall8) FTMcdIsChkDigit
'--------------------------------------
    Dim tSql$
    Dim orsTemp As ADODB.Recordset
On Error GoTo ErrHandle

    W_PRCxFindMallCardByIDMemberPrice = False
    tSql = "STP_PRCxFindMallCardByIDEx @ptAccNum= '" & ptCusID & "'"
    Call SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine, , , , , True, True) 'read master
    If Not (orsTemp.BOF Or orsTemp.EOF) Then
        xVB_CNMemberP.MallCardType = SP_FEDtChkString(orsTemp, "FTMcdCardType")
        xVB_CNMemberP.MallCardExpire = SP_FEDdChkDate(orsTemp, "FDMcdEndDate")
        xVB_CNMemberP.MallCardMemberPrice = SP_FEDtChkString(orsTemp, "FTMcdAlwPrice")
        xVB_CNMemberP.MallCardMemberPoint = SP_FEDtChkString(orsTemp, "FTMcdAlwPoint")
        W_PRCxFindMallCardByIDMemberPrice = True
        tW_IsChkDigitMallCard = SP_FEDtChkString(orsTemp, "FTMcdIsChkDigit")
    Else
        xVB_CNMemberP.MallCardType = ""
        xVB_CNMemberP.MallCardExpire = Date
        xVB_CNMemberP.MallCardMemberPrice = "N"
        xVB_CNMemberP.MallCardMemberPoint = "N"
        W_PRCxFindMallCardByIDMemberPrice = False
        tW_IsChkDigitMallCard = "0"
    End If
    Call SP_DATxRsNothing(orsTemp)
    Exit Function
ErrHandle:
    Call SP_DATxRsNothing(orsTemp)
    xVB_CNMemberP.MallCardType = ""
    xVB_CNMemberP.MallCardExpire = Date
    xVB_CNMemberP.MallCardMemberPrice = "N"
    xVB_CNMemberP.MallCardMemberPoint = "N"
     tW_IsChkDigitMallCard = "0"
    W_PRCxFindMallCardByIDMemberPrice = False
End Function

Private Sub W_PRCxClaerAllCD()
'---------------------------------------------
'Ret:
'Cmt: Delete all discount for Member Price on table CD
'--------------------------------------------
    Dim tSql As String
    tSql = "DELETE FROM  " & tVB_CNTblCD
    tSql = tSql & vbCrLf & " WHERE FTTmnNum = '" & tVB_CNTerminalNum & "'"
    tSql = tSql & vbCrLf & " AND FTShdTransNo = '" & tVB_Running & "'"
    tSql = tSql & vbCrLf & " AND FDShdTransDate = '" & tUT_SaleDate & "'"
    tSql = tSql & vbCrLf & " AND FNDctNo IN (" & nEN_Disc21MPrice & "," & nEN_Disc22MDisc & ")"
    Call SP_SQLvExecute(tSql)
End Sub



Private Sub otbAmount_Change()
        otbAmount.Text = SP_DEFtChangeNum(Trim(otbAmount.Text))
        otbAmount.SelStart = Len(otbAmount.Text)
End Sub

Private Sub otbAmount_GotFocus()
    Call SP_OTBxSelText(otbAmount)
End Sub

Private Sub otbAmount_KeyDown(KeyCode As Integer, Shift As Integer)
     Select Case KeyCode
        Case vbKeyDelete         'delete
            otbAmount.Text = ""
    End Select
End Sub

Private Sub otbAmount_KeyPress(KeyAscii As Integer)
    Call UT_ChkKeyAscii(KeyAscii)
    Select Case KeyAscii
        Case 8, 6           'backspace, delete
            'do nothing
        Case 46             'ทศนิยม
            KeyAscii = 0
        Case 13
            SendKeys "{TAB}"
'            Call ocmOk_Click
        Case Else
            If Not IsNumeric(Chr(KeyAscii)) Then
                KeyAscii = 0
            End If
    End Select
End Sub

Private Sub otbMemberID_GotFocus()
    Call SP_OTBxSelText(otbMemberID)
End Sub

Private Sub otbMemberID_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDelete         'delete
            If otbMemberID.Locked = False Then
                otbMemberID.Text = ""
            End If
    End Select
End Sub

Private Sub otbMemberID_KeyPress(KeyAscii As Integer)
    Call UT_ChkKeyAscii(KeyAscii)
    Select Case KeyAscii
        Case 8, 6           'backspace, delete
            'do nothing
        Case 46             'ทศนิยม
            KeyAscii = 0
        Case 13
            SendKeys "{TAB}"
'            Call ocmOk_Click
        Case Else
            If Not IsNumeric(Chr(KeyAscii)) Then
                KeyAscii = 0
            End If
    End Select
End Sub
