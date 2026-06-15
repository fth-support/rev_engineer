VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPrinter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'*KT 53-05-29 แก้ไขการเกิด Dot send โดยการแก้ไขจากตัวตังต้น 3.0.19

Private tW_TransTypeSub As String

Private nW_State As Boolean

Private oPrn As Printer

Private cW_BaseLineOffset As Double

Private cW_TextHeight As Double

Private cW_CurY As Double

Private cW_LineOffset As Double

Private uW_BMPInfo As BITMAP

Private bW_LineFeed As Boolean

Private bW_PrintEJ As Boolean

Private bW_PRMGenEJ As Boolean '*KT 51-12-16 Flag ที่บอกว่ามีการเรียกใช้งานจาก Tools Generate EJ

Private aW_EJ() As String
Private cW_PWidth As Double
Private nW_PColumn As Long

Public Enum eEN_PrintTo
        nEN_PrinterAndEJ = 0
        nEN_PrinterOnly = 1
        eEN_EJOnly = 2
End Enum

Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type


Private Type TEXTMETRIC
        tmHeight As Long
        tmAscent As Long
        tmDescent As Long
        tmInternalLeading As Long
        tmExternalLeading As Long
        tmAveCharWidth As Long
        tmMaxCharWidth As Long
        tmWeight As Long
        tmOverhang As Long
        tmDigitizedAspectX As Long
        tmDigitizedAspectY As Long
        tmFirstChar As Byte
        tmLastChar As Byte
        tmDefaultChar As Byte
        tmBreakChar As Byte
        tmItalic As Byte
        tmUnderlined As Byte
        tmStruckOut As Byte
        tmPitchAndFamily As Byte
        tmCharSet As Byte
End Type

Private bC_PrnByOPOS As Boolean
Private tC_LastTextPrint As String
Private bC_LastFontBold  As Boolean
Private bC_LastFontHeight  As Boolean


Private tC_TmnNum As String
Private tC_PthImageTmp As String
Private tC_PthImageCpnTmp As String

Private tC_PrinterType As String '0:Slip Printer 1:Receipt Printer

Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function GetTextMetrics Lib "gdi32" Alias "GetTextMetricsA" (ByVal hDC As Long, lpMetrics As TEXTMETRIC) As Long

Private nC_LineNo As Long
Private nC_LineNoEx As Long
Private nC_PageNo As Long
Private oW_Logo As IPictureDisp

Private nC_LineMax As Long

Private Type LineHeader
    tText As String
    cPosition As Double
    nAli As Long
    bLF As Boolean
End Type
Private oC_LineHeader(3) As LineHeader
Private bC_LineHeader As Boolean
Private nLevel As Long
'events
'Public Event OnPrintText(ByVal ptText As String, ByVal pcX As String, pbLF As Boolean)

Public Function SelectPrinter(ByVal poName As String) As Printer
Dim oPrinter As Printer
On Error GoTo ErrHandle
        If poName = vbNullString Then
                Set SelectPrinter = Nothing
                nW_State = 0
                Exit Function
        End If

        For Each oPrinter In Printers
                If oPrinter.DeviceName = poName Then
                        Set Printer = oPrinter
                        Set SelectPrinter = oPrinter
                        nW_State = 1
                        Exit Function
                End If
        Next
        MsgBox "Printer device not found. ", vbCritical 'tms_CN005219
        Exit Function
'Err.Raise 999, , "Printer device not found. "
ErrHandle:
     MsgBox "Printer device error.", vbCritical 'tms_CN005220
End Function

Property Let tTransTypeSub(New_tTransTypeSub As String) '*Ao 2013/08/31  Trans Type
    tW_TransTypeSub = New_tTransTypeSub
End Property

 '*KT 51-12-16 Flag ที่บอกว่ามีการเรียกใช้งานจาก Tools Generate EJ
Property Let bPRMGenEJ(New_bPRMGenEJ As Boolean)
    bW_PRMGenEJ = New_bPRMGenEJ
End Property

 '*KT 51-12-16 Flag ที่บอกว่ามีการเรียกใช้งานจาก Tools Generate EJ
Property Get bPRMGenEJ() As Boolean
    bPRMGenEJ = bW_PRMGenEJ
End Property

Property Get bState() As Boolean
    bState = nW_State
End Property

Property Let tTmnNum(New_tTmnNum As String)
    tC_TmnNum = New_tTmnNum
End Property

Property Get nTmnNum() As String
    nTmnNum = tC_TmnNum
End Property

Property Let bPrnByOPOS(New_bPrnByOPOS As Boolean)
    bC_PrnByOPOS = New_bPrnByOPOS
End Property

Property Get bPrnByOPOS() As Boolean
    bPrnByOPOS = bC_PrnByOPOS
End Property

Property Get tLastTextPrint() As String
    tLastTextPrint = tC_LastTextPrint
End Property

Public Property Let CurrentX(ByVal pnX As Double)
        oPrn.CurrentX = pnX
End Property

Public Property Get CurrentX() As Double
        CurrentX = oPrn.CurrentX
End Property

Public Property Let CurrentY(ByVal pnY As Double)
        oPrn.CurrentY = pnY
        cW_CurY = pnY
End Property

Public Property Get CurrentY() As Double
        CurrentY = oPrn.CurrentY
End Property

Public Property Let PrintEJ(ByVal pnPrint As Boolean)
        bW_PrintEJ = pnPrint
End Property

Public Property Get PrintEJ() As Boolean
        PrintEJ = bW_PrintEJ
End Property


Public Property Set Font(poFont As IFontDisp)
        Set oPrn.Font = poFont
End Property

Public Property Let LineOffset(ByVal pcOffset As Double)

        cW_LineOffset = pcOffset
End Property

Public Property Get LineOffset() As Double
        cW_LineOffset = oPrn.CurrentX
End Property

Public Property Get EJText() As String
Dim nTmp As Long
        
        For nTmp = 0 To UBound(aW_EJ)
            EJText = EJText & aW_EJ(nTmp) & vbCrLf
        Next
        
End Property

Public Function EJNewCapture(ByVal pcPageWidth As Double, ByVal pnPageColumn As Long)

        ReDim aW_EJ(0) As String
        cW_PWidth = pcPageWidth
        nW_PColumn = pnPageColumn
        
End Function

Public Function PrintPictureFromFile(ByVal ptFile As String _
, Optional ByVal pcX As Double _
, Optional ByVal pcY As Double = -1 _
, Optional ByVal pnAlign As VBRUN.AlignmentConstants = vbLeftJustify _
) As Boolean

On Error GoTo ErrHandler
Dim oPic As StdPicture
Dim nXcm As Double, nYcm As Double
Dim nWcm As Double, nHcm As Double

        Set oPic = LoadPicture(ptFile)
        GetObjectAPI oPic.handle, Len(uW_BMPInfo), uW_BMPInfo
       With uW_BMPInfo
       
            nWcm = (.bmWidth / oPrn.PrintQuality) * 2.54
            nHcm = (.bmHeight / oPrn.PrintQuality) * 2.54

       End With
       
        'set Y - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
                pcY = IIf(pcY = -1, oPrn.CurrentY, pcY)
                nYcm = pcY
                
        'set X - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

               pcX = IIf(pcX = -1, oPrn.CurrentX, pcX)
               
                Select Case pnAlign
                Case vbLeftJustify
                        nXcm = pcX
                Case vbCenter
                        nXcm = pcX - (nWcm / 2)
                Case vbRightJustify
                        nXcm = pcX - (nWcm)
                End Select
                
        'print - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
                Call oPrn.PaintPicture(oPic, nXcm, nYcm, nWcm, nHcm)
                
                cW_CurY = nYcm + nHcm + 0.4  'currentY
                
'return
PrintPictureFromFile = True

Exit Function '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
ErrHandler:

End Function

'Public Function PrintPicture(ByVal poPicture As IPictureDisp _
', Optional ByVal pcX As Double _
', Optional ByVal pcY As Double = -1 _
', Optional ByVal pnAlign As VBRUN.AlignmentConstants = vbLeftJustify _
') As Boolean
'
'On Error GoTo ErrHandler
'Dim oPic As StdPicture
'Dim nXcm As Double, nYcm As Double
'Dim nWcm As Double, nHcm As Double
'
'        Set oPic = poPicture
'        GetObjectAPI oPic.handle, Len(uW_BMPInfo), uW_BMPInfo
'       With uW_BMPInfo
'
'            nWcm = (.bmWidth / oPrn.PrintQuality) * 2.54
'            nHcm = (.bmHeight / oPrn.PrintQuality) * 2.54
'
'       End With
'
'        'set Y - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'                pcY = IIf(pcY = -1, oPrn.CurrentY, pcY)
'                nYcm = pcY
'
'        'set X - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'
'               pcX = IIf(pcX = -1, oPrn.CurrentX, pcX)
'
'                Select Case pnAlign
'                Case vbLeftJustify
'                        nXcm = pcX
'                Case vbCenter
'                        nXcm = pcX - (nWcm / 2)
'                Case vbRightJustify
'                        nXcm = pcX - (nWcm)
'                End Select
'
'        'print - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'                Call oPrn.PaintPicture(oPic, nXcm, nYcm, nWcm, nHcm)
'
'                cW_CurY = nYcm + nHcm + 0.4  'currentY
'
''return
'PrintPicture = True
'
'Exit Function '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'ErrHandler:
'
'End Function

Public Function PrintPicture(ByVal poPicture As IPictureDisp _
, Optional ByVal pcX As Double _
, Optional ByVal pcY As Double = -1 _
, Optional ByVal pnAlign As VBRUN.AlignmentConstants = vbLeftJustify _
, Optional ByVal pcW As Double _
, Optional ByVal pcH As Double _
) As Boolean

On Error GoTo ErrHandler
Dim oPic As StdPicture
Dim nXcm As Double, nYcm As Double
Dim nWcm As Double, nHcm As Double
    If Not bC_PrnByOPOS Then
'driver
        Set oPic = poPicture
        GetObjectAPI oPic.handle, Len(uW_BMPInfo), uW_BMPInfo

       With uW_BMPInfo
            nWcm = (.bmWidth / oPrn.PrintQuality) * 2.54
            nHcm = (.bmHeight / oPrn.PrintQuality) * 2.54
       End With
       
       If pcW > 0 And pcH > 0 Then
              nWcm = pcW
              nHcm = pcH
       End If

        'set Y - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
                pcY = IIf(pcY = -1, oPrn.CurrentY, pcY)
                nYcm = pcY
        'set X - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

               pcX = IIf(pcX = -1, oPrn.CurrentX, pcX)
                Select Case pnAlign
                Case vbLeftJustify
                        nXcm = pcX
                Case vbCenter
                        nXcm = pcX - (nWcm / 2)
                Case vbRightJustify
                        nXcm = pcX - (nWcm)
                End Select

        'print - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

                Call oPrn.PaintPicture(oPic, nXcm, nYcm, nWcm, nHcm)
                cW_CurY = nYcm + nHcm + 0.4  'currentY
    Else
'oopos
        tC_PthImageTmp = SP_GETtPthLoadImageFromField("TCNMReceiptJN a,TCNMTerminalMtn b " & vbCrLf, "FORjfLogo", "a.FTRjfCode =b.FTRjfCode AND b.FTTmnNum ='" & tC_TmnNum & "'")
        If tC_PthImageTmp <> "" Then
            Set oPic = poPicture
            GetObjectAPI oPic.handle, Len(uW_BMPInfo), uW_BMPInfo
'Slip Printer no logo
'            oOposSlip.RecLetterQuality = True
'            oOposSlip.MapMode = PCRW_MM_DOTS
'            oOposSlip.PrintBitmap PTR_S_RECEIPT, tC_PthImageTmp, uW_BMPInfo.bmWidth, DISP_BM_CENTER
'            oOposSlip.RecLetterQuality = False

            oOpos.RecLetterQuality = True
            oOpos.MapMode = PCRW_MM_DOTS
            oOpos.PrintBitmap PTR_S_RECEIPT, tC_PthImageTmp, uW_BMPInfo.bmWidth, DISP_BM_CENTER
            oOpos.RecLetterQuality = False

        Else
            MsgBox "Can not set image logo", vbCritical    '*Ao 09-07-23 'tms_CN005221
        End If
    End If
'return
PrintPicture = True

Exit Function '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
ErrHandler:

End Function

'Public Function PrintPicture20111110(ByVal poPicture As IPictureDisp _
', Optional ByVal pcX As Double _
', Optional ByVal pcY As Double = -1 _
', Optional ByVal pnAlign As VBRUN.AlignmentConstants = vbLeftJustify _
', Optional ByVal pcW As Double _
', Optional ByVal pcH As Double _
') As Boolean
'
'On Error GoTo ErrHandler
'Dim oPic As StdPicture
'Dim nXcm As Double, nYcm As Double
'Dim nWcm As Double, nHcm As Double
'    If Not bC_PrnByOPOS Then
''driver
'        Set oPic = poPicture
'        GetObjectAPI oPic.handle, Len(uW_BMPInfo), uW_BMPInfo
'
'       With uW_BMPInfo
'            nWcm = (.bmWidth / oPrn.PrintQuality) * 2.54
'            nHcm = (.bmHeight / oPrn.PrintQuality) * 2.54
'       End With
'
'       If pcW > 0 And pcH > 0 Then
'              nWcm = pcW
'              nHcm = pcH
'       End If
'
'        'set Y - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'                pcY = IIf(pcY = -1, oPrn.CurrentY, pcY)
'                nYcm = pcY
'        'set X - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'
'               pcX = IIf(pcX = -1, oPrn.CurrentX, pcX)
'                Select Case pnAlign
'                Case vbLeftJustify
'                        nXcm = pcX
'                Case vbCenter
'                        nXcm = pcX - (nWcm / 2)
'                Case vbRightJustify
'                        nXcm = pcX - (nWcm)
'                End Select
'
'        'print - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'
'                Call oPrn.PaintPicture(oPic, nXcm, nYcm, nWcm, nHcm)
'                cW_CurY = nYcm + nHcm + 0.4  'currentY
'    End If
''return
'PrintPicture2 = True
'
'Exit Function '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'ErrHandler:
'
'End Function

'Public Sub SetBitmap(ByVal poPicture As IPictureDisp)
''-----------------------------------------------------------
''Call: *KT 51-09-2008 SetBitmap
''Cmt:
''------------------------------------------------------------
'On Error GoTo ErrHandler
'    Dim oPic As StdPicture
'    bVB_ImageLogo = False
'    tC_PthImageTmp = SP_GETtPthLoadImageFromField("TCNMReceiptJN a,TCNMTerminalMtn b " & vbCrLf, "FORjfLogo", "a.FTRjfCode =b.FTRjfCode AND b.FTTmnNum ='" & tC_TmnNum & "'")
'    If tC_PthImageTmp <> "" Then
'        Set oPic = poPicture
'        GetObjectAPI oPic.handle, Len(uW_BMPInfo), uW_BMPInfo
''Slip Printer no logo
''        oOposSlip.RecLetterQuality = True
''        oOposSlip.SetBitmap 1, PTR_S_RECEIPT, tC_PthImageTmp, 384, DISP_BM_CENTER
''        oOposSlip.RecLetterQuality = False
'
'        oOpos.RecLetterQuality = True
'        oOpos.SetBitmap 1, PTR_S_RECEIPT, tC_PthImageTmp, 384, DISP_BM_CENTER
'        oOpos.RecLetterQuality = False
'
'        Call Kill(tC_PthImageTmp)
'         bVB_ImageLogo = True
'
'    Else
'            MsgBox "Can not set image logo", vbCritical    '*Ao 09-07-23 'tms_CN005221
'    End If
'    Exit Sub
'ErrHandler:
'    bVB_ImageLogo = False
'End Sub
Public Sub SetBitmap(ByVal poPicture As IPictureDisp)
'-----------------------------------------------------------
'Call: *KT 51-09-2008 SetBitmap
'Cmt:
'------------------------------------------------------------
Dim oFSO As New FileSystemObject
On Error GoTo ErrHandler
    Dim oPic As StdPicture
    bVB_ImageLogo = False       '*Ao 2013/01/24   Print FRONTTECH
    tC_PthImageTmp = SP_GETtPthLoadImageFromField("TCNMReceiptJN a,TCNMTerminalMtn b " & vbCrLf, "FORjfLogo", "a.FTRjfCode =b.FTRjfCode AND b.FTTmnNum ='" & tC_TmnNum & "'")
    If tC_PthImageTmp <> "" Then
'        Set oPic = poPicture
'        GetObjectAPI oPic.handle, Len(uW_BMPInfo), uW_BMPInfo
'        Select Case tVB_PrnDevice
'            Case "FP360SERPRT", "FP360PARPRT", "FP360USBPRT"
'                oOpos.SetBitmap 1, PTR_S_RECEIPT, tC_PthImageTmp, 384, DISP_BM_CENTER               '*Ao 2012/12/25
'            Case Else
'                oOpos.RecLetterQuality = True
'                oOpos.SetBitmap 1, PTR_S_RECEIPT, tC_PthImageTmp, 384, DISP_BM_CENTER
'                oOpos.RecLetterQuality = False
'        End Select
''        oOpos.RecLetterQuality = True
''        oOpos.SetBitmap 1, PTR_S_RECEIPT, tC_PthImageTmp, 384, DISP_BM_CENTER
''        oOpos.RecLetterQuality = False
'        Call Kill(tC_PthImageTmp)
'         bVB_ImageLogo = True
'        MsgBox tC_PthImageTmp
'        tC_PthImageTmp = "C:\Program Files\ISS3000\POS\Taka_LogoFP410.bmp"
        
'        384   ความกว้างของ Line_Width
'        oOpos.SetBitmap 1, PTR_S_RECEIPT, tC_PthImageTmp, 384, DISP_BM_CENTER          '35
'        oOpos.SetBitmap 1, PTR_S_RECEIPT, tC_PthImageTmp, 576, DISP_BM_CENTER          '48
'        oOpos.SetBitmap 1, PTR_S_RECEIPT, tC_PthImageTmp, 512, DISP_BM_CENTER          '48
        oOpos.SetBitmap 1, PTR_S_RECEIPT, tC_PthImageTmp, 512, DISP_BM_CENTER           '42
        'oOpos.RecLineWidth
        bVB_ImageLogo = True
        Call SP_PRNxGetLogo(tC_TmnNum, tC_PthImageTmp)                '*Ao 2013/05/10
        Call Kill(tC_PthImageTmp)
    Else
        MsgBox "Can not set image logo", vbCritical
    End If
    Exit Sub
ErrHandler:
    MsgBox "ErrHandler:Can not set image logo", vbCritical
    bVB_ImageLogo = False
End Sub

Public Sub SetBitmapCpn(ByVal poPicture As IPictureDisp, ByVal ptID As String)
'-----------------------------------------------------------
'Call: *KT 54-07-27 SetBitmap of promotion coupon
'Cmt:
'------------------------------------------------------------
On Error GoTo ErrHandler
    Dim oPic As StdPicture
    bVB_ImageLogoCpn = False
    tC_PthImageCpnTmp = SP_GETtPthLoadImageFromField("TCNMPmtCpnJN", "FOPcnLogo", "FTPcnID='" & ptID & "'")
    If tC_PthImageCpnTmp <> "" Then
        Set oPic = poPicture
        GetObjectAPI oPic.handle, Len(uW_BMPInfo), uW_BMPInfo
'Slip Printer no logo
'        oOposSlip.RecLetterQuality = True
'        oOposSlip.SetBitmap 2, PTR_S_RECEIPT, tC_PthImageCpnTmp, 384, DISP_BM_CENTER
'        oOposSlip.RecLetterQuality = False

        oOpos.RecLetterQuality = True
        oOpos.SetBitmap 2, PTR_S_RECEIPT, tC_PthImageCpnTmp, 384, DISP_BM_CENTER        '*Ao 2013/01/11  ต้องใช้ oOposSlip หรอไม่ ?
        oOpos.RecLetterQuality = False

        Call Kill(tC_PthImageCpnTmp)
         bVB_ImageLogoCpn = True
    Else
            MsgBox "Can not set image logo", vbCritical    '*Ao 09-07-23 'tms_CN005221
    End If
    Exit Sub
ErrHandler:
    bVB_ImageLogoCpn = False
End Sub

Public Sub PrintLogo()
'-----------------------------------------------------------
'Call: *KT 51-09-2008 SetBitmap
'Cmt:
'------------------------------------------------------------
Dim ESC As String
Dim LF As String
On Error GoTo ErrHandler
    ESC = Chr(27)
    LF = Chr(10)
    
    oOpos.Printnormal PTR_S_RECEIPT, ESC & "|1B" & LF

'Slip Printer no logo
    'oOposSlip.Printnormal PTR_S_RECEIPT, ESC & "|1B" & LF
    Exit Sub
ErrHandler:
    
End Sub

Public Sub PrintLogoCpn()
'-----------------------------------------------------------
'Call: SetBitmap for coupon
'Cmt: *KT 54-07-26 V.4.0.1 Phase4.0(Member Point) Promotion Coupon
'------------------------------------------------------------
Dim ESC As String
Dim LF As String
On Error GoTo ErrHandler
    ESC = Chr(27)
    LF = Chr(10)
    
    oOpos.Printnormal PTR_S_RECEIPT, ESC & "|2B" & LF
'Slip Printer no logo2
    'oOposSlip.Printnormal PTR_S_RECEIPT, ESC & "|2B" & LF
    Exit Sub
ErrHandler:
    
End Sub

Public Function PrintText(ByVal ptText As String _
, Optional ByVal pcX As Double = -1 _
, Optional ByVal pcY As Double = -1 _
, Optional ByVal pnAlign As VBRUN.AlignmentConstants = vbLeftJustify _
, Optional ByVal pbLF As Boolean = True _
, Optional ByVal pnPrintTo As eEN_PrintTo = nEN_PrinterAndEJ _
, Optional ByVal pcWidthLimit As Double _
, Optional ByVal pbCR As Boolean = False _
, Optional ByVal pbBold As Boolean = False _
, Optional ByVal pbHeight As Boolean = False _
) As Boolean

Dim nTmp As Long, nCol As Long
Dim cTextWidth As Double
Dim nXCol As Long, nTextDigit As Long
Dim nSpace As Long
Dim nSpaceBlank As Long
Dim nLenTha As Long
Dim nLenTextPrint As Long

'ESc Pos
Dim SEND As String
Dim ESC As String
Dim GS As String
Dim LF As String
Dim CR As String
Dim a As String

Dim nSpaceBlankOld As Long
Dim nLen As Long
Dim nLenTextFull  As Long
Dim tSlipCheck As String '*Bump 57-09-30
Dim nI As Long

ESC = Chr(27)
GS = Chr(29)
LF = Chr(10)
CR = Chr(13)
a = Chr(97)


On Error GoTo ErrHandler

'Initail Driver  for EJ and Print
    Call W_GETnBaseLineOffSet
    With oPrn
            'WidthLimit - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
            If pcWidthLimit > 0 Then
                    ptText = W_GETtLTrim(ptText, pcWidthLimit)
            End If
                    
            'set Y - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
                
            If pcY = -1 Then
                    If bW_LineFeed Then
                        cW_CurY = cW_CurY + ((cW_TextHeight) + cW_LineOffset)  'next line
                    End If
            Else
                    cW_CurY = pcY
            End If
            
            .CurrentY = cW_CurY - cW_BaseLineOffset
            bW_LineFeed = pbLF
                    
            
            'set X - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
            cTextWidth = .TextWidth(ptText)
            If pcX <> -1 Then
                    Select Case pnAlign
                    Case vbLeftJustify
                            .CurrentX = pcX
                    Case vbCenter
                            .CurrentX = pcX - (cTextWidth / 2)
                    Case vbRightJustify
                            .CurrentX = pcX - (cTextWidth)
                    End Select
            End If
    
    End With

    'EJ - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
    If bW_PrintEJ And (pnPrintTo = eEN_EJOnly Or pnPrintTo = nEN_PrinterAndEJ) Then
            nTmp = UBound(aW_EJ())
            If aW_EJ(nTmp) = vbNullString Then
                    aW_EJ(nTmp) = Space(nW_PColumn)
            End If
            
            nTextDigit = W_GETnTextDigit(ptText)
            'nTextDigit = Len(ptText)
            nXCol = (nW_PColumn * pcX) \ cW_PWidth
            
            Select Case pnAlign
            Case vbLeftJustify
                    nXCol = nXCol
            Case vbCenter
                    nXCol = nXCol - (nTextDigit / 2)
            Case vbRightJustify
                    nXCol = nXCol - (nTextDigit)
            End Select
            
            If Len(aW_EJ(nTmp)) >= nXCol + 1 Then
                Mid(aW_EJ(nTmp), nXCol + 1) = ptText
            End If
            
            If bW_LineFeed Then ReDim Preserve aW_EJ(nTmp + 1) As String  'for next line
    End If
    
If Not bW_PRMGenEJ Then '*KT 51-12-16 Flag แสดงว่ามีการเรียกใช้จาก Tools Generate EJ
   If Not bC_PrnByOPOS Then
'Driver
            'print - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
            '*KT 54-11-10 V.4.0.3
'            Select Case tUT_PrintType
'                    Case "2" 'แบบรวมพิมพ์โดยตรง
'                        Call UT_ThaJustAlig(ptText)
'                    Case "3" 'แบบรวมพิมพ์โดยวินโดว์
                        If pnPrintTo = nEN_PrinterAndEJ Or pnPrintTo = nEN_PrinterOnly Then
                                oPrn.Print ptText
                        Else
                                 bW_LineFeed = False
                        End If
'            End Select
    Else
'oopos
            If Not (pnPrintTo = nEN_PrinterAndEJ Or pnPrintTo = nEN_PrinterOnly) Then
                     bW_LineFeed = False
            End If
            
            If tC_PrinterType = "0" And bC_PrnByOPOS Then
'                pbBold = False                     '*BG 57-04-28 PRJ10001TK-PH2-External Spec-Goods Return-00.02.06 FSBT.docx
'                pbHeight = False
            End If
            
            If pnPrintTo = nEN_PrinterAndEJ Or pnPrintTo = nEN_PrinterOnly Then
                    If pbLF Then
                        nC_LineNo = nC_LineNo + 1
                        If nC_LineNo <> 0 Then nC_LineNoEx = nC_LineNo
'                        If tC_PrinterType = "0" Then
'                            If (nC_LineNo = 1) And (nC_PageNo > 1) Then                'Ao 2013/07/25   แก้ไขการพิมพ์ Slip ผิด
'                                oOposSlip.TransactionPrint PTR_S_SLIP, PTR_TP_TRANSACTION 'Initial start  print by OPOS
'                                oOposSlip.RecLineSpacing = 10
'                                oOposSlip.Printnormal PTR_S_SLIP, "...Document Continued" & LF
'                                nC_LineNo = nC_LineNo + 1
'
'                                If Not bC_LineHeader Then
'                                    bC_LineHeader = True
'                                    For nI = 0 To 3
'                                        Call PrintText(oC_LineHeader(nI).tText, oC_LineHeader(nI).cPosition, , oC_LineHeader(nI).nAli, oC_LineHeader(nI).bLF, nEN_PrinterOnly)
'                                    Next nI
'                                    Call PrintText("", 0, , vbCenter, True)
'
'                                End If
'                                bC_LineHeader = False
'                            End If
'                        End If
                    End If
                    
                    Select Case pnAlign
                        Case vbLeftJustify      'left nothing to do
                             nLenTha = cUT_Print.UT_nThaAbsoluteCount(tC_LastTextPrint)
                             nLenTextPrint = (Len(tC_LastTextPrint) - nLenTha) * IIf(bC_LastFontBold, IIf(bC_LastFontHeight, 1, 2), 1)
                             nSpaceBlank = IIf(tC_PrinterType = "0", nVB_MaxCharSlip, nVB_MaxChar) - nLenTextPrint
                            If pcX = 5 Then
                                nSpace = IIf(tC_PrinterType = "0", nVB_MaxCharSlip, nVB_MaxChar) - 33
                                If nLenTextPrint > nSpace Then nSpace = nLenTextPrint
                                nSpace = nSpace - nLenTextPrint
                                nSpaceBlank = IIf(tC_PrinterType = "0", nVB_MaxCharSlip, nVB_MaxChar) - nSpace
                                If nSpaceBlank = 0 Then nSpace = 0
                                If nSpace < 0 Then nSpace = 0
                           ElseIf pcX = 5 - 0.3 Then 'cXMK
                                nSpace = IIf(tC_PrinterType = "0", nVB_MaxCharSlip, nVB_MaxChar) - 11
                                If nLenTextPrint > nSpace Then nSpace = nLenTextPrint
                                nSpace = nSpace - nLenTextPrint
                                nSpaceBlank = IIf(tC_PrinterType = "0", nVB_MaxCharSlip, nVB_MaxChar) - nSpace
                                If nSpaceBlank = 0 Then nSpace = 0
                                If nSpace < 0 Then nSpace = 0
                                If tC_PrinterType = "0" Then
                                    oOposSlip.Printnormal PTR_S_SLIP, IIf(nSpace > 0, Space(nSpace), "") & IIf(pbBold, ESC & IIf(pbHeight, "|3C", "|2C") & Mid(ptText, 1, nSpaceBlank) & ESC & "|1C", Mid(ptText, 1, nSpaceBlank)) & IIf(pbLF, LF, "")
                                Else
                                    oOpos.Printnormal PTR_S_RECEIPT, IIf(nSpace > 0, Space(nSpace), "") & IIf(pbBold, ESC & IIf(pbHeight, "|3C", "|2C") & ESC & "E1" & Mid(ptText, 1, nSpaceBlank) & ESC & "|1C", Mid(ptText, 1, nSpaceBlank)) & IIf(pbLF, LF, "")
                                End If
                            ElseIf pcX = 3.75 Then 'cXHR
                                nSpace = IIf(tC_PrinterType = "0", nVB_MaxCharSlip, nVB_MaxChar) - (IIf(tC_PrinterType = "0", nVB_MaxCharSlip, nVB_MaxChar) / 4)
                                If nLenTextPrint > nSpace Then nSpace = nLenTextPrint
                                nSpace = nSpace - nLenTextPrint
                                nSpaceBlank = IIf(tC_PrinterType = "0", nVB_MaxCharSlip, nVB_MaxChar) - nSpace
                                If nSpaceBlank = 0 Then nSpace = 0
                                If nSpace < 0 Then nSpace = 0
                                If tC_PrinterType = "0" Then
                                    oOposSlip.Printnormal PTR_S_SLIP, IIf(nSpace > 0, Space(nSpace), "") & IIf(pbBold, ESC & IIf(pbHeight, "|3C", "|2C") & Mid(ptText, 1, nSpaceBlank) & ESC & "|1C", Mid(ptText, 1, nSpaceBlank)) & IIf(pbLF, LF, "")
                                Else
                                    oOpos.Printnormal PTR_S_RECEIPT, IIf(nSpace > 0, Space(nSpace), "") & IIf(pbBold, ESC & IIf(pbHeight, "|3C", "|2C") & ESC & "E1" & Mid(ptText, 1, nSpaceBlank) & ESC & "|1C", Mid(ptText, 1, nSpaceBlank)) & IIf(pbLF, LF, "")
                                End If
                            ElseIf pcX = 3.5 Then 'cXHR
                                nSpace = IIf(tC_PrinterType = "0", nVB_MaxCharSlip, nVB_MaxChar) - (IIf(tC_PrinterType = "0", nVB_MaxCharSlip, nVB_MaxChar) / 4) - 2
                                If nLenTextPrint > nSpace Then nSpace = nLenTextPrint
                                nSpace = nSpace - nLenTextPrint
                                nSpaceBlank = IIf(tC_PrinterType = "0", nVB_MaxCharSlip, nVB_MaxChar) - nSpace
                                If nSpaceBlank = 0 Then nSpace = 0
                                If nSpace < 0 Then nSpace = 0
                                If tC_PrinterType = "0" Then
                                    oOposSlip.Printnormal PTR_S_SLIP, IIf(nSpace > 0, Space(nSpace), "") & IIf(pbBold, ESC & IIf(pbHeight, "|3C", "|2C") & Mid(ptText, 1, nSpaceBlank) & ESC & "|1C", Mid(ptText, 1, nSpaceBlank)) & IIf(pbLF, LF, "")
                                Else
                                    oOpos.Printnormal PTR_S_RECEIPT, IIf(nSpace > 0, Space(nSpace), "") & IIf(pbBold, ESC & IIf(pbHeight, "|3C", "|2C") & ESC & "E1" & Mid(ptText, 1, nSpaceBlank) & ESC & "|1C", Mid(ptText, 1, nSpaceBlank)) & IIf(pbLF, LF, "")
                                End If
                            ElseIf pcX = 1.25 Then 'cXHL
                                nSpace = IIf(tC_PrinterType = "0", nVB_MaxCharSlip, nVB_MaxChar) / 4
                                If nLenTextPrint > nSpace Then nSpace = nLenTextPrint
                                nSpace = nSpace - nLenTextPrint
                                nSpaceBlank = IIf(tC_PrinterType = "0", nVB_MaxCharSlip, nVB_MaxChar) - nSpace
                                 If nSpaceBlank = 0 Then nSpace = 0
                                 If nSpace < 0 Then nSpace = 0
                                If tC_PrinterType = "0" Then
                                    oOposSlip.Printnormal PTR_S_SLIP, IIf(nSpace > 0, Space(nSpace), "") & IIf(pbBold, ESC & IIf(pbHeight, "|3C", "|2C") & Mid(ptText, 1, nSpaceBlank) & ESC & "|1C", Mid(ptText, 1, nSpaceBlank)) & IIf(pbLF, LF, "")
                                Else
                                    oOpos.Printnormal PTR_S_RECEIPT, IIf(nSpace > 0, Space(nSpace), "") & IIf(pbBold, ESC & IIf(pbHeight, "|3C", "|2C") & ESC & "E1" & Mid(ptText, 1, nSpaceBlank) & ESC & "|1C", Mid(ptText, 1, nSpaceBlank)) & IIf(pbLF, LF, "")
                                End If
                            ElseIf pcX = 2.275 Then 'cX15 '*KT 52-05-02 PH 1.8 ML-SF-046 ในกรณีที่ Barcode  20 ตัว และ ชิ่อสินค้ายาว 20 อักษร  ระบบพิมพ์ Receipt ออกมาโดยชื่อสินค้าตกบรรทัด
                                nSpace = 15
                                If nLenTextPrint > nSpace Then
                                    nSpace = nLenTextPrint
                                    nSpaceBlank = (IIf(tC_PrinterType = "0", nVB_MaxCharSlip, nVB_MaxChar) - nSpace)  '+ cUT_Print.UT_nThaAbsoluteCount(ptText)
                                    nSpace = nSpace - nLenTextPrint
                                    If nLenTextPrint < IIf(tC_PrinterType = "0", nVB_MaxCharSlip, nVB_MaxChar) Then
                                        nSpace = 1
                                        nSpaceBlank = nSpaceBlank - 1
                                    End If
                                    nSpaceBlankOld = nSpaceBlank
                                    nLen = cUT_Print.UT_nThaAbsoluteCount(Mid(ptText, 1, nSpaceBlank))
                                    nLenTextFull = cUT_Print.UT_nThaAbsoluteCount(ptText)
                                    Do While Not nLen = 0
                                        nSpaceBlank = nSpaceBlank + nLen
                                        nLen = cUT_Print.UT_nThaAbsoluteCount(Mid(ptText, 1, nSpaceBlank))
                                        If ((nSpaceBlank - nLen) = nSpaceBlankOld) Or (nLenTextFull = nLen) Then
                                            nLen = 0
                                        End If
                                    Loop
                                Else
                                    nSpace = nSpace - nLenTextPrint
                                    nSpaceBlank = IIf(tC_PrinterType = "0", nVB_MaxCharSlip, nVB_MaxChar) - nSpace
                                End If
                                If nSpace < 0 Then nSpace = 0
                                If tC_PrinterType = "0" Then
                                    oOposSlip.Printnormal PTR_S_SLIP, IIf(nSpace > 0, Space(nSpace), "") & IIf(pbBold, ESC & IIf(pbHeight, "|3C", "|2C") & Mid(ptText, 1, nSpaceBlank) & ESC & "|1C", Mid(ptText, 1, nSpaceBlank)) & IIf(pbLF, LF, "")
                                Else
                                    oOpos.Printnormal PTR_S_RECEIPT, IIf(nSpace > 0, Space(nSpace), "") & IIf(pbBold, ESC & IIf(pbHeight, "|3C", "|2C") & ESC & "E1" & Mid(ptText, 1, nSpaceBlank) & ESC & "|1C", Mid(ptText, 1, nSpaceBlank)) & IIf(pbLF, LF, "")
                                End If
                            ElseIf pcX = 2.5 Then 'cXC
                                nSpace = (IIf(tC_PrinterType = "0", nVB_MaxCharSlip, nVB_MaxChar) / 2)
                                If nLenTextPrint > nSpace Then nSpace = nLenTextPrint
                                nSpace = nSpace - nLenTextPrint
                                nSpaceBlank = IIf(tC_PrinterType = "0", nVB_MaxCharSlip, nVB_MaxChar) - nSpace
                                If nSpace < 0 Then nSpace = 0
                                If tC_PrinterType = "0" Then
                                    oOposSlip.Printnormal PTR_S_SLIP, IIf(nSpace > 0, Space(nSpace), "") & IIf(pbBold, ESC & IIf(pbHeight, "|3C", "|2C") & Mid(ptText, 1, nSpaceBlank) & ESC & "|1C", Mid(ptText, 1, nSpaceBlank)) & IIf(pbLF, LF, "")
                                Else
                                    oOpos.Printnormal PTR_S_RECEIPT, IIf(nSpace > 0, Space(nSpace), "") & IIf(pbBold, ESC & IIf(pbHeight, "|3C", "|2C") & ESC & "E1" & Mid(ptText, 1, nSpaceBlank) & ESC & "|1C", Mid(ptText, 1, nSpaceBlank)) & IIf(pbLF, LF, "")
                                End If
                            ElseIf pcX = 1.875 Then  'cXC-cXHC
                                nSpace = (IIf(tC_PrinterType = "0", nVB_MaxCharSlip, nVB_MaxChar) / 2) - 5
                                If nLenTextPrint > nSpace Then nSpace = nLenTextPrint
                                nSpace = nSpace - nLenTextPrint
                                nSpaceBlank = IIf(tC_PrinterType = "0", nVB_MaxCharSlip, nVB_MaxChar) - nSpace
                                If nSpace < 0 Then nSpace = 0
                                If tC_PrinterType = "0" Then
                                    oOposSlip.Printnormal PTR_S_SLIP, IIf(nSpace > 0, Space(nSpace), "") & IIf(pbBold, ESC & IIf(pbHeight, "|3C", "|2C") & Mid(ptText, 1, nSpaceBlank) & ESC & "|1C", Mid(ptText, 1, nSpaceBlank)) & IIf(pbLF, LF, "")
                                Else
                                    oOpos.Printnormal PTR_S_RECEIPT, IIf(nSpace > 0, Space(nSpace), "") & IIf(pbBold, ESC & IIf(pbHeight, "|3C", "|2C") & ESC & "E1" & Mid(ptText, 1, nSpaceBlank) & ESC & "|1C", Mid(ptText, 1, nSpaceBlank)) & IIf(pbLF, LF, "")
                                End If
                             ElseIf pcX = 1.08 Then
                                nSpace = 7
                                If nLenTextPrint > nSpace Then nSpace = nLenTextPrint
                                nSpace = nSpace - nLenTextPrint
                                nSpaceBlank = IIf(tC_PrinterType = "0", nVB_MaxCharSlip, nVB_MaxChar) - nSpace
                                If nSpace < 0 Then nSpace = 0
                                If tC_PrinterType = "0" Then
                                    oOposSlip.Printnormal PTR_S_SLIP, IIf(nSpace > 0, Space(nSpace), "") & IIf(pbBold, ESC & IIf(pbHeight, "|3C", "|2C") & Mid(ptText, 1, nSpaceBlank) & ESC & "|1C", Mid(ptText, 1, nSpaceBlank)) & IIf(pbLF, LF, "")
                                Else
                                    oOpos.Printnormal PTR_S_RECEIPT, IIf(nSpace > 0, Space(nSpace), "") & IIf(pbBold, ESC & IIf(pbHeight, "|3C", "|2C") & ESC & "E1" & Mid(ptText, 1, nSpaceBlank) & ESC & "|1C", Mid(ptText, 1, nSpaceBlank)) & IIf(pbLF, LF, "")
                                End If
                             ElseIf pcX = 1 Then  'cX5
                                nSpace = 5
                                If nLenTextPrint > nSpace Then nSpace = nLenTextPrint
                                nSpace = nSpace - nLenTextPrint
                                nSpaceBlank = IIf(tC_PrinterType = "0", nVB_MaxCharSlip, nVB_MaxChar) - nSpace
                                If nSpace < 0 Then nSpace = 0
                                If tC_PrinterType = "0" Then
                                    oOposSlip.Printnormal PTR_S_SLIP, IIf(nSpace > 0, Space(nSpace), "") & IIf(pbBold, ESC & IIf(pbHeight, "|3C", "|2C") & Mid(ptText, 1, nSpaceBlank) & ESC & "|1C", Mid(ptText, 1, nSpaceBlank)) & IIf(pbLF, LF, "")
                                Else
                                    oOpos.Printnormal PTR_S_RECEIPT, IIf(nSpace > 0, Space(nSpace), "") & IIf(pbBold, ESC & IIf(pbHeight, "|3C", "|2C") & ESC & "E1" & Mid(ptText, 1, nSpaceBlank) & ESC & "|1C", Mid(ptText, 1, nSpaceBlank)) & IIf(pbLF, LF, "")
                                End If
                             ElseIf pcX = 0.625 Then 'cXHL-cXHC
                                nSpace = 5
                                If nLenTextPrint > nSpace Then nSpace = nLenTextPrint
                                nSpace = nSpace - nLenTextPrint
                                nSpaceBlank = IIf(tC_PrinterType = "0", nVB_MaxCharSlip, nVB_MaxChar) - nSpace
                                If nSpace < 0 Then nSpace = 0
                                If tC_PrinterType = "0" Then
                                    oOposSlip.Printnormal PTR_S_SLIP, IIf(nSpace > 0, Space(nSpace), "") & IIf(pbBold, ESC & IIf(pbHeight, "|3C", "|2C") & Mid(ptText, 1, nSpaceBlank) & ESC & "|1C", Mid(ptText, 1, nSpaceBlank)) & IIf(pbLF, LF, "")
                                Else
                                    oOpos.Printnormal PTR_S_RECEIPT, IIf(nSpace > 0, Space(nSpace), "") & IIf(pbBold, ESC & IIf(pbHeight, "|3C", "|2C") & ESC & "E1" & Mid(ptText, 1, nSpaceBlank) & ESC & "|1C", Mid(ptText, 1, nSpaceBlank)) & IIf(pbLF, LF, "")
                                End If
                            ElseIf pcX = 0.425 Then 'cX3
                                  nSpace = 3
                                If nLenTextPrint > nSpace Then nSpace = nLenTextPrint
                                nSpace = nSpace - nLenTextPrint
                                nSpaceBlank = IIf(tC_PrinterType = "0", nVB_MaxCharSlip, nVB_MaxChar) - nSpace
                                If nSpace < 0 Then nSpace = 0
                                If tC_PrinterType = "0" Then
                                    oOposSlip.Printnormal PTR_S_SLIP, IIf(nSpace > 0, Space(nSpace), "") & IIf(pbBold, ESC & IIf(pbHeight, "|3C", "|2C") & Mid(ptText, 1, nSpaceBlank) & ESC & "|1C", Mid(ptText, 1, nSpaceBlank)) & IIf(pbLF, LF, "")
                                Else
                                    oOpos.Printnormal PTR_S_RECEIPT, IIf(nSpace > 0, Space(nSpace), "") & IIf(pbBold, ESC & IIf(pbHeight, "|3C", "|2C") & ESC & "E1" & Mid(ptText, 1, nSpaceBlank) & ESC & "|1C", Mid(ptText, 1, nSpaceBlank)) & IIf(pbLF, LF, "")
                                End If
                             ElseIf pcX = 0.25 Then 'cX2
                                nSpace = 2
                                If nLenTextPrint > nSpace Then nSpace = nLenTextPrint
                                nSpace = nSpace - nLenTextPrint
                                nSpaceBlank = IIf(tC_PrinterType = "0", nVB_MaxCharSlip, nVB_MaxChar) - nSpace
                                If nSpace < 0 Then nSpace = 0
                                If tC_PrinterType = "0" Then
                                     oOposSlip.Printnormal PTR_S_SLIP, IIf(nSpace > 0, Space(nSpace), "") & IIf(pbBold, ESC & IIf(pbHeight, "|3C", "|2C") & Mid(ptText, 1, nSpaceBlank) & ESC & "|1C", Mid(ptText, 1, nSpaceBlank)) & IIf(pbLF, LF, "")
                                Else
                                     oOpos.Printnormal PTR_S_RECEIPT, IIf(nSpace > 0, Space(nSpace), "") & IIf(pbBold, ESC & IIf(pbHeight, "|3C", "|2C") & ESC & "E1" & Mid(ptText, 1, nSpaceBlank) & ESC & "|1C", Mid(ptText, 1, nSpaceBlank)) & IIf(pbLF, LF, "")
                                End If
                             Else 'cXL
                                nSpace = 0
                                If tC_PrinterType = "0" Then
                                    oOposSlip.Printnormal PTR_S_SLIP, IIf(pbBold, ESC & IIf(pbHeight, "|3C", "|2C") & ptText & ESC & "|1C", ptText) & IIf(pbLF, LF, "")
                                Else
                                    oOpos.Printnormal PTR_S_RECEIPT, IIf(pbBold, ESC & IIf(pbHeight, "|3C", "|2C") & ESC & "E1" & ptText & ESC & "|1C", ptText) & IIf(pbLF, LF, "")
                                End If
                             End If
                        Case vbCenter      'center
                               If tC_PrinterType = "0" Then
                                    nLenTha = cUT_Print.UT_nThaAbsoluteCount(ptText)                  '*Ao 2014/07/19             V 4.2.138  TK-ISS3000PH2-524
                                    nLenTha = Len(ptText) - nLenTha
                                    nLenTextPrint = Len(ptText) * IIf(pbBold, IIf(pbHeight, 1, 2), 1)
                                    nSpaceBlank = nVB_MaxCharSlip - nLenTextPrint
                                    nSpace = nSpaceBlank \ 2
                                    If nSpace < 0 Then nSpace = 0
                                    oOposSlip.Printnormal PTR_S_SLIP, IIf(nSpace > 0, Space(nSpace), "") & IIf(pbBold, ESC & IIf(pbHeight, "|3C", "|2C") & ptText & ESC & "|1C", ptText) & IIf(pbLF, LF, "")
'                                    oOposSlip.Printnormal PTR_S_RECEIPT, Chr(27) & Chr(97) & "1" & IIf(pbBold, ESC & IIf(pbHeight, "|3C", "|2C") & ESC & "E1" & ptText & ESC & "|1C", ptText) & Chr(27) & Chr(97) & "0" & IIf(pbLF, LF, "")
                                Else
                                     oOpos.Printnormal PTR_S_RECEIPT, Chr(27) & Chr(97) & "1" & IIf(pbBold, ESC & IIf(pbHeight, "|3C", "|2C") & ESC & "E1" & ptText & ESC & "|1C", ptText) & Chr(27) & Chr(97) & "0" & IIf(pbLF, LF, "")
                                End If
                        Case vbRightJustify      'right justify
                             nLenTha = cUT_Print.UT_nThaAbsoluteCount(tC_LastTextPrint)
                             nLenTextPrint = (Len(tC_LastTextPrint) - nLenTha) * IIf(bC_LastFontBold, IIf(bC_LastFontHeight, 1, 2), 1)
                             nSpaceBlank = IIf(tC_PrinterType = "0", nVB_MaxCharSlip, nVB_MaxChar) - nLenTextPrint
                             If pcX = 5 Then 'cXR
                                nSpaceBlank = IIf(tC_PrinterType = "0", nVB_MaxCharSlip, nVB_MaxChar) - nLenTextPrint
                                nSpace = nSpaceBlank - (Len(ptText) * IIf(pbBold, IIf(pbHeight, 1, 2), 1))
                                If nSpace < 0 Then nSpace = 0
                            ElseIf pcX = (5 - 0.3) + 0.125 Then
                                nSpaceBlank = (IIf(tC_PrinterType = "0", nVB_MaxCharSlip, nVB_MaxChar) - 1) - nLenTextPrint
                                nSpace = nSpaceBlank - (Len(ptText) * IIf(pbBold, IIf(pbHeight, 1, 2), 1))
                                If nSpace < 0 Then nSpace = 0
                            ElseIf pcX = 5 - 0.3 Then 'cXMK
                                nSpaceBlank = (IIf(tC_PrinterType = "0", nVB_MaxCharSlip, nVB_MaxChar) - 2) - nLenTextPrint
                                nSpace = nSpaceBlank - (Len(ptText) * IIf(pbBold, IIf(pbHeight, 1, 2), 1))
                                If nSpace < 0 Then nSpace = 0
                             ElseIf pcX = 3.75 Then  'cXHR
                                nSpaceBlank = (IIf(tC_PrinterType = "0", nVB_MaxCharSlip, nVB_MaxChar) - (IIf(tC_PrinterType = "0", nVB_MaxCharSlip, nVB_MaxChar) / 4))
                                nSpace = nSpaceBlank - (Len(ptText) * IIf(pbBold, IIf(pbHeight, 1, 2), 1))
                                If nSpace < 0 Then nSpace = 0
                             ElseIf pcX = (3.75 + 0.5) Then 'cXHR + cX4
                                nSpaceBlank = (IIf(tC_PrinterType = "0", nVB_MaxCharSlip, nVB_MaxChar) - (IIf(tC_PrinterType = "0", nVB_MaxCharSlip, nVB_MaxChar) / 4)) + 4
                                nSpace = nSpaceBlank - (Len(ptText) * IIf(pbBold, IIf(pbHeight, 1, 2), 1))
                                If nSpace < 0 Then nSpace = 0
                             ElseIf pcX = 3.125 Then  'cXC + cXHC
                                nSpaceBlank = (IIf(tC_PrinterType = "0", nVB_MaxCharSlip, nVB_MaxChar) / 2) + (IIf(tC_PrinterType = "0", nVB_MaxCharSlip, nVB_MaxChar) / 8) - nLenTextPrint
                                nSpace = nSpaceBlank - Len(ptText)
                                If nSpace < 0 Then nSpace = 0
                             ElseIf pcX = 3 Then  'cXC+cX4
                                nSpaceBlank = (IIf(tC_PrinterType = "0", nVB_MaxCharSlip, nVB_MaxChar) / 2) + 4
                                nSpace = nSpaceBlank - (Len(ptText) * IIf(pbBold, IIf(pbHeight, 1, 2), 1))
                                If nSpace < 0 Then nSpace = 0
                             ElseIf pcX = 2.5 Then 'cXC
                                nSpaceBlank = IIf(tC_PrinterType = "0", nVB_MaxCharSlip, nVB_MaxChar) / 2
                                nSpace = nSpaceBlank - (Len(ptText) * IIf(pbBold, IIf(pbHeight, 1, 2), 1))
                                If nSpace < 0 Then nSpace = 0
                             ElseIf pcX = 2.25 Then 'cXC-cX2
                                nSpaceBlank = (IIf(tC_PrinterType = "0", nVB_MaxCharSlip, nVB_MaxChar) / 2) - 2
                                nSpace = nSpaceBlank - (Len(ptText) * IIf(pbBold, IIf(pbHeight, 1, 2), 1))
                                If nSpace < 0 Then nSpace = 0
                             ElseIf pcX = 2 Then  'cXC-cX4
                                nSpaceBlank = (IIf(tC_PrinterType = "0", nVB_MaxCharSlip, nVB_MaxChar) / 2) - 4
                                nSpace = nSpaceBlank - (Len(ptText) * IIf(pbBold, IIf(pbHeight, 1, 2), 1))
                                If nSpace < 0 Then nSpace = 0
                             ElseIf pcX = 1.25 Then 'cXHL
                                nSpace = IIf(tC_PrinterType = "0", nVB_MaxCharSlip, nVB_MaxChar) / 4
                                If nLenTextPrint > nSpace Then nSpace = nLenTextPrint
                                nSpaceBlank = IIf(tC_PrinterType = "0", nVB_MaxCharSlip, nVB_MaxChar) - nSpace
                                If nSpace < 0 Then nSpace = 0
                            ElseIf pcX = 0.5 Then 'cX4
                                nSpaceBlank = 4
                                nSpace = nSpaceBlank - (Len(ptText) * IIf(pbBold, IIf(pbHeight, 1, 2), 1))
                                If nSpace < 0 Then nSpace = 0
                             Else 'cXL
                                nSpace = nSpaceBlank - Len(ptText)
                                If nSpace < 0 Then nSpace = 0
                             End If
                            If nSpaceBlank > 0 Then
                                If pbCR Then                         '*Ao 2013/07/16
                                    pbBold = True
                                End If
                                If pbBold Then
                                    If tC_PrinterType = "0" Then
    '                                      oOposSlip.Printnormal PTR_S_SLIP, IIf(nSpace > 0, Space(nSpace), "") & ESC & IIf(pbHeight, "|3C", "|2C") & Mid(ptText, 1, nSpaceBlank) & ESC & "|1C" & IIf(pbLF, LF, "")          '*Ao 2013/07/16
                                        nSpace = 0
                                        oOposSlip.Printnormal PTR_S_SLIP, IIf(nSpace > 0, Space(nSpace), "") & IIf(pbBold, ESC & IIf(pbHeight, "|3C", "|2C") & ptText & ESC & "|1C", ptText) & IIf(pbLF, LF, "")
                                    Else
                                        oOpos.Printnormal PTR_S_RECEIPT, IIf(nSpace > 0, Space(nSpace), "") & ESC & IIf(pbHeight, "|3C", "|2C") & ESC & "E1" & Mid(ptText, 1, nSpaceBlank) & ESC & "|1C" & IIf(pbLF, LF, "")
                                    End If
                                Else
                                    If tC_PrinterType = "0" Then
                                       oOposSlip.Printnormal PTR_S_SLIP, IIf(nSpace > 0, Space(nSpace), "") & Mid(ptText, 1, nSpaceBlank) & IIf(pbLF, LF, "")
                                    Else
                                        oOpos.Printnormal PTR_S_RECEIPT, IIf(nSpace > 0, Space(nSpace), "") & Mid(ptText, 1, nSpaceBlank) & IIf(pbLF, LF, "")
                                    End If
                                End If
                            Else
                                If tC_PrinterType = "0" Then
                                    oOposSlip.Printnormal PTR_S_SLIP, "" & IIf(pbLF, LF, "")
                                Else
                                    oOpos.Printnormal PTR_S_RECEIPT, "" & IIf(pbLF, LF, "")
                                End If
                            End If
                    End Select
                    
                    If tC_PrinterType = "0" And bC_PrnByOPOS Then
                        If pbLF Then
                            tC_LastTextPrint = ""
                            bC_LastFontBold = False
                            bC_LastFontHeight = False
                        Else
                            Select Case pnAlign
                                Case vbLeftJustify      'left nothing to do
                                    tC_LastTextPrint = tC_LastTextPrint & Space(nSpace) & ptText
                                Case vbCenter      'center
                                    tC_LastTextPrint = tC_LastTextPrint & Space(nSpace) & ptText
                                Case vbRightJustify      'right justify
                                    tC_LastTextPrint = tC_LastTextPrint & Space(nSpace) & ptText
                            End Select
                            bC_LastFontBold = pbBold
                            bC_LastFontHeight = pbHeight
                        End If
'                        If nC_LineNo Mod nC_LineMax = 0 And nC_LineNo > 0 And pbLF Then
'                        If (nC_LineNo > nC_LineMax) And (nC_LineNo > 0) And pbLF Then '*BG 56-06-05  TK-ISS3000-078 ,TK -ISS3000 - 216 - ST,TK-ISS3000-181-ST
'                        If (nC_LineNo > nC_LineMax) And (nC_LineNo > 0) And pbLF And bVB_LastLine Then  '*BG 56-06-04 TK-ISS3000-078
'                            oOposSlip.Printnormal PTR_S_SLIP, "Document Continued..."           '*Ao 2013/07/25   ยังไม่มีการเช็คใน กรณีที่ พิมพ์พอดี  แล้วไม่ต้องพิมพ์ต่อ   *ยังคิดไม่ได้
'                            nC_LineNo = nC_LineNo + 1
'                            Call EndDoc(oW_Logo, False)
'                            wCNMsgErrSlip.bW_Wait = True
'                            wCNMsgErrSlip.bW_Printing = True
'                            wCNMsgErrSlip.nW_Line = nC_LineNoEx
'                            wCNMsgErrSlip.tW_Msg = "...Continued"
'                            wCNMsgErrSlip.tW_TransTypeSub = tW_TransTypeSub     '*Ao 2013/08/30  CommSheet TK-ISS3000-336-UAT
'                            wCNMsgErrSlip.Show vbModal
''                            MsgBox "ResultCode = " & oOposSlip.ResultCode & vbCrLf & "ResultCodeExtended = " & oOposSlip.ResultCodeExtended
'                            nLevel = oOposSlip.nCheckHealth(OPOS_CH_INTERNAL)       '*Ao 2013/07/26 CommSheet TK-ISS3000-347-UAT        V 4.0.62
'                            If nLevel = OPOS_E_ILLEGAL Then
'                                wCNMsgSlip.Show vbModal
'                            End If
'                            nC_PageNo = nC_PageNo + 1
'                            nC_LineNo = 0
'                            wCNMsgErrSlip.bW_Wait = False
'                            wCNMsgErrSlip.bW_Printing = False
'                            wCNMsgErrSlip.nW_Line = 0
'                            wCNMsgErrSlip.tW_Msg = "...Continued"
'                            wCNMsgErrSlip.tW_TransTypeSub = tW_TransTypeSub     '*Ao 2013/08/30  CommSheet TK-ISS3000-336-UAT
'                            wCNMsgErrSlip.Show vbModal
'                            oOposSlip.TransactionPrint PTR_S_SLIP, PTR_TP_TRANSACTION 'Initial start  print by OPOS
'                            oOposSlip.RecLineSpacing = 10
'                            oOposSlip.Printnormal PTR_S_SLIP, "...Document Continued" & LF
'                            nC_LineNo = nC_LineNo + 1
'
'                            If Not bC_LineHeader Then
'                                bC_LineHeader = True
'                                For nI = 0 To 3
'                                    Call PrintText(oC_LineHeader(nI).tText, oC_LineHeader(nI).cPosition, , oC_LineHeader(nI).nAli, oC_LineHeader(nI).bLF, nEN_PrinterOnly)
'                                Next nI
'                                Call PrintText("", 0, , vbCenter, True)
'                            End If
'                            bC_LineHeader = False
'                        End If
                        '*Ao 2014/04/23  '*Ao 2014/04/23  PRJ10001TK-PH2-External Spec-Auto Release Paper at Slip Printer
                        If (nC_LineNo > nC_LineMax) And (nC_LineNo > 0) And pbLF And bVB_LastLine Then
                            '*Bump 57-09-30 10.15am. CommSheet TK-ISS3000-555 Add log เพื่อ Monitor
                            tSlipCheck = ""
                            tSlipCheck = "nC_LineNo." & nC_LineNo & ",nC_LineMax:" & nC_LineMax & ",pbLF:" & pbLF & ",bVB_LastLine:" & bVB_LastLine
                            Call SP_CRTxLogPrg("Slip Printer: Slip printing in progress[Document Continued...]," & tSlipCheck & "cPrinter.PrintText: " & ",TerminalNum:" & tVB_CNTerminalNum & " TransNo.:" & tVB_Running & " TransType:" & tVB_Transaction)
                           '-------------
                            oOposSlip.Printnormal PTR_S_SLIP, "Document Continued..."
                            nC_LineNo = nC_LineNo + 1
                            nC_LineNoEx = nC_LineNo
                            
                            Call SP_CRTxLogPrg("SLIP PRINTER-cPrinter.PrintText-TransType:" & tW_TransTypeSub & ",print Document Continued..." & ",Page:" & nC_PageNo)  '*Bump 58-01-22 เก็บ Slip Printer -Monitor เครื่องแฮ้ง
                            
                            If oOposSlip.SlpEmpty = True Then
                                wCNMsgErrSlip.tW_Msg = ""                                                                   'Press Enter to print
                                wCNMsgErrSlip.tW_TransTypeSub = tW_TransTypeSub           '*Ao 2014/04/23  PRJ10001TK-PH2-External Spec-Auto Release Paper at Slip Printer-00.02.03
                                wCNMsgErrSlip.bW_Printing = False
                                wCNMsgErrSlip.bW_Wait = False
                                wCNMsgErrSlip.nW_Line = 0
                                wCNMsgErrSlip.Show vbModal
                                Call SP_CRTxLogPrg("SLIP PRINTER-cPrinter.PrintText-TransType:" & tW_TransTypeSub & ",Press Enter to print")  '*Bump 58-01-22 เก็บ Slip Printer -Monitor เครื่องแฮ้ง
                            End If
                            
                            Call EndDoc(oW_Logo, False)
                            wCNMsgErrSlip.tW_Msg = ""                                                                  'Slip printing in progress"
                            wCNMsgErrSlip.tW_TransTypeSub = tW_TransTypeSub           '*Ao 2014/04/23  PRJ10001TK-PH2-External Spec-Auto Release Paper at Slip Printer-00.02.03
                            wCNMsgErrSlip.bW_Printing = True
                            wCNMsgErrSlip.bW_Wait = True
                            wCNMsgErrSlip.nW_Line = nC_LineNoEx
                            wCNMsgErrSlip.Show vbModal
                            
                            Call SP_CRTxLogPrg("SLIP PRINTER-cPrinter.PrintText-TransType:" & tW_TransTypeSub & ",Slip printing in progress")  '*Bump 58-01-22 เก็บ Slip Printer -Monitor เครื่องแฮ้ง
                            
                            nLevel = oOposSlip.nCheckHealth(OPOS_CH_INTERNAL)       '*Ao 2013/07/26 CommSheet TK-ISS3000-347-UAT        V 4.0.62
                            If nLevel = OPOS_E_ILLEGAL Then
                                wCNMsgSlip.Show vbModal
                                Call SP_CRTxLogPrg("SLIP PRINTER-cPrinter.PrintText-TransType:" & tW_TransTypeSub & ",OPOS_E_ILLEGAL")  '*Bump 58-01-22 เก็บ Slip Printer -Monitor เครื่องแฮ้ง
                            End If
                            wCNMsgErrSlip.tW_Msg = "...Continued"
                            wCNMsgErrSlip.tW_TransTypeSub = tW_TransTypeSub
                            wCNMsgErrSlip.bW_Printing = False
                            wCNMsgErrSlip.bW_Wait = False
                            wCNMsgErrSlip.nW_Line = 0
                            wCNMsgErrSlip.Show vbModal
                                        
                            Call SP_CRTxLogPrg("SLIP PRINTER-cPrinter.PrintText-TransType:" & tW_TransTypeSub & ",set ...Continued")  '*Bump 58-01-22 เก็บ Slip Printer -Monitor เครื่องแฮ้ง
                             
                            nC_PageNo = nC_PageNo + 1
                            nC_LineNo = 0
                            oOposSlip.TransactionPrint PTR_S_SLIP, PTR_TP_TRANSACTION 'Initial start  print by OPOS
                            oOposSlip.RecLineSpacing = 10
                            oOposSlip.Printnormal PTR_S_SLIP, "...Document Continued" & LF
                            '*Bump 57-09-30 10.15am. CommSheet TK-ISS3000-555 Add log เพื่อ Monitor
                            tSlipCheck = ""
                            tSlipCheck = "nC_LineNo." & nC_LineNo & ",nC_LineMax:" & nC_LineMax & ",pbLF:" & pbLF & ",bVB_LastLine:" & bVB_LastLine & ",nC_PageNo:" & nC_PageNo
                            'Call SP_CRTxLogPrg("Slip Printer: Slip printing in progress[...Document Continued]," & tSlipCheck & "cPrinter.PrintText: " & ",TerminalNum:" & tVB_CNTerminalNum & " TransNo.:" & tVB_Running & " TransType:" & tVB_Transaction)
                            Call SP_CRTxLogPrg("SLIP PRINTER-cPrinter.PrintText-TransType:" & tW_TransTypeSub & ",print ...Document Continued" & ",Page:" & nC_PageNo)  '*Bump 58-01-22 เก็บ Slip Printer -Monitor เครื่องแฮ้ง
                            '-------------
                            nC_LineNo = nC_LineNo + 1
                            tC_LastTextPrint = ""               '*Ao 2014/07/18  TK-ISS3000PH2-522
                            nSpace = 0
                            bC_LineHeader = False
                            If Not bC_LineHeader Then
                                bC_LineHeader = True
                                For nI = 0 To 3
'                                    Call PrintText(oC_LineHeader(nI).tText, oC_LineHeader(nI).cPosition, , oC_LineHeader(nI).nAli, oC_LineHeader(nI).bLF, nEN_PrinterOnly)
                                '----------------------------------------------------------'*Ao 2014/07/18  TK-ISS3000PH2-522
                                    ptText = oC_LineHeader(nI).tText
                                    pbLF = oC_LineHeader(nI).bLF
                                    pnAlign = oC_LineHeader(nI).nAli
                                    Select Case nI
                                        Case 0                                  'Case vbLeftJustify
                                            pbBold = True
                                            pbHeight = False
                                            nSpace = 0
                                            oOposSlip.Printnormal PTR_S_SLIP, IIf(pbBold, ESC & IIf(pbHeight, "|3C", "|2C") & ptText & ESC & "|1C", ptText) & IIf(pbLF, LF, "")
                                        Case 1                                  'Case vbRightJustify
                                            pbBold = False
                                            pbHeight = False
                                            nLenTha = cUT_Print.UT_nThaAbsoluteCount(tC_LastTextPrint)
                                            nLenTextPrint = (Len(tC_LastTextPrint) - nLenTha) * IIf(bC_LastFontBold, IIf(bC_LastFontHeight, 1, 2), 1)
                                            nSpaceBlank = IIf(tC_PrinterType = "0", nVB_MaxCharSlip, nVB_MaxChar) - nLenTextPrint
                                            nSpace = nSpaceBlank - (Len(ptText) * IIf(pbBold, IIf(pbHeight, 1, 2), 1))
                                            If nSpace < 0 Then nSpace = 0
                                            oOposSlip.Printnormal PTR_S_SLIP, IIf(nSpace > 0, Space(nSpace), "") & Mid(ptText, 1, nSpaceBlank) & IIf(pbLF, LF, "")
                                        Case 2
                                            pbBold = True
                                            pbHeight = False
                                            nSpace = 0
                                            oOposSlip.Printnormal PTR_S_SLIP, IIf(pbBold, ESC & IIf(pbHeight, "|3C", "|2C") & ptText & ESC & "|1C", ptText) & IIf(pbLF, LF, "")
                                        Case 3
                                            pbBold = False
                                            pbHeight = False
                                            nLenTha = cUT_Print.UT_nThaAbsoluteCount(tC_LastTextPrint)
                                            nLenTextPrint = (Len(tC_LastTextPrint) - nLenTha) * IIf(bC_LastFontBold, IIf(bC_LastFontHeight, 1, 2), 1)
                                            nSpaceBlank = IIf(tC_PrinterType = "0", nVB_MaxCharSlip, nVB_MaxChar) - nLenTextPrint
                                            nSpace = nSpaceBlank - (Len(ptText) * IIf(pbBold, IIf(pbHeight, 1, 2), 1))
                                            If nSpace < 0 Then nSpace = 0
                                            oOposSlip.Printnormal PTR_S_SLIP, IIf(nSpace > 0, Space(nSpace), "") & Mid(ptText, 1, nSpaceBlank) & IIf(pbLF, LF, "")
                                    End Select
                                    
                                    If pbLF Then
                                        tC_LastTextPrint = ""
                                        bC_LastFontBold = False
                                        bC_LastFontHeight = False
                                        nC_LineNo = nC_LineNo + 1
                                    Else
                                        Select Case pnAlign
                                            Case vbLeftJustify
                                                tC_LastTextPrint = tC_LastTextPrint & Space(nSpace) & ptText
                                            Case vbCenter
                                                tC_LastTextPrint = tC_LastTextPrint & Space(nSpace) & ptText
                                            Case vbRightJustify
                                                tC_LastTextPrint = tC_LastTextPrint & Space(nSpace) & ptText
                                        End Select
                                        bC_LastFontBold = pbBold
                                        bC_LastFontHeight = pbHeight
                                    End If
                                    '----------------------------------------------------------'*Ao 2014/07/18  TK-ISS3000PH2-522
                                Next nI
'                                Call PrintText("", 0, , vbCenter, True)
                                oOposSlip.Printnormal PTR_S_SLIP, "" & LF
                                nC_LineNo = nC_LineNo + 1
                            End If
                            bC_LineHeader = False
                        End If
                    Else
                        If tC_PrinterType <> "0" And bC_PrnByOPOS Then
                            If pbLF Then
                                tC_LastTextPrint = ""
                                bC_LastFontBold = False
                                bC_LastFontHeight = False
                            Else
                                Select Case pnAlign
                                    Case vbLeftJustify      'left nothing to do
                                        tC_LastTextPrint = tC_LastTextPrint & Space(nSpace) & ptText
                                    Case vbCenter      'center
                                        tC_LastTextPrint = tC_LastTextPrint & Space(nSpace) & ptText
                                    Case vbRightJustify      'right justify
                                        tC_LastTextPrint = tC_LastTextPrint & Space(nSpace) & ptText
                                End Select
                                bC_LastFontBold = pbBold
                                bC_LastFontHeight = pbHeight
                            End If
                        End If
                    End If
            End If
      End If
End If
'return
PrintText = True

Exit Function '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
ErrHandler:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Function



Private Function W_GETnTextDigit(ByVal ptText As String) As Long
Dim tTmp As String, nTmp As Long

            tTmp = "ุ ู ่ ้ ๊ ๋ ์ ิ ี ึ ื "
            
            For nTmp = 1 To Len(tTmp)
                    Replace ptText, Mid(tTmp, nTmp, 1), vbNullString
            Next
            W_GETnTextDigit = Len(ptText)
            
        
End Function

Private Sub W_GETnBaseLineOffSet()
        Dim uTM As TEXTMETRIC
        Call GetTextMetrics(oPrn.hDC, uTM)
        cW_TextHeight = oPrn.TextHeight(Chr(uTM.tmDefaultChar))
        cW_BaseLineOffset = (cW_TextHeight * uTM.tmAscent) / uTM.tmHeight
        
End Sub

Private Sub Class_Initialize()
       
        Set oPrn = Printer
        oPrn.ScaleMode = vbCentimeters
         Call W_GETnBaseLineOffSet
         bW_LineFeed = True
         'oPrn.CurrentY = cW_BaseLineOffset
End Sub

Public Function NewPage() As Boolean
On Error GoTo ErrHandler
    
        oPrn.NewPage
        oPrn.CurrentY = cW_BaseLineOffset
    
'return
NewPage = True

Exit Function '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
ErrHandler:

End Function

Public Function KillDoc() As Boolean
On Error GoTo ErrHandler
    
        oPrn.KillDoc
    
'return
KillDoc = True

Exit Function '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
ErrHandler:

End Function

Public Sub ClearBufferOPOSPrn()
    'Dim DLE As String '*KT 51-12-22 Phase 2 Clear Buffer
    'Dim DC4 As String '*KT 51-12-22 Phase 2 Clear Buffer
On Error GoTo ErrHandler
    'DLE = Chr(16) '*KT 51-12-22 Phase 2 Clear Buffer
    'DC4 = Chr(20) '*KT 51-12-22 Phase 2 Clear Buffer
    With oOpos                          '*Ao 2013/01/24   Print FRONTTECH
        .ClearOutput
    End With
    Exit Sub
ErrHandler:
    
End Sub

Public Sub ClearBufferOPOSlip()
  
On Error GoTo ErrHandler
    'DLE = Chr(16) '*KT 51-12-22 Phase 2 Clear Buffer
    'DC4 = Chr(20) '*KT 51-12-22 Phase 2 Clear Buffer
    With oOposSlip                          '*Ao 2013/01/24   Print FRONTTECH
        .ClearOutput
    End With
    Exit Sub
ErrHandler:
    
End Sub

Public Function EndDoc(ByVal poPicture As IPictureDisp, Optional ByVal pbDrawer As Boolean = True, Optional ByVal pbPrnLogo As Boolean = True) As Boolean
    Dim LF As String
    Dim ESC As String
    Dim DLE As String '*KT 51-12-22 Phase 2 Clear Buffer
    Dim DC4 As String '*KT 51-12-22 Phase 2 Clear Buffer
    Dim nI As Long  '*KT 51-10-14 V 3.0.135/V 3.0.200
    Dim iRet&, tPrint$, ipcWritten&
On Error GoTo ErrHandler
    LF = Chr(10)
    ESC = Chr(27)
    DLE = Chr(16) '*KT 51-12-22 Phase 2 Clear Buffer
    DC4 = Chr(20) '*KT 51-12-22 Phase 2 Clear Buffer
On Error GoTo ErrHandler
        If bC_PrnByOPOS Then
             
        'oopos
            If tC_PrinterType = "0" Then
'                If bVB_CNPrnOPOS Then Call SP_SHWxErrPrinter(False, "0")       '*Ao 2013/07/18
                 With oOposSlip
    '*KT 51-09-25 Old การพิมพ์ Logo เก่า
    '                 .FeedCut
    '                .TransactionPrint PTR_S_RECEIPT, PTR_TP_NORMAL 'end print by OPOS
    '                 PrintPicture poPicture, 2.5, 0, vbCenter
'                    If pbPrnLogo Then
'                        If bVB_ImageLogo Then
'                            .Printnormal PTR_S_SLIP, "" & LF
'                             Call PrintLogo
'                        Else '*KT 51-10-14 V 3.0.135/V 3.0.200 กรณีที่ไม่มีการ Set Logo
'                             '-----------------------------------------------------------------------
'                             For nI = 1 To 8
'                                .Printnormal PTR_S_SLIP, "" & LF
'                             Next nI
'                             '----------------------------------------------------------------------
'                        End If
'                    End If

'                    For nI = nC_LineNo To 35                                               'Ao 2013/07/25    ไม่ต้องพิมพ์ ช่องว่าง
'                        .Printnormal PTR_S_SLIP, "" & LF
'                    Next nI
                    .TransactionPrint PTR_S_SLIP, PTR_TP_NORMAL 'end print by OPOS
'                    .Printnormal PTR_S_SLIP, ESC & Chr(101)
'                    If pbPrnLogo Then
'                        .CutPaper 90
'                    End If
'                    .Printnormal PTR_S_SLIP, DLE & Chr(64)
'                    .Printnormal PTR_S_SLIP, DLE & DC4 '*KT 51-12-22 Phase 2 Clear Buffer
'                    oOposSlip.ClearOutput
'                    Call SP_CHKbSlpRemoval                     '*Ao 2013/07/25
    '---------------------------------------------------------------
'                    If pbDrawer And bVB_OPOSCashDrawerEnable Then              '*Ao 2013/08/28   เช็คไม่ได้ ให้ไปเช็คหลังพิมพ์ Slip
'                        If .DrawerOpen Then wMsgDrawer.Show vbModal ': DoEvents *KT 53-02-03 เอาออกไม่มีผลต่อการแสดงหน้าจอ CashDrawer
'                    End If
                    If tC_PthImageTmp <> "" Then
                        If SP_FLEbManFile(tC_PthImageTmp, nEN_Fle1Exist) Then
                            Call Kill(tC_PthImageTmp)
                        End If
                    End If
                End With
            Else
                If bVB_CNPrnOPOS Then Call SP_SHWxErrPrinter(False, "1")
                With oOpos
    '*KT 51-09-25 Old การพิมพ์ Logo เก่า
    '                 .FeedCut
    '                .TransactionPrint PTR_S_RECEIPT, PTR_TP_NORMAL 'end print by OPOS
    '                 PrintPicture poPicture, 2.5, 0, vbCenter
                    '*Ao 2013/01/24   Print FRONTTECH
'                    If pbPrnLogo Then
'                        If bVB_ImageLogo Then
'                            .Printnormal PTR_S_RECEIPT, "" & LF
'                             Call PrintLogo
'                        Else '*KT 51-10-14 V 3.0.135/V 3.0.200 กรณีที่ไม่มีการ Set Logo
'                             '-----------------------------------------------------------------------
'                             For nI = 1 To 8
'                                .Printnormal PTR_S_RECEIPT, "" & LF
'                             Next nI
'                             '----------------------------------------------------------------------
'                        End If
'                    End If
                   .TransactionPrint PTR_S_RECEIPT, PTR_TP_NORMAL 'end print by OPOS
                    If pbPrnLogo Then
                        .CutPaper 90
                    End If
'                    .Printnormal PTR_S_RECEIPT, DLE & DC4 '*KT 51-12-22 Phase 2 Clear Buffer        '*Ao 2013/01/24   Print FRONTTECH
    '---------------------------------------------------------------
                    '*Eaw 56-11-19 ไม่ต้องเช็ค Cash Drawer
'                    If pbDrawer And bVB_OPOSCashDrawerEnable Then
'                        If .DrawerOpen Then
'                            wMsgDrawer.Show vbModal ': DoEvents *KT 53-02-03 เอาออกไม่มีผลต่อการแสดงหน้าจอ CashDrawer
'                            DoEvents
'                        End If
'                    End If
                    If tC_PthImageTmp <> "" Then
                        If SP_FLEbManFile(tC_PthImageTmp, nEN_Fle1Exist) Then
                            Call Kill(tC_PthImageTmp)
                        End If
                    End If
                End With
            End If
        Else
             'driver
'             Select Case tUT_PrintType
'                    Case "2"
'                         PrintPicture poPicture, 2.5, cW_CurY, vbCenter '*KT 54-11-10 V.4.0.3 แก้ไขการพิมพ์ Slip ด้วย Driver
'                         tPrint = Chr(27) & Chr(105) & vbCr '*KT 54-11-10 V.4.0.3 สั่งตัดกระดาษ
'                         iRet = WritePrinter(nUT_HdlePrinter, ByVal tPrint, Len(tPrint), ipcWritten)
'                        Call UT_PrnClosed
'                    Case "3"
'                        'PrintPicture poPicture, 2.5, 0, vbCenter '*KT 54-11-10 V.4.0.3 แก้ไขการพิมพ์ Slip ด้วย Driver
                        oPrn.EndDoc
'            End Select

            CurrentY = 0  '*KT 54-11-10 V.4.0.3 แก้ไขการพิมพ์ Slip ด้วย Driver
            '*Eaw 56-11-19 ไม่ต้องเช็ค Cash Drawer
'            'CashDrawer
'            If pbDrawer And bVB_OPOSCashDrawerEnable Then
'                If oOpos.DrawerOpen Then wMsgDrawer.Show vbModal ': DoEvents *KT 53-02-03 เอาออกไม่มีผลต่อการแสดงหน้าจอ CashDrawer
'            End If
        End If
'return
EndDoc = True

Exit Function '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
ErrHandler:

End Function

Private Function W_GETtLTrim(ByVal ptText As String, cWidth As Double) As String
Dim nTmp As Long, tTmp As String
On Error GoTo ErrHandler
            
            For nTmp = Len(ptText) To 1 Step -1
                    tTmp = Left(ptText, nTmp)
                    If oPrn.TextWidth(tTmp) <= cWidth Then Exit For
            Next
    
'return
W_GETtLTrim = tTmp

Exit Function '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
ErrHandler:

End Function

Public Property Set ActivePrinter(poPrinter As Printer)
        Set oPrn = poPrinter
End Property

Public Property Get ActivePrinter() As Printer
        Set ActivePrinter = oPrn
End Property

Public Property Let LineNo(pnLine As Long)
       nC_LineNo = pnLine
End Property

Public Property Get LineNo() As Long
        LineNo = nC_LineNo
End Property

Public Property Get LineNoEx() As Long
        LineNoEx = nC_LineNoEx
End Property

Public Property Let LineMax(pnLine As Long)
       nC_LineMax = pnLine
End Property

Public Property Get LineMax() As Long
        LineMax = nC_LineMax
End Property

Public Property Let PageNo(pnPage As Long)
       nC_PageNo = pnPage
End Property

Public Property Get PageNo() As Long
        PageNo = nC_PageNo
End Property

'
Public Property Let PrinterType(ptPrinterType As String)
       tC_PrinterType = ptPrinterType
End Property

Public Property Get PrinterType() As String
        PrinterType = tC_PrinterType
End Property

Public Sub LineHeader(ByVal pnIndex As Long, ByVal ptText As String, ByVal pcPosition As Double, ByVal pnAli As Long, ByVal pbLF As Boolean)
       oC_LineHeader(pnIndex).tText = ptText
       oC_LineHeader(pnIndex).cPosition = pcPosition
       oC_LineHeader(pnIndex).nAli = pnAli
       oC_LineHeader(pnIndex).bLF = pbLF
End Sub

Private Sub W_SHWxErrPrinter()
'------------------------------------
'Call:
'Cmt: แสดงหน้าจอ error result  ของ Printer
'------------------------------------
    With oOpos
        Select Case .ResultCode
            Case OPOS_E_EXTENDED
                Select Case .ResultCodeExtended
                    Case OPOS_EPTR_COVER_OPEN
                        wCNMsgErr.Show vbModal
                    Case OPOS_EPTR_REC_EMPTY
                        wCNMsgErr.Show vbModal
                End Select
        End Select
    End With
End Sub

'Public Sub UT_PrnOpen(ByVal ptDevice$)
''----------------------------------------------------------
''   Call:   ptDevice is device name of printer
''
''   Cmt:   open printer for line by line
''*KT 54-11-10 V.4.0.3 แก้ไขกรณีใช้ Driver แบบ รวมพิมพ์โดยตรง
''----------------------------------------------------------
'    Dim tPrint$, ipcWritten&
'    Dim iRet&, iDoc&, tDevice$
'    Dim MyDocInfo As DOCINFO
'
'    tDevice = Trim(ptDevice)
'
'    If tDevice = "" Then tDevice = Printer.DeviceName       '*GP 47-11-15
'    iRet = OpenPrinter(tDevice, nUT_HdlePrinter, 0)
'
'    If iRet = 0 Then        '*GP เช็คแค่ว่า Open ได้หรือไม่, ไม่สามารถเช็คได้ว่าเครื่องพิมพ์ เปิดหรือปิด
'        Call SP_MSGnShowing(tCS_MSG59, nCS_Warning)
'        Exit Sub
'    End If
'
'    bUT_PrinterOk = True        'printer open ok or normal
'
'    MyDocInfo.pDocName = "POS Printing"
'    MyDocInfo.pOutputFile = vbNullString
'    MyDocInfo.pDatatype = vbNullString
'    iDoc = StartDocPrinter(nUT_HdlePrinter, 1, MyDocInfo)
'    Call StartPagePrinter(nUT_HdlePrinter)
'End Sub
'
'Public Sub UT_PrnClosed()
''----------------------------------------------------------
''   Call:
''
''   Cmt:   just closed printer
''----------------------------------------------------------
'    Dim iRet&
'
'    If Trim(tUT_Printer) = "" Then       'no printer, exit
'        bUT_PrinterOk = False           'closed printer ok
'        Exit Sub
'    End If
'
'    If Not bUT_PrinterOk Then Exit Sub          'printer closed or open abnormal
'
'    Select Case UCase(tUT_Printer)
'        Case "TMU210", "TMU300", "TMU950", "IDP3541", "FUJITSUF01"       '*GPxxx 46-05-11 Fujitsu
'            iRet = ClosePrinter(nUT_HdlePrinter)
'            iRet = EndPagePrinter(nUT_HdlePrinter)
'            iRet = EndDocPrinter(nUT_HdlePrinter)
'        Case "IBMSUREONE"    'IBMSureOne
'
'
'    End Select
'
'    bUT_PrinterOk = False       'closed printer ok
'
'End Sub
'
'
'Public Sub UT_ThaJustAlig(ByVal ptPrint$, Optional ByVal piAlig%)
''----------------------------------------------------------
''   Call:   ptPrint is wording that want to print
''
''   Cmt:   justify thai wording to three levels
''               for all printer
''----------------------------------------------------------
'    Dim tPrint$, iRet%, ipcWritten&
'    Dim i%, iASC%, iASCp%
'    Dim tText$, bJust As Boolean
'    Dim bTFill As Boolean, bBFill As Boolean
'    Dim t1$, t2$, t3$, iP%, tBlank$
'    Dim iT1%, iT3%, iDiv%, iPdtW%
'    Dim tTemp$, tPrice$, tVat$
'    Dim nPMax%              'จำนวนอักษรที่ใส่ได้มากสุด ของกระดาษ
'    Dim nSaraUpDwn%
'    Dim nThaCntSara%      '*GPxxx 46-05-11 Fujitsu
'
'    On Error GoTo ErrHandle
'
'    nPMax = 40
'    'iPdtW = UT_iProductWidth()  '*GP 47-03-06
'    iPdtW = 25          '*GP 47-03-06
'    tBlank = Chr(32)     'ให้ค่าว่าง
'
'    tTemp = ptPrint
'    tText = UT_tToken(tTemp, ";")        'ให้ค่าตั้งต้นก่อน
'    nSaraUpDwn = UT_nThaAbsoluteCount(tText)        '*Tee 53-03-31 SFM ComSheet-BP-P1-105-Tee
'    If Trim(tTemp) <> "" Then
'        tPrice = UT_tToken(tTemp, ";")
'        If Trim(tTemp) <> "" Then tVat = UT_tToken(tTemp, ";")
'    End If
'
'    Select Case tUT_Printer         '*GPxxx 46-05-11 Fujitsu
'        Case "IDP3541", "FUJITSUF01"        'ไม่ต้องจัดระดับภาษาไทย
'             'aligment
'             If Len(tText) > nPMax + nSaraUpDwn Then tText = Left(tText, nPMax + nSaraUpDwn)      'ความกว้างจะต้องไม่เกิน nPMax
'             Select Case piAlig
'                 Case 0      'left nothing to do
'
'                 Case 1      'center
'                     iDiv = Round((nPMax - Len(tText)) / 2, 0)
'                     tText = Space(iDiv) & tText & Space(nPMax - (iDiv + Len(tText)))
'
'                 Case 2      'right justify
'                     tText = Space(nPMax - Len(tText)) & tText
'
'             End Select
'             If tPrice <> "" Then                            'จัดระดับ หรือไม่จัดระดับ ภาษาไทยก็ยังมีผลต่อการพิมพ์ราคาสินค้า เพราะ Len(ไทย-มีสระ) <> Len(ไทย-ไมีมีสระ)
'                 'tText = Left(tText, iPdtW - 3)       'set iPdtW'*TW 50-03-22 เอาออกเลย
'
'                 nThaCntSara = UT_nThaAbsoluteCount(tText)       'ความกว้างชื่อสินค้า -  นับสระภาษาไทย
'                 If iPdtW < (Len(tText) - nThaCntSara) Then tText = Left(tText, Len(tText) - (Len(tText) - nThaCntSara - iPdtW))
'                        'tText = tText & " " & String(13 - Len(Format(Trim(tPrice), "#,###,##0.00")), " ") & Format(Trim(tPrice), "#,###,##0.00") & " " & tVat
'                 ' Else
'                 If Len(Format(Trim(tPrice), "#,###,##0.00")) > 13 Then '*TW 50-06-06
'                        tText = tText & String(iPdtW - (Len(tText) - nThaCntSara), " ") & vbCrLf & String(40 - Len(Format(Trim(tPrice), "#,###,##0.00") & " " & tVat), " ") & Format(Trim(tPrice), "#,###,##0.00") & " " & tVat
'                 Else
'                        tText = tText & String(iPdtW - (Len(tText) - nThaCntSara), " ") & String(13 - Len(Format(Trim(tPrice), "#,###,##0.00")), " ") & Format(Trim(tPrice), "#,###,##0.00") & " " & tVat
'                End If
'                 'End If
'
'             End If
'
'             tText = tText & vbCrLf
'             iRet = WritePrinter(nUT_HdlePrinter, ByVal tText, Len(tText), ipcWritten)
'             bUT_NotJust = True               'เก็บว่าไม่มีการ justify จริง      'ไม่มีผลอะไร ขึ้นอยู่กับเครื่องพิมพ์
'
'        Case Else                   'จัดระดับภาษาไทย
'            'เช็คก่อนว่าต้องตัดคำหรือไม่ less process
'            For i = 1 To Len(tText)
'                iASC = Asc(Mid(tText, i, 1))
'                Select Case iASC
'                    Case 209, 212 To 218, 232 To 237            'all สระบน, ล่าง
'                        bJust = True
'                        Exit For
'                    Case Else
'                End Select
'            Next i
'
'            If bJust Then           'ถ้ามีการตัดคำ
'                iP = 0
'                For i = 1 To Len(tText)
'                    iASC = Asc(Mid(tText, i, 1))
'                    Select Case iASC
'                        Case 212 To 215, 209    'สระอิ[  ิ  ], สระอี[  ี  ], สระอึ[  ึ  ], สระอือ[  ื  ], ไม้หันอากาศ[  ั  ]
'                            t1 = t1 & Chr(iASC)
'                            iT1 = iT1 + 1
'                            bTFill = True
'                            iP = 1
'
'                        Case 231                            'ไม้ไตคู้[  ็  ]
'                            t1 = t1 & Chr(iASC)
'                            bTFill = True
'                            iT1 = iT1 + 1
'                            iP = 2
'
'                        Case 236                            'ไม้การันต์[  ์ ]
'                            If (iP = 1) And (iASCp = 212) Then
'                                'ถ้าก่อนหน้านี้เป็นสระอิ ต้องลบสระอิก่อนหน้านี้ก่อน แล้วเปลี่ยนเป็น  ->  ิ์
'                                t1 = Left(t1, Len(t1) - 1)
'                                iT1 = iT1 - 1
'                                t1 = t1 & Chr(UT_iThaConvertSaraIKaran)            'call function get ค่า Ascii   ิ์
'                            Else
'                                t1 = t1 & Chr(iASC)
'                            End If
'                            bTFill = True
'                            iT1 = iT1 + 1
'                            iP = 3
'
'                        Case 232 To 235             'ไม้เอก[  ่ ], ไม้โท[  ้  ], ไม้ตรี[  ๊  ], ไม้จัตวา[  ๋ ]
'                            If iP = 1 Then
'                                'ถ้าก่อนหน้านี้เป็น ชุดที่ 1 ต้องลบตัวก่อนหน้านี้ก่อนเปลี่ยนให้เป็น -> สระติดกัน แยกตาม case ตัวเดิม
'                                t1 = Left(t1, Len(t1) - 1)
'                                iT1 = iT1 - 1
'                                t1 = t1 & Chr(UT_iThaConvertSara(iASC, iASCp))          'call function get ค่า Ascii ชุด สระ อิ, อี, อึ, อือ
'                            Else
'                                t1 = t1 & Chr(iASC)
'                            End If
'                            bTFill = True
'                            iT1 = iT1 + 1
'                            iP = 4
'
'                        Case 216 To 217             'สระอุ[  ุ  ] กับ สระอู[  ู  ]
'                            t3 = t3 & Chr(iASC)
'                            bBFill = True
'                            iT3 = iT3 + 1
'                            iP = 5
'
'                        Case Else                          'ตัวอักขระอื่นๆ
'                            t2 = t2 & Chr(iASC)
'                            Select Case iP
'                                Case 1, 2, 3, 4, 5
'                                    If (bTFill = False) Then t1 = t1 & tBlank
'                                    If (bBFill = False) Then t3 = t3 & tBlank
'                                Case 6
'                                    t1 = t1 & tBlank
'                                    t3 = t3 & tBlank
'                            End Select
'                            bBFill = False
'                            bTFill = False
'                            iP = 6
'                        End Select
'                        iASCp = iASC                    'เก็บค่าตัวเก่าไว้
'                Next i
'
'                If (tPrice <> "") And (IsNumeric(tPrice)) Then
'                    t1 = Left(t1, iPdtW - 3)             'set iPdtW
'
'                    t2 = t2 & Space(iT1 + iT3)      'บวกค่า T1 กับ T3
'                    t2 = Left(t2, iPdtW - 3)               'set iPdtW           '*GPxxx ไม่แน่ใจว่าจะต้อง fixed=0.00 หรือเปล่า เพราะกระดาษมีความกว้างจำกัด
'                    t2 = t2 & String(iPdtW - Len(t2), " ") & String(13 - Len(SP_FMTcCurPoint(Trim(tPrice))), " ") & SP_FMTcCurPoint(Trim(tPrice)) & " " & tVat
'
'                    t3 = Left(t3, iPdtW - 3)
'                End If
'
'                'aligment
'                If Len(t1) > nPMax Then t1 = Left(t1, nPMax)           'สำหรับเครื่องพิมพ์ TMU ความกว้างจะต้องไม่เกิน nPMax
'                If Len(t2) > nPMax Then t2 = Left(t2, nPMax)
'                If Len(t3) > nPMax Then t3 = Left(t3, nPMax)
'                Select Case piAlig
'                    Case 0      'left nothing to do
'
'                    Case 1      'center
'                        iDiv = Round((nPMax - Len(t2)) / 2, 0)
'                        t1 = Space(iDiv) & t1 & Space(nPMax - (iDiv + Len(t1)))
'                        t2 = Space(iDiv) & t2 & Space(nPMax - (iDiv + Len(t2)))
'                        t3 = Space(iDiv) & t3 & Space(nPMax - (iDiv + Len(t3)))
'
'                    Case 2      'right justify
'                        t1 = Space(nPMax - (Len(t1) + 1)) & t1
'                        t2 = Space(nPMax - Len(t2)) & t2
'                        t3 = Space(nPMax - (Len(t3) + 1)) & t3
'
'                End Select
'
'                'print t1, check t1 ถ้าเป็น  blank ทั้งหมดจะไม่พิมพ์ เพราะบรรทัดจะถูกดันให้กว้าง
'                bUT_T1Print = False
'                For i = 1 To Len(t1)
'                    If Mid(t1, i, 1) <> tBlank Then
'                        tText = t1 & vbCrLf
'                        iRet = WritePrinter(nUT_HdlePrinter, ByVal tText, Len(tText), ipcWritten)
'                        bUT_T1Print = True
'                        Exit For
'                    End If
'                Next i
'
'                'ถ้าก่อนหน้านี้ไม่พิมพ์ t3 และตัวปัจจุบันก็ไม่มีพิมพ์ t1 ต้องพิมพ์ blank
'                If (Not bUT_T1Print) And (Not bUT_T3Print) Then
'                    tPrint = tBlank & vbCrLf              'พิมพ์  blank
'                    iRet = WritePrinter(nUT_HdlePrinter, ByVal tPrint, Len(tPrint), ipcWritten)
'                End If
'
'                'print t2
'                tText = t2 & vbCrLf
'                iRet = WritePrinter(nUT_HdlePrinter, ByVal tText, Len(tText), ipcWritten)
'                'print t3, check t3 ถ้าเป็น  blank ทั้งหมดจะไม่พิมพ์ เพราะบรรทัดจะถูกดันให้กว้าง
'                bUT_T3Print = False
'                For i = 1 To Len(t3)
'                    If Mid(t3, i, 1) <> tBlank Then
'                        tText = t3 & vbCrLf
'                        iRet = WritePrinter(nUT_HdlePrinter, ByVal tText, Len(tText), ipcWritten)
'                        bUT_T3Print = True           'เช็คว่ามีการพิมพ์ t3 จริง
'                        Exit For
'                    End If
'                Next i
'                bUT_NotJust = False             'เก็บว่ามีการ justify
'            Else
'                If tUT_Printer <> "IDP3541" Then        '*GP idp
'                    If (Not bUT_T3Print) Or bUT_NotJust Then              'เช็คว่าก่อนหน้านี้พิมพ์ t3? กับ ก่อนหน้านี้ justified? (แก้เรื่อง line spacing)
'                        tPrint = tBlank & vbCrLf              'พิมพ์  blank
'                        iRet = WritePrinter(nUT_HdlePrinter, ByVal tPrint, Len(tPrint), ipcWritten)
'                    End If
'                End If
'
'                'aligment
'                If Len(tText) > nPMax Then tText = Left(tText, nPMax)           'สำหรับเครื่องพิมพ์ TMU ความกว้างจะต้องไม่เกิน nPMax
'                Select Case piAlig
'                    Case 0      'left nothing to do
'
'                    Case 1      'center
'                        iDiv = Round((nPMax - Len(tText)) / 2, 0)
'                        tText = Space(iDiv) & tText & Space(nPMax - (iDiv + Len(tText)))
'
'                    Case 2      'right justify
'                        tText = Space(nPMax - Len(tText)) & tText
'
'                End Select
'
'                If tPrice <> "" Then
'                    tText = Left(tText, iPdtW - 3)       'set iPdtW
'                    tText = tText & String(iPdtW - Len(tText), " ") & String(13 - Len(SP_FMTcCurPoint(Trim(tPrice), False)), " ") & SP_FMTcCurPoint(Trim(tPrice), False) & " " & tVat
'                End If
'
'                tText = tText & vbCrLf
'                iRet = WritePrinter(nUT_HdlePrinter, ByVal tText, Len(tText), ipcWritten)
'
'                bUT_NotJust = True               'เก็บว่าไม่มีการ justify จริง
'
'            End If
'
'    End Select
'
'    Exit Sub
'ErrHandle:
'    'Call SP_MSGShowError("mUT.UT_ThaJustAlig")
'End Sub
'
'
'Public Function UT_nThaAbsoluteCount(ByVal ptPrint$) As Integer
''----------------------------------------------------------
''   Call:   ptPrint is wording that want to print
''
''   Cmt:   justify thai wording to three levels
''               for all printer
''----------------------------------------------------------
'    Dim tPrint$, iRet%, ipcWritten&
'    Dim i%, iASC%, iASCp%
'    Dim tText$, bJust As Boolean
'    Dim bTFill As Boolean, bBFill As Boolean
'    Dim t1$, t2$, t3$, iP%, tBlank$
'    Dim iT1%, iT3%, iDiv%, iPdtW%
'    Dim tTemp$, tPrice$, tVat$
'
'    UT_nThaAbsoluteCount = 0
'    'iPdtW = UT_iProductWidth()
'
'    tBlank = Chr(32)     'ให้ค่าว่าง
'
'    tTemp = ptPrint
'    tText = UT_tToken(tTemp, ";")        'ให้ค่าตั้งต้นก่อน
'
''    If Trim(tTemp) <> "" Then
''        tPrice = UT_tToken(tTemp, ";")
''        If Trim(tTemp) <> "" Then tVat = UT_tToken(tTemp, ";")
''    End If
'
'    'เช็คก่อนว่าต้องตัดคำหรือไม่ less process
'    For i = 1 To Len(tText)
'        iASC = Asc(Mid(tText, i, 1))
'        Select Case iASC
'                Case 209, 212 To 218, 231, 232 To 237           'all สระบน, ล่าง           '*Ao 51-08-13 เพิ่ม 231   'ไม้ไตคู้[  ็  ]
'                bJust = True
'                Exit For
'            Case Else
'        End Select
'    Next i
'
'    If bJust Then           'ถ้ามีการตัดคำ
'        iP = 0
'        For i = 1 To Len(tText)
'            iASC = Asc(Mid(tText, i, 1))
'            Select Case iASC
'                Case 212 To 215, 209    'สระอิ[  ิ  ], สระอี[  ี  ], สระอึ[  ึ  ], สระอือ[  ื  ], ไม้หันอากาศ[  ั  ]
'                    t1 = t1 & Chr(iASC)
'                    iT1 = iT1 + 1
'                    bTFill = True
'                    iP = 1
'
'                Case 231                            'ไม้ไตคู้[  ็  ]
'                    t1 = t1 & Chr(iASC)
'                    bTFill = True
'                    iT1 = iT1 + 1
'                    iP = 2
'
'                Case 236                            'ไม้การันต์[  ์ ]
'                    If (iP = 1) And (iASCp = 212) Then
'                        'ถ้าก่อนหน้านี้เป็นสระอิ ต้องลบสระอิก่อนหน้านี้ก่อน แล้วเปลี่ยนเป็น  ->  ิ์
'                        't1 = Left(t1, Len(t1) - 1)
'                        'iT1 = iT1 - 1
'                        't1 = t1 & Chr(UT_iThaConvertSaraIKaran)            'call function get ค่า Ascii   ิ์
'                        t1 = t1 & Chr(iASC)
'                    Else
'                        t1 = t1 & Chr(iASC)
'                    End If
'                    bTFill = True
'                    iT1 = iT1 + 1
'                    iP = 3
'
'                Case 232 To 235             'ไม้เอก[  ่ ], ไม้โท[  ้  ], ไม้ตรี[  ๊  ], ไม้จัตวา[  ๋ ]
'                    If iP = 1 Then
'                        'ถ้าก่อนหน้านี้เป็น ชุดที่ 1 ต้องลบตัวก่อนหน้านี้ก่อนเปลี่ยนให้เป็น -> สระติดกัน แยกตาม case ตัวเดิม
'                        't1 = Left(t1, Len(t1) - 1)
'                        'iT1 = iT1 - 1
'                        't1 = t1 & Chr(UT_iThaConvertSara(iASC, iASCp))          'call function get ค่า Ascii ชุด สระ อิ, อี, อึ, อือ
'                        t1 = t1 & Chr(iASC)
'                    Else
'                        t1 = t1 & Chr(iASC)
'                    End If
'                    bTFill = True
'                    iT1 = iT1 + 1
'                    iP = 4
'
'                Case 216 To 217             'สระอุ[  ุ  ] กับ สระอู[  ู  ]
'                    t3 = t3 & Chr(iASC)
'                    bBFill = True
'                    iT3 = iT3 + 1
'                    iP = 5
'
'                Case Else                          'ตัวอักขระอื่นๆ
'                    t2 = t2 & Chr(iASC)
'                    Select Case iP
'                        Case 1, 2, 3, 4, 5
'                            If (bTFill = False) Then t1 = t1 & tBlank
'                            If (bBFill = False) Then t3 = t3 & tBlank
'                        Case 6
'                            t1 = t1 & tBlank
'                            t3 = t3 & tBlank
'                    End Select
'                    bBFill = False
'                    bTFill = False
'                    iP = 6
'                End Select
'                iASCp = iASC                    'เก็บค่าตัวเก่าไว้
'        Next i
'        UT_nThaAbsoluteCount = iT1 + iT3
'    Else
'        UT_nThaAbsoluteCount = 0
'    End If
'
'End Function
'
'Public Function UT_iThaConvertSaraIKaran() As Integer
''----------------------------------------------------------
''   Call:
''
''   Ret:    คืนค่า Ascii code สำหรับสระอิ + การันต์  [  ิ์  ]
''----------------------------------------------------------
'
'    'default ของ TMU300, TMU950
'    UT_iThaConvertSaraIKaran = 154
'
'    Select Case UCase(tUT_Printer)
'        Case "TMU210"          'Set EpsonIR310 -TMU210 Dot-Matrix
'            UT_iThaConvertSaraIKaran = 136
'
'        Case "IBMSUREONE"          'Set IBM
'
'        Case "FUJITSUF01"          'Set Fujitsu
'
'        Case "TMU300"          'TMU300
'            'get default
'
'        Case "TMU950"          'TMU950
'            'get default
'
'    End Select
'End Function
'
'Public Function UT_iThaConvertSara(ByVal piASC%, ByVal piASCp%) As Integer
''----------------------------------------------------------
''   Call:   piASC is current Ascii
''               piASCp is previous Ascii
''
''   Ret:    คืนค่า Ascii code สำหรับสระอิ, อี, อึ, อือ
''----------------------------------------------------------
'    Dim iSara1%, iSara2%, iSara3%, iSara4%, iSara5%
'
'    'set default ที่ TMU300, TMU950
'    iSara1 = 82
'    iSara2 = 77
'    iSara3 = 13
'    iSara4 = -19
'    iSara5 = 86
'
'    Select Case UCase(tUT_Printer)
'        Case "TMU210"          'Set EpsonIR310 -TMU210 Dot-Matrix
'            iSara1 = 100
'            iSara2 = 95
'            iSara3 = 91
'            iSara4 = 87
'            iSara5 = 104
'
'        Case "IBMSUREONE"          'Set IBM
'
'        Case "FUJITSUF01"          'Set Fujitsu
'
'        Case "TMU300"          'TMU300
'            'get default
'
'        Case "TMU950"          'TMU950
'            'get default
'
'    End Select
'
'    Select Case piASCp           'เทียบกับค่าก่อนหน้านี้
'        Case 212
'            UT_iThaConvertSara = piASC - iSara1            '  ิ่  ,  ิ้  ,  ิ๊  ,  ิ๋         ชุดสระอิ
'        Case 213
'            UT_iThaConvertSara = piASC - iSara2            '  ี่  ,  ี้  ,  ี๊  , ี๋          ชุดสระอี
'        Case 214
'            UT_iThaConvertSara = piASC - iSara3            '  ึ่  ,  ึ้  ,  ึ๊  ,  ึ๋         ชุดสระอื
'        Case 215
'            UT_iThaConvertSara = piASC - iSara4           '  ื่  ,  ื้  ,  ื๊  ,  ื๋         ชุดสระอือ
'        Case 209
'            UT_iThaConvertSara = piASC - iSara5            '  ั่  ,  ั้  ,  ั๊  ,  ั๋         ชุดไม้หันอากาศ
'    End Select
'End Function
