VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMVoucherApi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public oC_RESAuthMVoucherInterf As AdaMobileVoucher.cRESAuthInf
Public oC_RESVerifyMVoucherInterf As AdaMobileVoucher.cRESVerifyInf
Public oC_RESUpdateMVoucherInterf As AdaMobileVoucher.cRESUpdateInf

Private tC_AuthUrl As String
Private tC_VerifyUrl As String
Private tC_UpdateUrl As String
Private tC_client_id As String
Private tC_client_secret As String
Private tC_grant_type As String
Private tC_Scope As String
Private tC_ContentTypeAuth As String
Private tC_HKeyToken As String
Private nC_hosttimeout  As Integer
Private tC_VoucherNo As String
Private tC_POSID As String
Private tC_UpdateTransID As String
Private tC_Update_datetime As String

Private nC_ExpireIn As Long
Private dC_AuthDateTime As Date

Public Property Get nExpireIn() As Long
    nExpireIn = nC_ExpireIn
End Property

Public Property Get dAuthDateTime() As Date
    dAuthDateTime = dC_AuthDateTime
End Property

Property Get tPOSID() As String
    tVoucherNo = tC_POSID
End Property

Property Let tPOSID(New_tValue As String)
     tC_POSID = New_tValue
End Property

Property Get tVoucherNo() As String
    tVoucherNo = tC_VoucherNo
End Property
Property Let tVoucherNo(New_tValue As String)
    tC_VoucherNo = New_tValue
End Property

Property Let tUpdateTransID(New_tValue As String)
    tC_UpdateTransID = New_tValue
End Property

Property Let tUpdate_datetime(New_tValue As String)
    tC_Update_datetime = New_tValue
End Property

Property Let tAuthURL(New_tValue As String)
    tC_AuthUrl = New_tValue
End Property

Property Let tVerifyURL(New_tValue As String)
    tC_VerifyUrl = New_tValue
End Property

Property Let tUpdateURL(New_tValue As String)
    tC_UpdateUrl = New_tValue
End Property

Property Let tClient_id(New_tValue As String)
    tC_client_id = New_tValue
End Property

Property Let tClient_secret(New_tValue As String)
    tC_client_secret = New_tValue
End Property

Property Let tGrant_type(New_tValue As String)
    tC_grant_type = New_tValue
End Property

Property Let tScope(New_tValue As String)
    tC_Scope = New_tValue
End Property

Property Let tContentTypeAuth(New_tValue As String)
    tC_ContentTypeAuth = New_tValue
End Property

Property Let nHostTimeOut(New_nValue As Integer)
    nC_hosttimeout = New_nValue
End Property

'Property Get nHostTimeOut()
'    nHostTimeOut = nC_hosttimeout
'End Property

Public Property Get tHKeyToken()
    tHKeyToken = tC_HKeyToken
End Property
'
'Public Property Let tHKeyToken(New_tValue As String)
'    tC_HKeyToken = New_tValue
'End Property

Public Function C_APIbAuth() As Boolean
    Dim oREQInterf As New AdaMobileVoucher.cREQAuthInf
    Dim oCallPay As New AdaMobileVoucher.cVoucher
    Dim tAPIRes As String
    Dim tReq As String
    Dim nLoop As Long
On Error GoTo ErrHandle
    C_APIbAuth = False
   Set oC_RESAuthMVoucherInterf = New AdaMobileVoucher.cRESAuthInf
    'oC_RESAuthMVoucherInterf.GETtRespErrMsg = ""
    'Params
    'set Url
    oCallPay.tC_Url = Replace(Replace(tC_AuthUrl, Chr(13), ""), Chr(10), "")
    'TimeOut
     oCallPay.nC_TimeOut = nC_hosttimeout

    'Body
    oREQInterf.tClient_id = tC_client_id
    oREQInterf.tClient_secret = tC_client_secret
    oREQInterf.tGrant_type = tC_grant_type
    oREQInterf.tScope = tC_Scope
    oREQInterf.tContentType = tC_ContentTypeAuth
    
    'Header  Key
'    For nLoop = 0 To UBound(aVB_MVoucherHKey)
'        Dim oREQHKeyInterf As New AdaMobileVoucher.cREQHeaderInterf
'        Call oREQHKeyInterf.SETxValue(aVB_MVoucherHKey(nLoop).tHkeyName, aVB_MVoucherHKey(nLoop).tHkeyVal)
'        Call oREQInterf.ADDxHeader(oREQHKeyInterf.Get)
'    Next nLoop

    'Respone
    Set oC_RESAuthMVoucherInterf = oCallPay.SENDoAuthAPI(oREQInterf)
    'Req
    Call C_WRTxLog("Auth", False, oCallPay.tC_Url, oCallPay.tC_ReqJson)
    'Res
    Call C_WRTxLog("Auth", True, oCallPay.tC_Url, oCallPay.tC_ResJson)
    If oC_RESAuthMVoucherInterf.GETtRespAccessToken <> "" Then
        C_APIbAuth = True
        tC_HKeyToken = oC_RESAuthMVoucherInterf.GETtRespAccessToken
        nC_ExpireIn = oC_RESAuthMVoucherInterf.GETnExpireIn
        dC_AuthDateTime = Format(Now, "yyyy/MM/dd HH:mm:ss")
    Else
        C_APIbAuth = False
    End If
    Set oREQInterf = Nothing
    Set oCallPay = Nothing
    Exit Function
ErrHandle:
        C_APIbAuth = False
        If oC_RESAuthMVoucherInterf.GETtRespAccessToken = "" Then
            'oC_RESAuthMVoucherInterf.GEToRespError.GETnRespErrCode = Err.Number
            'oC_RESAuthMVoucherInterf.GEToRespError.GETtRespErrMsg = Err.Description
        End If
        Call C_WRTxLog("Auth", True, "", "", "Err No: " & Err.Number & "Err Desc: " & Err.Description)
        If oCallPay.tC_Url <> "" Or oCallPay.tC_ResJson <> "" Then
            Call C_WRTxLog("Auth", True, oCallPay.tC_Url, oCallPay.tC_ResJson)
        End If
        Set oREQInterf = Nothing
       Set oCallPay = Nothing
End Function

Public Function C_APIbVerify(ByRef pnErr As Long, ByRef ptErr As String) As Boolean
    Dim oREQInterf As New AdaMobileVoucher.cREQVerifyInf
    Dim oCallPay As New AdaMobileVoucher.cVoucher
    Dim tAPIRes As String
    Dim tReq As String
    Dim nLoop As Long
    Dim nLn As Integer
On Error GoTo ErrHandle
    nLn = 0
    C_APIbVerify = False
    Set oC_RESVerifyMVoucherInterf = New AdaMobileVoucher.cRESVerifyInf
    'oC_RESVerifyMVoucherInterf.GETtRespErrMsg = ""
    'Params
    'set Url
    nLn = 1
    oCallPay.tC_Url = Replace(Replace(tC_VerifyUrl, Chr(13), ""), Chr(10), "")
    'TimeOut
     oCallPay.nC_TimeOut = nC_hosttimeout

    'Body
    oREQInterf.tVoucherNo = tC_VoucherNo
    nLn = 2
    'Header  Key
    For nLoop = 0 To UBound(aVB_MVoucherHKey)
        Dim oREQHKeyInterf As New AdaMobileVoucher.cREQHeaderInterf
        Call oREQHKeyInterf.SETxValue(aVB_MVoucherHKey(nLoop).tHkeyName, aVB_MVoucherHKey(nLoop).tHkeyVal)
        Call oREQInterf.ADDxHeader(oREQHKeyInterf.Get)
    Next nLoop
    nLn = 3
    'Respone
    Set oC_RESVerifyMVoucherInterf = oCallPay.SENDoVerifAPI(oREQInterf)
    nLn = 4
    'Req
    Call C_WRTxLog("Verify", False, oCallPay.tC_Url, oCallPay.tC_ReqJson)
    'Res
    Call C_WRTxLog("Verify", True, oCallPay.tC_Url, oCallPay.tC_ResJson)
    If oC_RESVerifyMVoucherInterf.GETbRespSuccesss Then
    
        C_APIbVerify = True
    Else
        C_APIbVerify = False
    End If
    nLn = 5
    Set oREQInterf = Nothing
    Set oCallPay = Nothing
    nLn = 6
    Exit Function
ErrHandle:

        C_APIbVerify = False
        If Not oC_RESVerifyMVoucherInterf Is Nothing Then
            If nLn > 3 Then
                If oC_RESVerifyMVoucherInterf.GETtRespStatus = "" Then
        '            oC_RESVerifyMVoucherInterf.GETtRespStatus = "ADAERROR"
        '            oC_RESVerifyMVoucherInterf.GETnRespErrCode = Err.Number
        '            oC_RESVerifyMVoucherInterf.GETtRespErrMsg = Err.Description
                End If
                Call C_WRTxLog("Verify", True, "", "", "Err No: " & Err.Number & "Err Desc: " & Err.Description)
                If oCallPay.tC_Url <> "" Or oCallPay.tC_ResJson <> "" Then
                    Call C_WRTxLog("Verify", True, oCallPay.tC_Url, oCallPay.tC_ResJson)
                End If
            Else
                ptErr = Err.Description
                pnErr = Err.Number
            End If
        Else
            ptErr = Err.Description
            pnErr = Err.Number
        End If
        Set oREQInterf = Nothing
       Set oCallPay = Nothing
End Function

    Public Function C_APIbUpdate() As Boolean
    Dim oREQInterf As New AdaMobileVoucher.cREQUpdateInf
    Dim oCallPay As New AdaMobileVoucher.cVoucher
    Dim tAPIRes As String
    Dim tReq As String
    Dim nLoop As Long
On Error GoTo ErrHandle
    C_APIbUpdate = False
   Set oC_RESUpdateMVoucherInterf = New AdaMobileVoucher.cRESUpdateInf
    'oC_RESUpdateMVoucherInterf.GETtRespStatus = ""
    'Params
    'set Url
    oCallPay.tC_Url = Replace(Replace(tC_UpdateUrl, Chr(13), ""), Chr(10), "")
    'TimeOut
     oCallPay.nC_TimeOut = nC_hosttimeout

    'Body
    oREQInterf.tPOSID = tC_POSID
    oREQInterf.tVoucherNo = tC_VoucherNo
    oREQInterf.tUpdateTransID = tC_UpdateTransID
    oREQInterf.tUpdate_datetime = tC_Update_datetime
    
    'Header  Key
    For nLoop = 0 To UBound(aVB_MVoucherHKey)
        Dim oREQHKeyInterf As New AdaMobileVoucher.cREQHeaderInterf
        Call oREQHKeyInterf.SETxValue(aVB_MVoucherHKey(nLoop).tHkeyName, aVB_MVoucherHKey(nLoop).tHkeyVal)
        Call oREQInterf.ADDxHeader(oREQHKeyInterf.Get)
    Next nLoop

    'Respone
    Set oC_RESUpdateMVoucherInterf = oCallPay.SENDoUpdateAPI(oREQInterf)
    'Req
    Call C_WRTxLog("Update", False, oCallPay.tC_Url, oCallPay.tC_ReqJson)
    'Res
    Call C_WRTxLog("Update", True, oCallPay.tC_Url, oCallPay.tC_ResJson)
    If oC_RESUpdateMVoucherInterf.GETbRespSuccesss Then
         If oC_RESUpdateMVoucherInterf.GETnRespCode = 200 Then
            C_APIbUpdate = True
         End If
    Else
        C_APIbUpdate = False
    End If
    Set oREQInterf = Nothing
    Set oCallPay = Nothing
    Exit Function
ErrHandle:
        C_APIbUpdate = False
        If oC_RESUpdateMVoucherInterf.GETtRespStatus = "" Then
'            oC_RESUpdateMVoucherInterf.GETtRespStatus = "ADAERROR"
'            oC_RESUpdateMVoucherInterf.GEToRespError.GETnRespErrCode = Err.Number
'            oC_RESUpdateMVoucherInterf.GEToRespError.GETtRespErrMsg = Err.Description
        End If
        Call C_WRTxLog("Update", True, "", "", "Err No: " & Err.Number & "Err Desc: " & Err.Description)
        If oCallPay.tC_Url <> "" Or oCallPay.tC_ResJson <> "" Then
            Call C_WRTxLog("Update", True, oCallPay.tC_Url, oCallPay.tC_ResJson)
        End If
        Set oREQInterf = Nothing
       Set oCallPay = Nothing
End Function

Public Sub C_WRTxLog(ByVal ptAct As String, ByVal pbRest As Boolean, ByVal ptURL As String, ByVal ptJsonResp As String, Optional ByVal ptErr As String = "", Optional ByVal ptInfo As String = "")
'----------------------------------------------------------
' Call
' Cmt :   Create File Log
'----------------------------------------------------------
    Dim oLogFile As TextStream
    Dim oFSO  As New FileSystemObject
    Dim nLof As Long
    Dim tFileName$
    Dim tFld$
    Dim oFile As File
    Dim tMsg As String
On Error GoTo ErrHandler
        If ptInfo <> "" Then
             tMsg = "Info : " & ptInfo & vbNewLine
        ElseIf ptErr <> "" Then
             tMsg = "Error : " & ptErr & vbNewLine
        Else
             tMsg = "EVENT : " & ptAct & vbNewLine & _
                                                            "URL : " & ptURL & vbNewLine & _
                                                            "JSON " & IIf(pbRest, "RES", "REQ") & " : " & vbNewLine & ptJsonResp & vbNewLine
        End If
        tFileName = Left(App.Path, InStrRev(App.Path, "\") - 1) + "\LOG\" & "mVoucher-" & Format(Date, "yyyyMMdd") & ".Log"
        If Not SP_FLEbManFile(tFileName, nEN_Fle1Exist) Then
                tFileName = oFSO.GetParentFolderName(oFSO.GetParentFolderName(App.Path))
                tFileName = tFileName & "\LOG\mVoucher-" & Format(Date, "yyyyMMdd") & ".Log"
                If Not SP_FLEbManFile(tFileName, nEN_Fle1Exist) Then
                     tFileName = Left(App.Path, InStrRev(App.Path, "\") - 1) + "\LOG\" & "mVoucher-" & Format(Date, "yyyyMMdd") & ".Log"
                End If
        End If
        tFld = oFSO.GetParentFolderName(tFileName)
        If Not SP_FLDbManFolder(tFld, nEN_Fld1Exist) Then
           Call SP_FLDbManFolder(tFld, nEN_Fld3Create)
        End If
        If Not SP_FLEbManFile(tFileName, nEN_Fle1Exist) Then
            oFSO.CreateTextFile (tFileName)
            Set oFile = oFSO.GetFile(tFileName)
            oFile.Attributes = Normal
        Else
            Set oFile = oFSO.GetFile(tFileName)
            oFile.Attributes = Normal
        End If
        Set oLogFile = oFile.OpenAsTextStream(ForAppending)
        oLogFile.WriteLine "V." & App.Major & "." & App.Minor & "." & App.Revision & " Date:" & Date & " Time:" & Time$
        oLogFile.WriteLine tMsg
        oLogFile.Close
        Set oFSO = Nothing

        Open tFileName For Append As #1
        nLof = LOF(1)
        Seek #1, nLof + 1
        Print #1, "Date:" & Date & " Time:" & Time$ & " TerminalNum:" & tVB_CNTerminalNum
        Print #1, ptErr
        Close #1
        Set oFSO = Nothing
        Exit Sub '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
ErrHandler:
        Close #1
        Set oFSO = Nothing
End Sub

Public Function C_CHKbExpireToken(ByVal pdDate As Date, pnExpireIn As Long) As Boolean
    If DateDiff("S", Date, pdDate) >= (pnExpireIn - 5) Then
        C_CHKbExpireToken = True
    Else
        C_CHKbExpireToken = False
    End If
End Function
