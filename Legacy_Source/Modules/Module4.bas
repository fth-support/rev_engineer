Attribute VB_Name = "Module4"
Option Explicit
Public Declare Function SleepEx Lib "kernel32" (ByVal dwMilliseconds As Long, ByVal bAlertable As Long) As Long

'Public Declare Function WSACleanup Lib "wsock32.dll" () As Long
'Public Declare Function IcmpCreateFile Lib "Icmp.dll" () As Long
'Public Declare Function IcmpCloseHandle Lib "Icmp.dll" (ByVal IcmpHandle As Long) As Long
'Public Declare Sub RtlMoveMemory Lib "kernel32" (hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)
'Public Declare Function WSAStartup Lib "wsock32.dll" (ByVal wVersionRequired As Long, lpWSAData As WSAdata) As Long
'Public Declare Function IcmpSendEcho Lib "Icmp.dll" (ByVal IcmpHandle As Long, ByVal DestinationAddress As Long, ByVal RequestData As String, ByVal RequestSize As Long, ByVal RequestOptions As Long, ReplyBuffer As ICMP_ECHO_REPLY, ByVal ReplySize As Long, ByVal TimeOut As Long) As Long
'
'
'Public Const PING_TIMEOUT = 200
'Public Const WS_VERSION_REQD = &H101
'Public Const WS_VERSION_MAJOR = WS_VERSION_REQD \ &H100 And &HFF&
'Public Const WS_VERSION_MINOR = WS_VERSION_REQD And &HFF&
'Public Const MIN_SOCKETS_REQD = 1
'Public Const MAX_WSADescription = 256
'Public Const MAX_WSASYSStatus = 128
''Winsock
'Declare Function gethostbyname& Lib "wsock32.dll" (ByVal hostname$)
'Public Type Hostent
'    hName As Long
'    hAliases As Long
'    hAddrType As Integer
'    hLength As Integer
'    hAddrList As Long
'End Type
'Public Type WSAdata
'    wVersion As Integer
'    wHighVersion As Integer
'    szDescription(0 To MAX_WSADescription) As Byte
'    szSystemStatus(0 To MAX_WSASYSStatus) As Byte
'    wMaxSockets As Long
'    wMaxUDPDG As Long
'    dwVendorInfo As Long
'End Type
'Public Type ICMP_ECHO_REPLY
'    Address         As Long
'    Status          As Long
'    RoundTripTime   As Long
'    DataSize        As Long  'formerly integer
'    DataPointer     As Long
'    'Options         As ICMP_OPTIONS
'    data            As String * 250
'End Type
'Public Type ICMP_OPTIONS
'    Ttl             As Byte
'    Tos             As Byte
'    Flags           As Byte
'    OptionsSize     As Byte
'    OptionsData     As Long
'End Type
''***
'Public Function AP_Ping(ByVal ptHostOrIP As String, Optional pnTimeOutmSec As Long = PING_TIMEOUT) As Boolean
'Dim echoValues As ICMP_ECHO_REPLY
'Dim pos As Integer
'Dim Count As Integer
'Dim returnIp As Collection
'
'On Error GoTo ErrHandle
'    If Trim(ptHostOrIP) = "" Then
'        AP_Ping = False
'        Exit Function
'    End If
'
'    If AP_SocketsInitialize() Then
'
'        If InStr(1, ptHostOrIP, ".", vbTextCompare) <> 0 Then
'            If IsNumeric(Mid(ptHostOrIP, 1, InStr(1, ptHostOrIP, ".") - 1)) = False Then
'                Set returnIp = AP_ResolveIpaddress(ptHostOrIP)
'                If returnIp.Count = 0 Then
'                    AP_Ping = False
'                    Exit Function
'                Else
'                    ptHostOrIP = returnIp.Item(1)
'                End If
'            End If
'        End If
'
'        'ping an ip address, passing the
'        'address and the ECHO structure
'        Call AP_PingAddress((ptHostOrIP), echoValues, pnTimeOutmSec)
'
'        If Left$(echoValues.data, 1) <> Chr$(0) Then
'           pos = InStr(echoValues.data, Chr$(0))
'           echoValues.data = Left$(echoValues.data, pos - 1)
'        Else
'              echoValues.data = ""
'        End If
'
'        AP_SocketsCleanup
'
'        If echoValues.Status <> 0 Then
'            AP_Ping = False
'        Else
'            AP_Ping = True
'        End If
'    End If
'    Exit Function
'ErrHandle:
'    AP_Ping = False
'End Function
'
'Private Function AP_PingAddress(szAddress As String, ECHO As ICMP_ECHO_REPLY, Optional TimeOut As Long = PING_TIMEOUT) As Long
'Dim hPort As Long
'Dim dwAddress As Long
'Dim sDataToSend As String
'Dim iOpt As Long
'
'   sDataToSend = "Echo This"
'   dwAddress = AP_AddressStringToLong(szAddress)
'
'   hPort = IcmpCreateFile()
'
'   If IcmpSendEcho(hPort, dwAddress, sDataToSend, Len(sDataToSend), 0, ECHO, Len(ECHO), TimeOut) Then
'        'the ping succeeded,
'        '.Status will be 0
'        '.RoundTripTime is the time in ms for
'        '               the ping to complete,
'        '.Data is the data returned (NULL terminated)
'        '.Address is the Ip address that actually replied
'        '.DataSize is the size of the string in .Data
'         AP_PingAddress = ECHO.RoundTripTime
'   Else
'        AP_PingAddress = ECHO.Status * -1
'   End If
'
'   Call IcmpCloseHandle(hPort)
'
'End Function
'
'
'Public Function AP_ResolveIpaddress(ByVal ptHostName As String) As Collection
''---------------------------------------
''*WD-SFM 49-03-02
''---------------------------------------
'Dim hostent_addr As Long
'Dim host As Hostent
'Dim hostip_addr As Long
'Dim temp_ip_address() As Byte
'Dim i As Integer
'Dim ip_address As String
'Dim Count As Integer
'
'    If AP_SocketsInitialize() Then
'
'        Set AP_ResolveIpaddress = New Collection
'        hostent_addr = gethostbyname(ptHostName)
'
'        If hostent_addr = 0 Then
'            Call AP_SocketsCleanup
'            Exit Function
'        End If
'
'        RtlMoveMemory host, hostent_addr, LenB(host)
'        RtlMoveMemory hostip_addr, host.hAddrList, 4
'
'        'get all of the IP address if machine is  multi-homed
'
'        Do
'            ReDim temp_ip_address(1 To host.hLength)
'            RtlMoveMemory temp_ip_address(1), hostip_addr, host.hLength
'
'            For i = 1 To host.hLength
'                ip_address = ip_address & temp_ip_address(i) & "."
'            Next
'            ip_address = Mid$(ip_address, 1, Len(ip_address) - 1)
'            AP_ResolveIpaddress.Add ip_address
'            ip_address = ""
'            host.hAddrList = host.hAddrList + LenB(host.hAddrList)
'            RtlMoveMemory hostip_addr, host.hAddrList, 4
'         Loop While (hostip_addr <> 0)
'
'        AP_SocketsCleanup
'    End If
'End Function
'
'Public Function AP_SocketsInitialize() As Boolean
'Dim WSAD As WSAdata
'Dim X As Integer
'Dim szLoByte As String
'Dim szHiByte As String
'Dim szBuf As String
'
'    X = WSAStartup(WS_VERSION_REQD, WSAD)
'
'   'check for valid response
'   'MsgBox "Windows Sockets for 32 bit Windows environments is not successfully responding."
'    If X <> 0 Then Exit Function
'
'   'check that the version of sockets is supported
'    If AP_LoByte(WSAD.wVersion) < WS_VERSION_MAJOR Or _
'       (AP_LoByte(WSAD.wVersion) = WS_VERSION_MAJOR And _
'        AP_HiByte(WSAD.wVersion) < WS_VERSION_MINOR) Then
'
'        szHiByte = Trim$(Str$(AP_HiByte(WSAD.wVersion)))
'        szLoByte = Trim$(Str$(AP_LoByte(WSAD.wVersion)))
'        szBuf = "Windows Sockets Version " & szLoByte & "." & szHiByte
'        szBuf = szBuf & " is not supported by Windows " & _
'                          "Sockets for 32 bit Windows environments."
'        'MsgBox szBuf, vbExclamation
'        Exit Function
'    End If
'
'   'check that there are available sockets
'    If WSAD.wMaxSockets < MIN_SOCKETS_REQD Then
'        szBuf = "This application requires a minimum of " & _
'                 Trim$(Str$(MIN_SOCKETS_REQD)) & " supported sockets."
'        'MsgBox szBuf, vbExclamation
'        Exit Function
'    End If
'
'    AP_SocketsInitialize = True
'
'End Function
'
'
'Private Sub AP_SocketsCleanup()
'Dim X As Long
'
'  'need to use a var (insread of embedding
'  'in the If..Then call) becuse the function
'  'returns the error code if failed.
'   X = WSACleanup()
'
'   If X <> 0 Then
'       MsgBox "Windows Sockets error " & Trim$(Str$(X)) & " occurred in Cleanup.", vbExclamation
'   End If
'
'End Sub
'
'Private Function AP_AddressStringToLong(ByVal tmp As String) As Long
'Dim i As Integer
'Dim parts(1 To 4) As String
'
'    i = 0
'
'    If InStr(1, tmp, ".", vbTextCompare) = 0 Then
'        AP_AddressStringToLong = gethostbyname(tmp)
'    Else
'        'we have to extract each part of the
'        '123.456.789.123 string, delimited by
'        'a period
'        While InStr(tmp, ".") > 0
'          i = i + 1
'          parts(i) = Mid(tmp, 1, InStr(tmp, ".") - 1)
'          tmp = Mid(tmp, InStr(tmp, ".") + 1)
'        Wend
'
'        i = i + 1
'        parts(i) = tmp
'
'        If i <> 4 Then
'          AP_AddressStringToLong = 0
'          Exit Function
'        End If
'
'        'build the long value out of the
'        'hex of the extracted strings
'        AP_AddressStringToLong = Val("&H" & Right("00" & Hex(parts(4)), 2) & Right("00" & Hex(parts(3)), 2) & Right("00" & Hex(parts(2)), 2) & Right("00" & Hex(parts(1)), 2))
'   End If
'End Function
'
'Private Function AP_LoByte(ByVal pnParam As Long) As Integer
'    AP_LoByte = pnParam And &HFF&
'End Function
'
'Private Function AP_HiByte(ByVal pnParam As Long) As Integer
'    AP_HiByte = pnParam \ &H100 And &HFF&
'End Function
'
