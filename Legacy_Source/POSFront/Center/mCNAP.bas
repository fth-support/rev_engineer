Attribute VB_Name = "mCNAP"
Option Explicit
'*Ao 53-03-15
Public Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
Public Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
'Playsound returns immediately
'Do not use SND_SYNC
Public Const SND_ASYNC& = &H1

Public Declare Function GetForegroundWindow Lib "user32" () As Long

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Public Declare Sub ReleaseCapture Lib "user32" ()

'*WD 48-02-03 For File's Information

'***

'*WD Add 48-01-07
Public Const MAX_PATH = 260
Public Const FILE_ATTRIBUTE_NORMAL = &H80
Public Const SHGFI_SMALLICON = &H1
Public Const SHGFI_SYSICONINDEX = &H4000
Public Const SHGFI_USEFILEATTRIBUTES = &H10
Public Const SHGFI_TYPENAME = &H400
Public Const ILD_TRANSPARENT = &H1         ' display transparent
Public Const OPEN_EXISTING = 3
Public Type SHFILEINFO
  hIcon As Long                       '  out: icon
  iIcon As Long                       '  out: icon index
  dwAttributes As Long                '  out: SFGAO_ flags
  szDisplayName As String * MAX_PATH  '  out: display name (or path)
  szTypeName As String * 80           '  out: type name
End Type
'*

'********** API Type **********
'for pos printer
Public Type DOCINFO
    pDocName As String
    pOutputFile As String
    pDatatype As String
End Type

Public Type BROWSEINFO
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type

Public Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Public Type POINTAPI
    X As Long
    Y As Long
End Type

Public Const RAS95_MaxEntryName = 256
Public Const RAS95_MaxDeviceType = 16
Public Const RAS95_MaxDeviceName = 32

Public Type RASCONN95
    dwSize As Long
    hRasCon As Long
    szEntryName(RAS95_MaxEntryName) As Byte
    szDeviceType(RAS95_MaxDeviceType) As Byte
    szDeviceName(RAS95_MaxDeviceName) As Byte
End Type

Public Type RASCONNSTATUS95
    dwSize As Long
    RasConnState As Long
    dwError As Long
    szDeviceType(RAS95_MaxDeviceType) As Byte
    szDeviceName(RAS95_MaxDeviceName) As Byte
End Type

Public Const RAS_MAXDEVICETYPE = 16
Public Const RAS_MAXENTRYNAME As Integer = 256
Public Const RAS_MAXDEVICENAME As Integer = 128
Public Const RAS_RASCONNSIZE As Integer = 412

Public Type RASENTRYNAME95
    'set dwsize to 264
    dwSize As Long
    szEntryName(RAS95_MaxEntryName) As Byte
End Type

Public Type RasConn
    dwSize As Long
    hRasConn As Long
    szEntryName(RAS_MAXENTRYNAME) As Byte
    szDeviceType(RAS_MAXDEVICETYPE) As Byte
    szDeviceName(RAS_MAXDEVICENAME) As Byte
End Type

'********** API Constant **********
'constant
Public Const EWX_LOGOFF = 0
Public Const EWX_SHUTDOWN = 1
Public Const EWX_REBOOT = 2
Public Const EWX_FORCE = 4
Public Const SPI_SCREENSAVERRUNNING = 97
Public Const BIF_RETURNONLYFSDIRS = &H1
Public Const BIF_DONTGOBELOWDOMAIN As Long = &H2
Public Const BIF_RETURNFSANCESTORS As Long = &H8
Public Const BIF_BROWSEFORCOMPUTER As Long = &H1000
Public Const BIF_BROWSEFORPRINTER As Long = &H2000
Public Const BIF_BROWSEINCLUDEFILES As Long = &H4000
Public Const BIF_NEWDIALOGSTYLE As Long = &H40

Public Const SW_SHOWNORMAL = 1
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40

'-----------------------Start Regional Setting *WD Add 47-11-24----------------------------------------
'constant
Public Const LOCALE_SLANGUAGE As Long = &H2     'localized name of language
Public Const LOCALE_SSHORTDATE As Long = &H1F   'short date format string
Public Const LOCALE_SLONGDATE As Long = &H20    'long date format string
Public Const DATE_LONGDATE As Long = &H2
Public Const DATE_SHORTDATE As Long = &H1
Public Const HWND_BROADCAST As Long = &HFFFF&
Public Const WM_SETTINGCHANGE As Long = &H1A

Public Const LOCALE_ICALENDARTYPE As Long = &H1009
'Public Const LOCALE_SSHORTDATE As Long = &H1F
Public Const LOCALE_USER_DEFAULT As Long = &H400

'Public Const LOCALE_SLANGUAGE As Long = &H2     'localized name of language

'Calendar ID Values.
Public Const CAL_GREGORIAN = 1                      'Gregorian (localized) calendar
Public Const CAL_GREGORIAN_US = 2                   'Gregorian (U.S.) calendar
Public Const CAL_JAPAN = 3                          'Japanese Emperor Era calendar
Public Const CAL_TAIWAN = 4                         'Taiwan Region Era calendar
Public Const CAL_KOREA = 5                          'Korean Tangun Era calendar
Public Const CAL_HIJRI = 6                          'Hijri (Arabic Lunar) calendar
Public Const CAL_THAI = 7                           'Thai calendar
Public Const CAL_HEBREW = 8                         'Hebrew (Lunar) calendar
Public Const CAL_GREGORIAN_ME_FRENCH = 9            'Gregorian Middle East French calendar
Public Const CAL_GREGORIAN_ARABIC = 10              'Gregorian Arabic calendar
Public Const CAL_GREGORIAN_XLIT_ENGLISH = 11        'Gregorian Transliterated English calendar
Public Const CAL_GREGORIAN_XLIT_FRENCH = 12         'Gregorian Transliterated French calendar

' Calendar Enumeration Value.
Public Const ENUM_ALL_CALENDARS As Long = &HFFFFFFFF 'enumerate all calendars
Public Const CAL_ICALINTVALUE As Long = &H1         'calendar type
Public Const CAL_SCALNAME As Long = &H2             'native name of calendar
Public Const CAL_IYEAROFFSETRANGE As Long = &H3     'starting years of eras
Public Const CAL_SERASTRING As Long = &H4           'era name for IYearOffsetRanges
Public Const CAL_SSHORTDATE As Long = &H5           'short date format string
Public Const CAL_SLONGDATE As Long = &H6            'long date format string
Public Const CAL_SDAYNAME1 As Long = &H7            'native name for Monday
Public Const CAL_SDAYNAME2 As Long = &H8            'native name for Tuesday
Public Const CAL_SDAYNAME3 As Long = &H9            'native name for Wednesday
Public Const CAL_SDAYNAME4 As Long = &HA            'native name for Thursday
Public Const CAL_SDAYNAME5 As Long = &HB            'native name for Friday
Public Const CAL_SDAYNAME6 As Long = &HC            'native name for Saturday
Public Const CAL_SDAYNAME7 As Long = &HD            'native name for Sunday
Public Const CAL_SABBREVDAYNAME1 As Long = &HE      'abbreviated name for Monday
Public Const CAL_SABBREVDAYNAME2 As Long = &HF      'abbreviated name for Tuesday
Public Const CAL_SABBREVDAYNAME3 As Long = &H10     'abbreviated name for Wednesday
Public Const CAL_SABBREVDAYNAME4 As Long = &H11     'abbreviated name for Thursday
Public Const CAL_SABBREVDAYNAME5 As Long = &H12     'abbreviated name for Friday
Public Const CAL_SABBREVDAYNAME6 As Long = &H13     'abbreviated name for Saturday
Public Const CAL_SABBREVDAYNAME7 As Long = &H14     'abbreviated name for Sunday
Public Const CAL_SMONTHNAME1 As Long = &H15         'native name for January
Public Const CAL_SMONTHNAME2 As Long = &H16         'native name for February
Public Const CAL_SMONTHNAME3 As Long = &H17         'native name for March
Public Const CAL_SMONTHNAME4 As Long = &H18         'native name for April
Public Const CAL_SMONTHNAME5 As Long = &H19         'native name for May
Public Const CAL_SMONTHNAME6 As Long = &H1A         'native name for June
Public Const CAL_SMONTHNAME7 As Long = &H1B         'native name for July
Public Const CAL_SMONTHNAME8 As Long = &H1C         'native name for August
Public Const CAL_SMONTHNAME9 As Long = &H1D         'native name for September
Public Const CAL_SMONTHNAME10 As Long = &H1E        'native name for October
Public Const CAL_SMONTHNAME11 As Long = &H1F        'native name for November
Public Const CAL_SMONTHNAME12 As Long = &H20        'native name for December
Public Const CAL_SMONTHNAME13 As Long = &H21        'native name for 13th month (if any)
Public Const CAL_SABBREVMONTHNAME1 As Long = &H22   'abbreviated name for January
Public Const CAL_SABBREVMONTHNAME2 As Long = &H23   'abbreviated name for February
Public Const CAL_SABBREVMONTHNAME3 As Long = &H24   'abbreviated name for March
Public Const CAL_SABBREVMONTHNAME4 As Long = &H25   'abbreviated name for April
Public Const CAL_SABBREVMONTHNAME5 As Long = &H26   'abbreviated name for May
Public Const CAL_SABBREVMONTHNAME6 As Long = &H27   'abbreviated name for June
Public Const CAL_SABBREVMONTHNAME7 As Long = &H28   'abbreviated name for July
Public Const CAL_SABBREVMONTHNAME8 As Long = &H29   'abbreviated name for August
Public Const CAL_SABBREVMONTHNAME9 As Long = &H2A   'abbreviated name for September
Public Const CAL_SABBREVMONTHNAME10 As Long = &H2B  'abbreviated name for October
Public Const CAL_SABBREVMONTHNAME11 As Long = &H2C  'abbreviated name for November
Public Const CAL_SABBREVMONTHNAME12 As Long = &H2D  'abbreviated name for December
Public Const CAL_SABBREVMONTHNAME13 As Long = &H2E  'abbreviated name for 13th month (if any)

'WINVER >= =&H0500 only
Public Const CAL_SYEARMONTH As Long = &H2F          'year month format string
Public Const HORZRES = 8     '*Tee 49-01-24
Public Const VERTRES = 10     '*Tee 49-01-24

'********** API Declaration **********
Public gstrISPName As String
Public ReturnCode2 As Long
Public Const ERROR_SUCCESS = 0&
Public Const APINULL = 0&
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_CURRENT_USER = &H80000001 '*TON 58-12-01
Public ReturnCode1 As Long
Public Declare Function RasHangUp Lib "RasApi32.dll" Alias "RasHangUpA" (ByVal hRasConn As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
'Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long

Public Declare Function RasEnumEntries Lib "RasApi32.dll" Alias "RasEnumEntriesA" (ByVal reserved As String, ByVal lpszPhonebook As String, lprasentryname As Any, lpcb As Long, lpcEntries As Long) As Long

Public Declare Function RasEnumConnections Lib "RasApi32.dll" Alias "RasEnumConnectionsA" (lpRasCon As Any, lpcb As Long, lpcConnections As Long) As Long
Public Declare Function RasGetConnectStatus Lib "RasApi32.dll" Alias "RasGetConnectStatusA" (ByVal hRasCon As Long, lpStatus As Any) As Long
Public Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal himl As Long, ByVal i As Long, ByVal hDCDest As Long, ByVal X As Long, ByVal Y As Long, ByVal Flags As Long) As Long
Public Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbSizeFileInfo As Long, ByVal uFlags As Long) As Long

'shell32
Public Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long 'ITEMIDLIST
Public Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Public Declare Function ShellEx Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As Any, ByVal lpDirectory As Any, ByVal nShowCmd As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'advapi32
Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

'winspool   control POS printer
Public Declare Function EndDocPrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Public Declare Function StartDocPrinter Lib "winspool.drv" Alias "StartDocPrinterA" (ByVal hPrinter As Long, ByVal Level As Long, pDocInfo As DOCINFO) As Long
Public Declare Function EndPagePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Public Declare Function StartPagePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Public Declare Function ClosePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Public Declare Function OpenPrinter Lib "winspool.drv" Alias "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, ByVal pDefault As Long) As Long
Public Declare Function WritePrinter Lib "winspool.drv" (ByVal hPrinter As Long, pBuf As Any, ByVal cdBuf As Long, pcWritten As Long) As Long

'kernel32
Public Declare Function GetVolumeInformation& Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal pVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long)
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function SleepEx Lib "kernel32" (ByVal dwMilliseconds As Long, ByVal bAlertable As Long) As Long
Public Declare Function GetDateFormat Lib "kernel32" Alias "GetDateFormatA" (ByVal Locale As Long, ByVal dwFlags As Long, lpDate As SYSTEMTIME, ByVal lpFormat As String, ByVal lpDateStr As String, ByVal cchDate As Long) As Long
Public Declare Function GetVersion Lib "kernel32" () As Long
'*WD 48-08-30
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
'***

'user32
Public Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function ExitWindowsEx& Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long)
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function Putfocus Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long

'-----------------------End Regional Setting ----------------------------------------
Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long     '*Tee 49-01-24

'for window functions (API)
Public Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function GetDiskFreeSpaceEx Lib "kernel32" Alias "GetDiskFreeSpaceExA" (ByVal lpRootPathName As String, lpFreeBytesAvailableToCaller As Currency, lpTotalNumberOfBytes As Currency, lpTotalNumberOfFreeBytes As Currency) As Long      '*Tee 47-02-13
Public Declare Function GetFileInformationByHandle Lib "kernel32" (ByVal hFile As Long, lpFileInformation As BY_HANDLE_FILE_INFORMATION) As Long

'-----------------------Start Regional Setting *WD Add 47-11-24------------------------------------------
'kernel32
Public Declare Function SetLocaleInfo Lib "kernel32" Alias "SetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String) As Long
Public Declare Function GetSystemDefaultLCID Lib "kernel32" () As Long
Public Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
Public Declare Function EnumCalendarInfo Lib "kernel32" Alias "EnumCalendarInfoA" (ByVal lpCalInfoEnumProc As Long, ByVal Locale As Long, ByVal Calendar As Long, ByVal CalType As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function EnumDateFormats Lib "kernel32" Alias "EnumDateFormatsA" (ByVal lpDateFmtEnumProc As Long, ByVal Locale As Long, ByVal dwFlags As Long) As Long
'user32
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const WM_CLOSE = &H10

Public Const CCHDEVICENAME = 32
Public Const CCHFORMNAME = 32
'Screen
Public Type DEVMODE
    dmDeviceName As String * CCHDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * CCHFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type

Public Const BITSPIXEL = 12

'/* Return values for ChangeDisplaySettings */
Public Const DISP_CHANGE_SUCCESSFUL = 0
Public Const DISP_CHANGE_RESTART = 1
Public Const DISP_CHANGE_FAILED = -1
Public Const DISP_CHANGE_BADMODE = -2
Public Const DISP_CHANGE_NOTUPDATED = -3
Public Const DISP_CHANGE_BADFLAGS = -4
Public Const DISP_CHANGE_BADPARAM = -5


' /* Flags for ChangeDisplaySettings */
Public Const CDS_UPDATEREGISTRY = &H1
Public Const CDS_TEST = &H2
Public Const CDS_FULLSCREEN = &H4
Public Const CDS_GLOBAL = &H8
Public Const CDS_SET_PRIMARY = &H10
Public Const CDS_RESET = &H40000000
Public Const CDS_SETRECT = &H20000000
Public Const CDS_NORESET = &H10000000

Public Const DM_BITSPERPEL = &H40000
Public Const DM_PELSWIDTH = &H80000
Public Const DM_PELSHEIGHT = &H100000
Public Const DM_DISPLAYFLAGS = &H200000
Public Const DM_DISPLAYFREQUENCY = &H400000



Public Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lpInitData As DEVMODE, ByVal dwFlags As Long) As Long
Public Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (lpszDeviceName As Any, ByVal iModeNum As Long, lpDevMode As Any) As Boolean


'*WD-SFM 49-03-02
Public Const PING_TIMEOUT = 200
Public Const WS_VERSION_REQD = &H101
Public Const WS_VERSION_MAJOR = WS_VERSION_REQD \ &H100 And &HFF&
Public Const WS_VERSION_MINOR = WS_VERSION_REQD And &HFF&
Public Const MIN_SOCKETS_REQD = 1
Public Const MAX_WSADescription = 256
Public Const MAX_WSASYSStatus = 128

Private nCN_SPNumModes As Long
Private oCN_SPDM() As DEVMODE

Public Type WIN32_FIND_DATA
        dwFileAttributes As Long
        ftCreationTime As FILETIME
        ftLastAccessTime As FILETIME
        ftLastWriteTime As FILETIME
        nFileSizeHigh As Long
        nFileSizeLow As Long
        dwReserved0 As Long
        dwReserved1 As Long
        cFileName As String * MAX_PATH
        cAlternate As String * 14
End Type

'*ZEAH 08-02-28
Public Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Public Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
Public Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Public Declare Function SystemTimeToFileTime Lib "kernel32" (lpSystemTime As SYSTEMTIME, lpFileTime As FILETIME) As Long

'*ZEAH 08-02-28
Public Const FILE_ATTRIBUTE_ARCHIVE = &H20
Public Const FILE_ATTRIBUTE_COMPRESSED = &H800
Public Const FILE_ATTRIBUTE_DIRECTORY = &H10
Public Const FILE_ATTRIBUTE_HIDDEN = &H2
Public Const FILE_ATTRIBUTE_READONLY = &H1
Public Const FILE_ATTRIBUTE_SYSTEM = &H4
Public Const FILE_ATTRIBUTE_TEMPORARY = &H100
Public Const INVALID_HANDLE_VALUE = -1

Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long


Public Declare Function WSACleanup Lib "wsock32.dll" () As Long
Public Declare Function IcmpCreateFile Lib "Icmp.dll" () As Long
Public Declare Function IcmpCloseHandle Lib "Icmp.dll" (ByVal IcmpHandle As Long) As Long
Public Declare Sub RtlMoveMemory Lib "kernel32" (hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)
Public Declare Function WSAStartup Lib "wsock32.dll" (ByVal wVersionRequired As Long, lpWSAData As WSAdata) As Long
Public Declare Function IcmpSendEcho Lib "Icmp.dll" (ByVal IcmpHandle As Long, ByVal DestinationAddress As Long, ByVal RequestData As String, ByVal RequestSize As Long, ByVal RequestOptions As Long, ReplyBuffer As ICMP_ECHO_REPLY, ByVal ReplySize As Long, ByVal TimeOut As Long) As Long

'Winsock
Declare Function gethostbyname& Lib "wsock32.dll" (ByVal hostname$)
Public Type Hostent
    hName As Long
    hAliases As Long
    hAddrType As Integer
    hLength As Integer
    hAddrList As Long
End Type
Public Type WSAdata
    wVersion As Integer
    wHighVersion As Integer
    szDescription(0 To MAX_WSADescription) As Byte
    szSystemStatus(0 To MAX_WSASYSStatus) As Byte
    wMaxSockets As Long
    wMaxUDPDG As Long
    dwVendorInfo As Long
End Type
Public Type ICMP_ECHO_REPLY
    Address         As Long
    Status          As Long
    RoundTripTime   As Long
    DataSize        As Long  'formerly integer
    DataPointer     As Long
    'Options         As ICMP_OPTIONS
    data            As String * 250
End Type
Public Type ICMP_OPTIONS
    Ttl             As Byte
    Tos             As Byte
    Flags           As Byte
    OptionsSize     As Byte
    OptionsData     As Long
End Type
'***

'*KT 51-08-08 System Date
Public Declare Sub GetSystemTime Lib "kernel32.dll" (lpSystemTime As SYSTEMTIME)
Public Declare Function SetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME) As Long
'*Ao 52-02-23
Public Declare Function GetSystemMetrics Lib "user32.dll" (ByVal nIndex As Long) As Long

Public Declare Function GetInputState Lib "user32" () As Long '*KT 52-09-10 การทำให้คำสั่ง DoEvent CPU ทำงานสั้นลง

Public Function AP_APIDisableCtrlAltDel()
' ----------------------------------------------------------
'   Call:
'
'   Cmt : Lock Keypress Ctrl+Alt+Del button
' ----------------------------------------------------------
    Dim iRet As Integer
    Dim bOld As Boolean
    
    iRet = SystemParametersInfo(SPI_SCREENSAVERRUNNING, True, bOld, 0)
End Function

Public Function AP_APIEnableCtrlAltDel()
' ----------------------------------------------------------
'   Call:
'
'   Cmt : Un Lock KeyPress Ctrl+Alt+Del
' ----------------------------------------------------------
    Dim iRet As Integer
    Dim bOld As Boolean
    
    iRet = SystemParametersInfo(SPI_SCREENSAVERRUNNING, False, bOld, 0)
End Function

Public Function AP_APILogUserOff()
'----------------------------------------------------------
'   Call:
'
'   Cmt : Return to log Off User
' ---------------------------------------------------------
    ExitWindowsEx EWX_LOGOFF, 0
End Function

Public Function AP_APIForceClose()
'----------------------------------------------------------
'   Call:
'
'   Cmt :
'----------------------------------------------------------
    ExitWindowsEx EWX_FORCE, 0
End Function

Public Function AP_APIShutdown()
' ----------------------------------------------------------
'   Call:
'
'   Cmt:    to shut down your computer
' ----------------------------------------------------------
    ExitWindowsEx EWX_SHUTDOWN, 0
    ExitWindowsEx EWX_SHUTDOWN, 0
    ExitWindowsEx EWX_SHUTDOWN, 0
End Function

Public Function AP_APIReboot()
' ----------------------------------------------------------
'   Call:
'
'   Cmt:    to restart  your Computer
' ----------------------------------------------------------
    ExitWindowsEx EWX_REBOOT, 0
    ExitWindowsEx EWX_REBOOT, 0
    ExitWindowsEx EWX_REBOOT, 0
End Function

Public Sub AP_APIMouseShow()
' ----------------------------------------------------------
'   Call:
'
'   Cmt:    shows your mouse cursor
' ----------------------------------------------------------
    Dim nRet As Long
    
    nRet = ShowCursor(True)
End Sub

Public Sub AP_APIMouseHide()
'----------------------------------------------------------
'   Call:
'
'   Cmt:    hide your mouse cursor
'---------------------------------------------------------
    Dim nRet As Long
    
    nRet = ShowCursor(False)
End Sub

Public Function AP_APIShowStartButton()
'---------------------------------------------------------
'   Call:
'
'   Cmt:    show start up button in desktop
'---------------------------------------------------------
    Dim handle As Long
    Dim FindClass As Long
    
    FindClass& = FindWindow("Shell_TrayWnd", "")
    handle& = FindWindowEx(FindClass&, 0, "Button", vbNullString)
    ShowWindow handle&, 1
End Function

Public Function AP_APIShowTaskbar()
'---------------------------------------------------------
'   Call:
'
'   Cmt:    show taskbar (normal)
'---------------------------------------------------------
    Dim handle As Long
    
    handle& = FindWindow("Shell_TrayWnd", vbNullString)
    ShowWindow handle&, 1
End Function

Public Function AP_APIHideTaskbar()
'---------------------------------------------------------
'   Call:
'
'   Cmt:    to visible taskbar(Lock Taskbar)
'---------------------------------------------------------
    Dim handle As Long
    
    handle& = FindWindow("Shell_TrayWnd", vbNullString)
    ShowWindow handle&, 0
End Function

Public Function AP_APIHideStartButton()
'---------------------------------------------------------
'   Call:
'
'   Cmt:    hide start button , do not working start menu
'---------------------------------------------------------
    Dim handle As Long
    Dim FindClass As Long
    
    FindClass& = FindWindow("Shell_TrayWnd", "")
    handle& = FindWindowEx(FindClass&, 0, "Button", vbNullString)
    ShowWindow handle&, 0
End Function

Public Sub AP_APIWindowsLog()
' ---------------------------------------------------------
'   Call:
'
'   Cmt:    Set windows to lock working with system
' ---------------------------------------------------------
    Call AP_APIHideTaskbar
    Call AP_APIHideStartButton
    Call AP_APIDisableCtrlAltDel
End Sub

Public Sub AP_APIWindowsUnLog()
'---------------------------------------------------------
'   Call:
'
'   Cmt: set windows to unlock system windows
'---------------------------------------------------------
    Call AP_APIShowTaskbar
    Call AP_APIShowStartButton
    Call AP_APIEnableCtrlAltDel
End Sub
'
'Public Sub SP_FRMxMoveForm(poForm As Form)
''   -------------------------------------------------------------------------------------------
''      Use               : MoveForm  โดยไม่ต้องคลิ้กที่ ToolBar
''      Parameter  :
''      Call               : เรียกใน Form_MourseDown Call API_MoveForm(Me)
''      Return          :
''      API                :
''                              Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
''      Design By  :  Tee  & Air 07/05/2001 00:35 PM
''   -------------------------------------------------------------------------------------------
'    ReleaseCapture
'   Call SendMessage(poForm.hWnd, &HA1, 2, 0&)
'End Sub


Public Function AP_APIbIsWindowsXP() As Boolean
'-----------------------------------------------------------
'   Call :
'   Cmt :
'   Developer : SOMJAI
'   Date Create : 07-29-2005 :  17:46:24
'-----------------------------------------------------------
    On Error GoTo ErrHandle
    Dim Ver As Long, WinVer As Long
    Dim tWinVersion$
    
    Ver = GetVersion()
    WinVer = Ver And &HFFFF&
    'retrieve the windows version
    tWinVersion = Format((WinVer Mod 256) + ((WinVer \ 256) / 100), "Fixed")

    If SP_VALcText2Double(Trim(tWinVersion)) >= 5.01 Then
        AP_APIbIsWindowsXP = True
    End If
    Exit Function
ErrHandle:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Function

Private Sub AP_SocketsCleanup()
Dim X As Long
   
  'need to use a var (insread of embedding
  'in the If..Then call) becuse the function
  'returns the error code if failed.
   X = WSACleanup()

   If X <> 0 Then
       MsgBox "Windows Sockets error " & Trim$(Str$(X)) & " occurred in Cleanup.", vbExclamation
   End If
    
End Sub

Public Function AP_ResolveIpaddress(ByVal ptHostName As String) As Collection
'---------------------------------------
'*WD-SFM 49-03-02
'---------------------------------------
Dim hostent_addr As Long
Dim host As Hostent
Dim hostip_addr As Long
Dim temp_ip_address() As Byte
Dim i As Integer
Dim ip_address As String
Dim Count As Integer

    If AP_SocketsInitialize() Then

        Set AP_ResolveIpaddress = New Collection
        hostent_addr = gethostbyname(ptHostName)

        If hostent_addr = 0 Then
            Call AP_SocketsCleanup
            Exit Function
        End If

        RtlMoveMemory host, hostent_addr, LenB(host)
        RtlMoveMemory hostip_addr, host.hAddrList, 4

        'get all of the IP address if machine is  multi-homed

        Do
            ReDim temp_ip_address(1 To host.hLength)
            RtlMoveMemory temp_ip_address(1), hostip_addr, host.hLength

            For i = 1 To host.hLength
                ip_address = ip_address & temp_ip_address(i) & "."
            Next
            ip_address = Mid$(ip_address, 1, Len(ip_address) - 1)
            AP_ResolveIpaddress.Add ip_address
            ip_address = ""
            host.hAddrList = host.hAddrList + LenB(host.hAddrList)
            RtlMoveMemory hostip_addr, host.hAddrList, 4
         Loop While (hostip_addr <> 0)

        AP_SocketsCleanup
    End If
End Function

'Public Function AP_ResolveIP(ByVal ptHostName As String) As String
''--------------------------------------------------
''*WD-SFM 49-03-02
''Cmt : Hostname to IP Section
''--------------------------------------------------
'Dim retColl As Collection
'
'    Set retColl = AP_ResolveIpaddress(ptHostName)
'
'    If retColl.Count > 0 Then
'        AP_ResolveIP = CStr(retColl.Item(1))
'    End If
'
'End Function

Public Function AP_ResolveIP(ByVal ptHostName As String) As String
'--------------------------------------------------
'*WD-SFM 49-03-02
'Cmt : Hostname to IP Section
'--------------------------------------------------
Dim retColl As Collection
Dim n As Integer
Dim tTmpIP As String
    Set retColl = AP_ResolveIpaddress(ptHostName)
    If retColl.Count > 0 Then
        For n = 1 To retColl.Count
            tTmpIP = CStr(retColl.Item(1))
            If (Left$(tTmpIP, 2) = "19") Or (Left$(tTmpIP, 2) = "17") Or (Left$(tTmpIP, 2) = "10") Then AP_ResolveIP = tTmpIP
        Next n
    End If
End Function

Public Function AP_GetIPFromHostName(ByVal ptHostName As String) As String
On Error GoTo ErrHandle
    '*Tao 59-11-01
    Dim nbytes As Long
    Dim tHosent As Long
    Dim tName As Long
    Dim tAddress As Long
    Dim ptrIPAddress As Long
    Dim tIPAddress As String
    On Error Resume Next
    tIPAddress = Space$(4)
    tHosent = gethostbyname(ptHostName & vbNullChar)
    If tHosent <> 0 Then
        tAddress = tHosent + 12
        CopyMemory tAddress, ByVal tAddress, 4            'get the IP address
        CopyMemory ptrIPAddress, ByVal tAddress, 4
        CopyMemory ByVal tIPAddress, ByVal ptrIPAddress, 4
        AP_GetIPFromHostName = CStr(Asc(tIPAddress)) & "." & CStr(Asc(Mid$(tIPAddress, 2, 1))) & "." & CStr(Asc(Mid$(tIPAddress, 3, 1))) & "." & CStr(Asc(Mid$(tIPAddress, 4, 1)))
    Else
        AP_GetIPFromHostName = ptHostName
    End If
    Exit Function
ErrHandle:
    Call SP_CRTxLogPrg("ErrorNumber:" & Err.Number & " ErrorDesc:" & Err.Description & " : mCNAP.AP_GetIPFromHostName")
End Function

Private Function AP_LoByte(ByVal pnParam As Long) As Integer
    AP_LoByte = pnParam And &HFF&
End Function

Private Function AP_HiByte(ByVal pnParam As Long) As Integer
    AP_HiByte = pnParam \ &H100 And &HFF&
End Function
'
Public Function AP_SocketsInitialize() As Boolean
Dim WSAD As WSAdata
Dim X As Integer
Dim szLoByte As String
Dim szHiByte As String
Dim szBuf As String

    X = WSAStartup(WS_VERSION_REQD, WSAD)

   'check for valid response
   'MsgBox "Windows Sockets for 32 bit Windows environments is not successfully responding."
    If X <> 0 Then Exit Function

   'check that the version of sockets is supported
    If AP_LoByte(WSAD.wVersion) < WS_VERSION_MAJOR Or _
       (AP_LoByte(WSAD.wVersion) = WS_VERSION_MAJOR And _
        AP_HiByte(WSAD.wVersion) < WS_VERSION_MINOR) Then

        szHiByte = Trim$(Str$(AP_HiByte(WSAD.wVersion)))
        szLoByte = Trim$(Str$(AP_LoByte(WSAD.wVersion)))
        szBuf = "Windows Sockets Version " & szLoByte & "." & szHiByte
        szBuf = szBuf & " is not supported by Windows " & _
                          "Sockets for 32 bit Windows environments."
        'MsgBox szBuf, vbExclamation
        Exit Function
    End If

   'check that there are available sockets
    If WSAD.wMaxSockets < MIN_SOCKETS_REQD Then
        szBuf = "This application requires a minimum of " & _
                 Trim$(Str$(MIN_SOCKETS_REQD)) & " supported sockets."
        'MsgBox szBuf, vbExclamation
        Exit Function
    End If

    AP_SocketsInitialize = True

End Function

Private Function AP_AddressStringToLong(ByVal tmp As String) As Long
Dim i As Integer
Dim parts(1 To 4) As String
   
    i = 0
    
    If InStr(1, tmp, ".", vbTextCompare) = 0 Then
        AP_AddressStringToLong = gethostbyname(tmp)
    Else
        'we have to extract each part of the
        '123.456.789.123 string, delimited by
        'a period
        While InStr(tmp, ".") > 0
          i = i + 1
          parts(i) = Mid(tmp, 1, InStr(tmp, ".") - 1)
          tmp = Mid(tmp, InStr(tmp, ".") + 1)
        Wend
        
        i = i + 1
        parts(i) = tmp
        
        If i <> 4 Then
          AP_AddressStringToLong = 0
          Exit Function
        End If
        
        'build the long value out of the
        'hex of the extracted strings
        AP_AddressStringToLong = Val("&H" & Right("00" & Hex(parts(4)), 2) & Right("00" & Hex(parts(3)), 2) & Right("00" & Hex(parts(2)), 2) & Right("00" & Hex(parts(1)), 2))
   End If
End Function
'
Private Function AP_PingAddress(szAddress As String, ECHO As ICMP_ECHO_REPLY, Optional TimeOut As Long = PING_TIMEOUT) As Long
Dim hPort As Long
Dim dwAddress As Long
Dim sDataToSend As String
Dim iOpt As Long

   sDataToSend = "Echo This"
   dwAddress = AP_AddressStringToLong(szAddress)

   hPort = IcmpCreateFile()

   If IcmpSendEcho(hPort, dwAddress, sDataToSend, Len(sDataToSend), 0, ECHO, Len(ECHO), TimeOut) Then
        'the ping succeeded,
        '.Status will be 0
        '.RoundTripTime is the time in ms for
        '               the ping to complete,
        '.Data is the data returned (NULL terminated)
        '.Address is the Ip address that actually replied
        '.DataSize is the size of the string in .Data
         AP_PingAddress = ECHO.RoundTripTime
   Else
        AP_PingAddress = ECHO.Status * -1
   End If

   Call IcmpCloseHandle(hPort)

End Function

Public Function AP_Ping(ByVal ptHostOrIP As String, Optional pnTimeOutmSec As Long = PING_TIMEOUT) As Boolean
Dim echoValues As ICMP_ECHO_REPLY
Dim pos As Integer
Dim Count As Integer
Dim returnIp As Collection

On Error GoTo ErrHandle
    If Trim(ptHostOrIP) = "" Then
        AP_Ping = False
        Exit Function
    End If

    If AP_SocketsInitialize() Then

        If InStr(1, ptHostOrIP, ".", vbTextCompare) <> 0 Then
            If IsNumeric(Mid(ptHostOrIP, 1, InStr(1, ptHostOrIP, ".") - 1)) = False Then
                Set returnIp = AP_ResolveIpaddress(ptHostOrIP)
                If returnIp.Count = 0 Then
                    AP_Ping = False
                    Exit Function
                Else
                    ptHostOrIP = returnIp.Item(1)
                End If
            End If
        End If

        'ping an ip address, passing the
        'address and the ECHO structure
        Call AP_PingAddress((ptHostOrIP), echoValues, pnTimeOutmSec)

        If Left$(echoValues.data, 1) <> Chr$(0) Then
           pos = InStr(echoValues.data, Chr$(0))
           echoValues.data = Left$(echoValues.data, pos - 1)
        Else
              echoValues.data = ""
        End If

        AP_SocketsCleanup

        If echoValues.Status <> 0 Then
            AP_Ping = False
        Else
            AP_Ping = True
        End If
    End If
    Exit Function
ErrHandle:
    AP_Ping = False
End Function


Public Sub SP_PRCxScreen(ByVal pnWidth As Long, ByVal pnHeight As Long, pForm As Form, Optional ByVal pnBits As Long = 60)
'--------------------------------------------------------------------------------------
'*KT 50-01-11
'Cmt:   ทำการ Get Resolution ทั้งหมดของ Windows แล้วหาว่าปัจจุบัน Screen เป็นแบบขนาดที่ต้องการหรือไม่ ถ้าไม่ให้ set resolution
'--------------------------------------------------------------------------------------
On Error GoTo ErrHandle
    Dim nMaxModes As Long, nList As Long, nScreen As Long
    Dim nBits As Long, nWidth As Long, nHeight As Long, nCount As Long
    Dim tTmp As String
    Dim aScreen(1000) As String
    Static nAccess As Byte
    Call SP_CRTxWriteLogTerminated("ErrorNumber:0" & " ErrorDesc:" & " Form:" & "mCNAP" & " Events:SP_PRCxScreen[1]")
    nVB_ScreenDefault = False
    
    If GetSetting("AdaPosMall", "Desktop", "Resolution") <> "" Then
        nVB_OldPositionScreen = CInt(GetSetting("AdaPosMall", "Desktop", "Resolution"))
    End If
    
    nAccess = nAccess + 1
    If nAccess > 1 Then Exit Sub
    
    nBits = GetDeviceCaps(pForm.hDC, BITSPIXEL)
    pnBits = nBits
    nWidth = GetDeviceCaps(pForm.hDC, HORZRES)
    nHeight = GetDeviceCaps(pForm.hDC, VERTRES)
    
    nMaxModes = 8
    ReDim oCN_SPDM(0 To nMaxModes) As DEVMODE
    nCN_SPNumModes = 0
    nList = EnumDisplaySettings(ByVal 0, nCN_SPNumModes, oCN_SPDM(nCN_SPNumModes))
    nCount = 0
    Do While nList
        aScreen(nCount) = oCN_SPDM(nCN_SPNumModes).dmPelsWidth & "x" & oCN_SPDM(nCN_SPNumModes).dmPelsHeight & "x" & oCN_SPDM(nCN_SPNumModes).dmBitsPerPel
        nCount = nCount + 1
        If nBits = oCN_SPDM(nCN_SPNumModes).dmBitsPerPel And _
            nWidth = oCN_SPDM(nCN_SPNumModes).dmPelsWidth And _
            nHeight = oCN_SPDM(nCN_SPNumModes).dmPelsHeight Then
            If GetSetting("AdaPosMall", "Desktop", "Resolution") = "" Then
                nVB_OldPositionScreen = nCount - 1 'เก็บตำแหน่งของ Resolution ก่อนเข้าโปรแกรม
                Call SaveSetting("AdaPosMall", "Desktop", "Resolution", nVB_OldPositionScreen)
            End If
            If (nWidth = pnWidth And nHeight = pnHeight And nBits = pnBits) Then
                nVB_ScreenDefault = True
            End If
        End If
        nCN_SPNumModes = nCN_SPNumModes + 1
        If nCN_SPNumModes > nMaxModes Then
            nMaxModes = nMaxModes + 8
            ReDim Preserve oCN_SPDM(0 To nMaxModes) As DEVMODE
        End If
        nList = EnumDisplaySettings(ByVal 0, nCN_SPNumModes, oCN_SPDM(nCN_SPNumModes))
    Loop
    
    nCN_SPNumModes = nCN_SPNumModes - 1
    If nVB_ScreenDefault = False Then
        tTmp = pnWidth & "x" & pnHeight & "x" & nBits
        For nList = 0 To nCount - 1
            If tTmp = aScreen(nList) Then
                Call SP_SETxScreen(nList)
                SleepEx 1500, 0
                Exit Sub
            End If
        Next nList
    End If
    Exit Sub
    Call SP_CRTxWriteLogTerminated("ErrorNumber:0" & " ErrorDesc:" & " Form:" & "mCNAP" & " Events:SP_PRCxScreen[2]")
    Exit Sub
ErrHandle:
    Call SP_CRTxWriteLogTerminated("ErrorNumber:" & Err.Number & " ErrorDesc:" & Err.Description & " Form:" & "mCNAP" & " Events:SP_PRCxScreen[3]")
End Sub

Public Sub SP_SETxScreen(ByVal pnPosition As Long)
'--------------------------------------------------------------------------
'Call:
'Cmt:   ทำการ set resolution ตาม list ที่ต้องการ
'--------------------------------------------------------------------------
On Error GoTo ErrHandle
    Dim nList  As Long, nFlags As Long, nP As Long

'   If GetSetting("AdaPosTNS", "StoreBack", "Handle") <> "0" Then Exit Sub
   'ดูว่าโปรแกรมหน้าร้านเปิดอยู่หรือไม่ ถ้าเปิดอยู่จะไม่ทำการ set resolution กลับเหมือนเดิม  ตัวมันเองปล่อยให้ผ่านได้ครับ
Ln:
    nP = pnPosition
    oCN_SPDM(nP).dmFields = DM_BITSPERPEL Or DM_PELSWIDTH Or DM_PELSHEIGHT
    nFlags = CDS_UPDATEREGISTRY
    nList = ChangeDisplaySettings(oCN_SPDM(nP), nFlags)
    Select Case nList
        Case DISP_CHANGE_RESTART
            nList = MsgBox("This change will not take effect until you reboot the system. Reboot now?", vbYesNo)
            If nList = vbYes Then
                nFlags = 0
                nList = ExitWindowsEx(EWX_REBOOT, nFlags)
            End If
        Case DISP_CHANGE_SUCCESSFUL
        
    End Select
    Exit Sub
ErrHandle:
    If Err.Number = 13 Then
        GoTo Ln
    Else
        Call SP_MSGnShowing("tms_CN006004", tCS_CNMsgErr)
    End If
End Sub

