Attribute VB_Name = "mAP"
Option Explicit

'*TW 51-07-15 Service
Public Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion(1 To 128) As Byte
End Type
Public Const VER_PLATFORM_WIN32_NT = 2&
'---------------------------------------------------------------------------------

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long
Public Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hwnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)
Public Declare Function SetWindowText Lib "user32.dll" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long


'*TW 51-07-15 Service
Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

'Private Declare Function ShellExecute Lib "shell32" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Sub sendkeys(Text As Variant, Optional wait As Boolean = False)
    Dim WshShell As Object
    Set WshShell = CreateObject("wscript.shell")
    WshShell.sendkeys CStr(Text), wait
    Set WshShell = Nothing
End Sub
