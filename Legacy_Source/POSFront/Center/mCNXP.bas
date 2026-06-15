Attribute VB_Name = "mCNXP"
'
'Option Explicit
'
'Private Declare Function InitCommonControlsEx Lib "comctl32.dll" (init As InitCommonControlsExType) As Boolean
'Private Declare Function ActivateWindowTheme Lib "uxtheme.dll" Alias "SetWindowTheme" (ByVal hWnd As Long, Optional ByVal pszSubAppName As Long = 0, Optional ByVal pszSubIdList As Long = 0) As Long
'Private Declare Function DeactivateWindowTheme Lib "uxtheme.dll" Alias "SetWindowTheme" (ByVal hWnd As Long, Optional ByRef pszSubAppName As String = " ", Optional ByRef pszSubIdList As String = " ") As Long
'Private Declare Function IsThemeActiveXP Lib "uxtheme.dll" Alias "IsThemeActive" () As Boolean
'Private Declare Function IsAppThemed Lib "uxtheme.dll" () As Boolean
'Private Declare Function OpenThemeData Lib "uxtheme.dll" (ByVal hWnd As Long, ByVal pszClassList As Long) As Long
'Private Declare Function DrawThemeBackground Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal lHDC As Long, ByVal iPartId As Long, ByVal iStateId As Long, pRect As Rect, pClipRect As Rect) As Long
'Private Declare Function CloseThemeData Lib "uxtheme.dll" (hTheme As Long) As Long
'Private Declare Function EnableThemeDialogTexture Lib "uxtheme.dll" (ByVal hWnd As Long, ByVal dwFlags As Long) As Long
'Private Declare Function RedrawWindow Lib "USER32" (ByVal hWnd As Long, lprcUpdate As Rect, Optional hrgnUpdate As Long, Optional fuRedraw As Long) As Long
'Private Declare Function GetWindowRect Lib "USER32" (ByVal hWnd As Long, lpRect As Rect) As Long
'Private Declare Function SendMessage Lib "USER32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'Private Declare Function GetComboBoxInfo Lib "USER32" (ByVal hwndCombo As Long, CBInfo As COMBOBOXINFO) As Long
'Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
'Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
'Private Declare Function GetVersion Lib "kernel32" () As Long
'
'Private Declare Function PathIsNetworkPath Lib "shlwapi.dll" Alias "PathIsNetworkPathA" (ByVal pszPath As String) As Boolean
'
'Private Const ETDT_DISABLE      As Long = &H1
'Private Const ETDT_ENABLE       As Long = &H2
'
'Private Const RDW_UPDATENOW     As Long = &H100
'
'Private Const ICC_USEREX_CLASSES As Long = &H200
'
'Private Const ECM_FIRST         As Long = &H1500
'Private Const EM_SHOWBALLOONTIP As Long = (ECM_FIRST + 3)
'Private Const EM_HIDEBALLOONTIP As Long = (ECM_FIRST + 4)
'
'Private m_bIsManifestActive     As Boolean
'Private bIsVbRunning            As Boolean
'
'Private Type InitCommonControlsExType
'    dwSize  As Long     'size of this structure
'    dwICC   As Long     'flags indicating which classes to be initialized
'End Type
'
'Private Type Rect
'    Left    As Long
'    Top     As Long
'    Right   As Long
'    Bottom  As Long
'End Type
'
'Private Type BALLOONTIP
'    cbStruct As Long
'    pszTitle As String
'    pszText As String
'    tIcon As Long
'End Type
'
'Private Type COMBOBOXINFO
'   cbSize As Long
'   rcItem As Rect
'   rcButton As Rect
'   stateButton  As Long
'   hwndCombo  As Long
'   hwndEdit  As Long
'   hwndList As Long
'End Type
'
'Enum BalloonTipIconConstants
'    balNone = 0
'    balExcalmation = 1
'    balInformation = 2
'    balCritical = 3
'End Enum
'
'Private Function XP_OBJbInitCommonControls() As Boolean
'    Dim InitCC As InitCommonControlsExType
'
'    With InitCC
'        .dwSize = Len(InitCC)
'        .dwICC = ICC_USEREX_CLASSES
'    End With
'
'    XP_OBJbInitCommonControls = InitCommonControlsEx(InitCC)         'initialize the common controls
'End Function
'
'
'Private Function XP_OBJbCheckVB() As Boolean
'    bIsVbRunning = True
'    XP_OBJbCheckVB = True
'End Function
'
'
'Private Function XP_OBJnGetWindowTheme(hWnd As Long, Optional PartID As String) As Long
'    'this will retrive the current hTheme used by the window..
'    If XP_OBJbIsWindowsXP() = False Or XP_OBJbIsThemingSupported() = False Or XP_OBJbIsVBRunning = True Then Exit Function
'    If IsThemeActiveXP() = False Then Exit Function
'
'    Dim hTheme As Long
'    If PartID = "" Then PartID = "Window"
'    hTheme = OpenThemeData(hWnd, StrPtr(PartID))
'    CloseThemeData hTheme
'    XP_OBJnGetWindowTheme = hTheme
'
'End Function
'
'Private Function XP_OBJtGetWinVersion() As String
'    Dim Ver As Long, WinVer As Long
'    Ver = GetVersion()
'    WinVer = Ver And &HFFFF&
'    'retrieve the windows version
'    XP_OBJtGetWinVersion = Format((WinVer Mod 256) + ((WinVer \ 256) / 100), "Fixed")
'End Function
'
'Private Function XP_OBJtAddDirSep(Path As String) As String
'    Dim DirSep As String
'
'    If PathIsNetworkPath(Path) = True Then
'        DirSep = "/"
'    Else
'        DirSep = "\"
'    End If
'
'    If Right(Trim(Path), Len(DirSep)) <> DirSep Then
'        XP_OBJtAddDirSep = Trim(Path) & DirSep
'    Else
'        XP_OBJtAddDirSep = Path
'    End If
'
'End Function
'
'
'Private Function XP_OBJbHideTextBalloonTip(Control As Control) As Boolean
'
'    Dim hWnd As Long
'
'    If XP_OBJbIsWindowsXP() = False Or XP_OBJbIsThemingSupported() = False Or XP_OBJbIsVBRunning = True Then Exit Function
'    If IsThemeActiveXP() = False Then Exit Function
'
'    Select Case UCase(TypeName(Control))
'
'        Case "TEXTBOX"
'            hWnd = Control.hWnd
'        Case "RICHTEXTBOX"
'            hWnd = Control.hWnd
'        Case "COMBOBOX"
'            If (Control.Style = 0 Or 1) Then
'                Dim Cbo As COMBOBOXINFO
'                Cbo.cbSize = Len(Cbo)
'                Call GetComboBoxInfo(Control.hWnd, Cbo)
'                hWnd = Cbo.hwndEdit
'            Else
'                Exit Function
'            End If
'        Case Else
'            hWnd = Control.hWnd
'    End Select
'
'    XP_OBJbHideTextBalloonTip = SendMessage(hWnd, EM_HIDEBALLOONTIP, 0&, 0&)
'
'End Function
'
'Function XP_OBJbIsThemingSupported() As Boolean
'
'    Dim hLib As Long                    'module handle..
'    hLib = LoadLibrary("uxtheme.dll")   'retrive the module handle.
'    Call FreeLibrary(hLib)              'unload the dll
'    XP_OBJbIsThemingSupported = CBool(hLib)    'if the return value = 0 then
'                                        'the dll does not exist,
'                                        'otherwise, the dll is there..
'End Function
'
'
'
'
'
'Private Function XP_OBJbIsXPThemed(hWnd As Long) As Boolean
'
'    'check if the object is using a visual style..
'
'    If XP_OBJbIsWindowsXP() = False Or XP_OBJbIsThemingSupported() = False Or XP_OBJbIsVBRunning = True Then Exit Function
'    If IsThemeActiveXP() = False Then Exit Function
'
'
'    Dim hTheme As Long
'
'    hTheme = OpenThemeData(hWnd, StrPtr("Window"))  'do the theme test
'
'    Call CloseThemeData(hTheme)                     'close the theme data handle
'
'    XP_OBJbIsXPThemed = CBool(hTheme)                      'if zero, return False.. else return true..
'
'
'End Function
'
'
'Public Function XP_OBJbShowTextBalloonTip(Control As Control, Prompt As String, Optional Title As String, Optional TitleIcon As BalloonTipIconConstants) As Boolean
'
'    'This function will show an EDIT balloon tip..
'    'this function will only apply to a normal text box, a richtext box or a combobox
'    'with syle 0 or 1...
'    'any other controls passed to this function will return false (as i know!)
'
'    Dim Bal As BALLOONTIP
'    Dim hWnd As Long
'
'    If XP_OBJbIsWindowsXP() = False Or XP_OBJbIsThemingSupported() = False Or XP_OBJbIsVBRunning = True Then Exit Function
'    If IsThemeActiveXP() = False Then Exit Function
'
'    Select Case UCase(TypeName(Control))
'        Case "COMBOBOX"
'            If (Control.Style = 0 Or Control.Style = 1) Then
'                Dim Cbo As COMBOBOXINFO
'                Cbo.cbSize = Len(Cbo)
'                Call GetComboBoxInfo(Control.hWnd, Cbo)
'                hWnd = Cbo.hwndEdit
'            Else
'                Exit Function
'            End If
'        Case "TEXTBOX"
'            hWnd = Control.hWnd
'        Case "RICHTEXTBOX"
'            hWnd = Control.hWnd
'        Case Else
'            hWnd = Control.hWnd
'    End Select
'
'    With Bal
'        .cbStruct = Len(Bal)                    'set the structure size
'        .pszTitle = StrConv(Title, vbUnicode)   'convert the title into unicode format..
'        .pszText = StrConv(Prompt, vbUnicode)   'convert the prompt into unicode format..
'        .tIcon = TitleIcon                      'set the title icon
'    End With
'
'    'show the balloon tip..
'
'    XP_OBJbShowTextBalloonTip = SendMessage(hWnd, EM_SHOWBALLOONTIP, 0&, Bal)
'
'
'End Function
'
'Private Function XP_OBJxToggleVisualStyles(Frm As Form, Enable As Boolean, Optional ToggleFormBorder As Boolean = True)
'
'    'Enable/diable a form theming ..
'
'    On Error GoTo ErrorHandler
'
'    If XP_OBJbIsWindowsXP() = False Or XP_OBJbIsThemingSupported() = False Or XP_OBJbIsVBRunning = True Then Exit Function
'    If IsThemeActiveXP() = False Then Exit Function
'
'
'    Dim fControls   As String   'This is the forbidden controls (controls with no .hWnd/cannot be skinned)
'                                'i didn't use an array cause i found out that it's too slow..
'
'    Dim Ctl         As Control
'    Dim sTypeName   As String
'
'    For Each Ctl In Frm.Controls                            'loop through the controls collection of the form..
'        sTypeName = UCase(TypeName(Ctl))                    'Get the typename of the control.
'        If InStr(1, fControls, sTypeName) = 0 Then          'look for the control type name in the forbidden controls list, if found, do nothing..
'            Select Case Enable                              'activate/deactivate theming
'                Case True:  Call XP_OBJbEnableXPLook(Ctl)
'                Case False: Call XP_OBJbDisableXPLook(Ctl)
'            End Select
'            If TypeName(Ctl) = "PICTUREBOX" Then Ctl.Refresh    'refresh any pictureboxes in the form..
'        End If
'    Next
'
'    If ToggleFormBorder = True Then
'        Select Case Enable                                      'activate/deactivate the form theming..
'            Case True
'                Call XP_OBJbEnableXPLook(Frm)
'            Case False
'                Call XP_OBJbDisableXPLook(Frm): Call XP_OBJxDisableXPDlgBackground(Frm)
'        End Select
'    End If
'
'    Set Ctl = Nothing       'erase the ctl variable from memory..
'
'    Frm.Refresh             'refresh the form
'
'   'Debug.Print fControls
'    Exit Function
'
'ErrorHandler:                                   'This is the error handling section...
'
'    If Err.Number = 438 Then                    'object doesn't have a ".hWnd" property..
'        'Debug.Print "Error: The Object '" & Ctl.Name & "' doesn't have a '.hwnd' property.."
'        fControls = sTypeName & "," & fControls 'add this typename into the forbidden list..
'        Resume Next                             'skip the line where the error happened, and proceed to the next line..
'    Else                                        'unexpected error..
'        Err.Raise Err.Number                    'show the error..
'    End If
'End Function
'
'
'Private Function XP_OBJbEnableXPLook(ByRef Object As Object) As Boolean
'    'this function will draw the object using windows xp visual styles..
'    'note: the object MUST have a handle
'
'    On Error GoTo ErrHandler:
'
'    Dim wRECT   As Rect
'
'    GetWindowRect Object.hWnd, wRECT   'retrive the object region.
'
'    If XP_OBJbIsWindowsXP() = False Or XP_OBJbIsThemingSupported() = False Or XP_OBJbIsVBRunning = True Then Exit Function
'    If IsThemeActiveXP() = False Then Exit Function
'
'    ActivateWindowTheme (Object.hWnd) 'try to enable theming
'
'    If XP_OBJbIsXPThemed(Object.hWnd) <> 0 Then
'        'ok
'        XP_OBJbEnableXPLook = True
'    Else
'        'error
'        GoTo ErrHandler
'    End If
'
'    Call RedrawWindow(Object.hWnd, wRECT, , RDW_UPDATENOW) 'refresh the object
'
'    Exit Function
'ErrHandler:
'    XP_OBJbEnableXPLook = False
'    Exit Function
'End Function
'
'Private Function XP_OBJbDisableXPLook(ByRef Object As Object) As Boolean
'    'this function will disable the object's visual style..
'    'note: the object MUST have a handle
'    'same as the XP_OBJbEnableXPLook function..
'
'    Dim wRECT As Rect
'
'    On Error GoTo ErrHandler:
'
'    GetWindowRect Object.hWnd, wRECT
'
'    If XP_OBJbIsWindowsXP() = False Or XP_OBJbIsThemingSupported() = False Or XP_OBJbIsVBRunning = True Then Exit Function
'    If IsThemeActiveXP() = False Then Exit Function
'
'    DeactivateWindowTheme (Object.hWnd)
'
'    If XP_OBJbIsXPThemed(Object.hWnd) = False Then
'        XP_OBJbDisableXPLook = True
'    Else
'        GoTo ErrHandler
'    End If
'
'    Call RedrawWindow(Object.hWnd, wRECT, , RDW_UPDATENOW)
'
'    Exit Function
'
'ErrHandler:
'    XP_OBJbDisableXPLook = False
'    Exit Function
'End Function
'
'Private Function XP_OBJxDrawTabBackground(oPictureBox As Object, Optional sTab As Object)
'
'    On Error Resume Next
'    'Draw a TabStrip control's background texture in a picture box..
'    'this is a good example on how to draw controls using "uxtheme.dll" API calls..
'
'    Dim hTheme          As Long         'The theme handle
'    Dim dRECT           As Rect         'The drawing Region
'    Dim tabHwnd         As Long
'    Const TAB_BODY      As Integer = 10 'this is the PartID of the tabstrip background..
'
'    If XP_OBJbIsWindowsXP() = False Or XP_OBJbIsThemingSupported() = False Or XP_OBJbIsVBRunning = True Then Exit Function
'    If IsThemeActiveXP() = False Then Exit Function
'
'    tabHwnd = sTab.hWnd
'
'    If tabHwnd <> 0 Then
'        If XP_OBJbIsXPThemed(sTab.hWnd) = False Then oPictureBox.Cls: Exit Function  'if the frame theming is disabled, Clear the picture box and exit
'    End If
'
'    oPictureBox.Cls
'    oPictureBox.AutoRedraw = False
'
'    'copy the picturebox measurements into the RECT object
'
'    dRECT.Left = 0
'    dRECT.Top = 0
'    dRECT.Right = oPictureBox.ScaleX(oPictureBox.Width, oPictureBox.ScaleMode, vbPixels)
'    dRECT.Bottom = oPictureBox.ScaleY(oPictureBox.Height, oPictureBox.ScaleMode, vbPixels)
'
'    hTheme = OpenThemeData(oPictureBox.hWnd, StrPtr("TAB"))      'Retrive the handle of the current theme being used.
'
'    If hTheme <> 0 Then
'        Call DrawThemeBackground(hTheme, oPictureBox.hdc, TAB_BODY, 0, dRECT, dRECT) 'draw the tab background on the picture box
'    Else
'        oPictureBox.Cls
'    End If
'
'    oPictureBox.AutoRedraw = True
'
'    CloseThemeData hTheme           'close the theme data handle..
'
'End Function
'
'
'Private Sub XP_OBJxEnableXPDlgBackground(Frm As Form)
'    If XP_OBJbIsWindowsXP() = False Or XP_OBJbIsThemingSupported() = False Or XP_OBJbIsVBRunning = True Then Exit Sub
'    If IsThemeActiveXP() = False Then Exit Sub
'    Call EnableThemeDialogTexture(Frm.hWnd, ETDT_ENABLE)
'End Sub
'
'Private Sub XP_OBJxDisableXPDlgBackground(Form As Form)
'    If XP_OBJbIsWindowsXP() = False Or XP_OBJbIsThemingSupported() = False Or XP_OBJbIsVBRunning = True Then Exit Sub
'    If IsThemeActiveXP() = False Then Exit Sub
'    Call EnableThemeDialogTexture(Form.hWnd, ETDT_DISABLE)
'End Sub
'
'Public Function XP_OBJbIsVBRunning() As Boolean
'
'    'this function will check if vb is running..
'    'I made it public cause you maight need it in your project..
'
'    Debug.Assert (XP_OBJbCheckVB) = True   '-  this works only if vb is running..
'                                    '   the original purpose of the "assert"
'                                    '   function is to make sure that a function(or variable)
'                                    '   will return the specified value .. else, the
'                                    '   program will pause..
'
'                                    '-  what we do here is:
'                                    '   call another function, "XP_OBJbCheckVB"
'                                    '   through the "assert" function
'                                    '   that will set the value of the
'                                    '   "bIsVbRunning" to true if called..
'                                    '   ("Assert" will only call the function
'                                    '   when VB is running "Debug mode"..)
'
'                                    '   I Hope this would be useful somehow.. :D
'
'    XP_OBJbIsVBRunning = bIsVbRunning
'    bIsVbRunning = False
'
'End Function
'
'Private Function XP_OBJbIsWindowsXP() As Boolean
'    If Val(Trim(XP_OBJtGetWinVersion)) >= 5.01 Then
'        XP_OBJbIsWindowsXP = True
'    End If
'End Function
'
'Private Function XP_OBJtVb5Replace(Expression As String, Find As String, ReplaceWith As String, Optional Start As Long = 1, Optional Count As Long = -1, Optional Compare As VbCompareMethod = vbBinaryCompare) As String
'
'On Error GoTo ExitProcedure
'
'    'I wrote this function to add the "replace" capability to vb5
'    'so you can now use XP_OBJxMainXPStyle module in vb5
'    'NOTE: IF YOU USE VB6, YOU CAN COMMENT THIS FUNCTION..
'
'    Dim iFind As Long
'    Dim nextStart As Long
'    Dim sCount As Long
'
'    iFind = InStr(Start, Expression, Find, Compare)
'
'    nextStart = Start
'
'    If iFind = -1 Then
'        XP_OBJtVb5Replace = Expression
'        Exit Function
'    Else
'        Do
'            If sCount >= Count And Count <> -1 Then
'                Exit Do
'            End If
'            iFind = InStr(nextStart, Expression, Find, Compare)
'            If iFind = 0 Then Exit Do
'            sCount = sCount + 1
'            Expression = Left(Expression, iFind - 1) & ReplaceWith & Mid(Expression, iFind + Len(Find))
'            If nextStart < Len(Expression) Then
'                nextStart = iFind + Len(ReplaceWith) + 1
'            Else
'                Exit Do
'            End If
'
'        Loop
'    End If
'
'ExitProcedure:
'
'    XP_OBJtVb5Replace = Expression
'
'End Function
'
'' ______________________________________________________________________________________________________________
''|                                                                                                              |
''| This is the main function of this module..                                                                   |
''| This function will write the program manifest, restart the program(optional),                                |
''| delete the manifest(optional too)..                                                                          |
''|______________________________________________________________________________________________________________|
''|                                                                                                              |
''|   [ Parameter ]  +   [ Description ]                                                                         |
''|______________________________________________________________________________________________________________|
''|                  |                                                                                           |
''|   [Autorestart] -|    [True]: The program will automatically restart.  (the visual styles will be activated) |
''|                  |   [False]: The program will continue normally.      (will not be skinned till restarted)  |
''|                  |                                                                                           |
''|                  |           If set to [True], CreateNew will be set to [False] by default.                  |
''| -----------------|------------------------------------------------------------------------------------------ |
''|                  |                                                                                           |
''|   [Autohide   ] -|    [True]: The manifest file will not be visible.                                         |
''|                  |   [False]: The manifest file will remain.                                                 |
''|                  |                                                                                           |
''|                  |           If set to [True], CreateNew will be set to [False] by default.                  |
''| -----------------|------------------------------------------------------------------------------------------ |
''|                  |                                                                                           |
''|   [CreateNew  ] -|    [True]: Write a fresh manifest file.                                                   |
''|                  |   [False]: Nothing.                                                                       |
''|                  |                                                                                           |
''|                  |           Only for use with Autorestart=[False] Or Autohide=[False]                       |
''|__________________l___________________________________________________________________________________________|
''
'
'Public Function XP_OBJxMainXPStyle(Optional AutoRestart As Boolean = True, Optional Autohide As Boolean = True, Optional CreateNew As Boolean = False) As Boolean
'
'    If XP_OBJbIsWindowsXP = False Or XP_OBJbIsVBRunning Or XP_OBJbIsThemingSupported = False Then Exit Function
'    If IsThemeActiveXP = False Then Exit Function
'
'    Const IsVB6 As Boolean = True   'change to false if you are using vb5
'
'    On Error Resume Next
'
'    Dim XML             As String
'    Dim ManifestCheck   As String
'    Dim strManifest     As String
'    Dim FreeFileNo      As Integer
'
'    If AutoRestart = True Or Autohide = True Then CreateNew = False
'
'    '(put the XML into a string)
'    XML = ("<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?> " & vbCrLf & "<assembly " & vbCrLf & "   xmlns=""urn:schemas-microsoft-com:asm.v1"" " & vbCrLf & "   manifestVersion=""1.0"">" & vbCrLf & "<assemblyIdentity " & vbCrLf & "    processorArchitecture=""x86"" " & vbCrLf & "    version=""EXEVERSION""" & vbCrLf & "    type=""win32""" & vbCrLf & "    name=""COMPANYNAME.EXENAME""/>" & vbCrLf & "    <description>EXEDESCRIPTION</description>" & vbCrLf & "    <dependency>" & vbCrLf & "    <dependentAssembly>" & vbCrLf & "    <assemblyIdentity" & vbCrLf & "         type=""win32""" & vbCrLf & "         name=""Microsoft.Windows.Common-Controls""" & vbCrLf & "         version=""6.0.0.0""" & vbCrLf & "         publicKeyToken=""6595b64144ccf1df""" & vbCrLf & "         language=""*""" & vbCrLf & "         processorArchitecture=""x86""/>" & vbCrLf & "    </dependentAssembly>" & vbCrLf & "    </dependency>" & vbCrLf & "</assembly>" & vbCrLf & "")
'
'    strManifest = XP_OBJtAddDirSep(App.Path) & App.EXEName & ".exe.manifest"        'retrive the name of the manifest file
'
'    ManifestCheck = Dir(strManifest, vbNormal + vbSystem + vbHidden + vbReadOnly + vbArchive) 'look for the manifest file.
'
'    If ManifestCheck = "" Or CreateNew = True Then           'if not found.. or the "CreateNew" parameter is set to true, make a new one..
'
'        If IsVB6 = True Then                                 'see if vb is version 6 or not!
'            'VISUAL BASIC  6
'            XML = Replace(XML, "EXENAME", App.EXEName & ".exe")             'Replace the string "EXENAME" with the program's exe file name.
'            XML = Replace(XML, "EXEVERSION", App.Major & "." & App.Minor & "." & App.Revision & ".0") 'Replace the "EXEVERSION" string.
'            XML = Replace(XML, "EXEDESCRIPTION", App.FileDescription)       'Replace the app DESCRIPTION.
'            XML = Replace(XML, "COMPANYNAME", App.CompanyName)
'        Else
'            'VISUAL BASIC  5
'            XML = XP_OBJtVb5Replace(XML, "EXENAME", App.EXEName & ".exe")
'            XML = XP_OBJtVb5Replace(XML, "EXEVERSION", App.Major & "." & App.Minor & "." & App.Revision & ".0")
'            XML = XP_OBJtVb5Replace(XML, "EXEDESCRIPTION", App.FileDescription)
'            XML = XP_OBJtVb5Replace(XML, "COMPANYNAME", App.CompanyName)
'        End If
'
'        FreeFileNo = FreeFile                           'get the next free file
'        If ManifestCheck <> "" Then                     'this should be the "CreateNew" mode..
'            SetAttr strManifest, vbNormal
'            Kill (strManifest)
'        End If
'
'        Open strManifest For Binary As #(FreeFileNo)    'open the file in binary mode
'            Put #(FreeFileNo), , XML                    'use "put" to write to the file.. note that "put" (binary mode) is much faster than "print"(output mode)..
'        Close #(FreeFileNo)                             'close the file.
'        SetAttr strManifest, vbHidden + vbSystem        'set the file attributes to "Hidden, System"
'        XP_OBJxMainXPStyle = False                                 'return false.. this means that the application is not yet using visual styles..
'
'        If AutoRestart = True Then                      'if in automode (default), the program will restart.
'
'            Shell App.Path & "\" & App.EXEName & ".exe" & _
'            Space(1) & Command$, vbNormalFocus
'
'                                                        'restart the program and bypass command line parameters (if any)..
'            End                                         'end the session.
'        End If
'
'    Else                                                'the manifest file exists.
'
'
'        If Autohide = True Then                         'if the autohide mode is enabled (default), then we should delete the file..
'                                                        'in normal conditions, the manifest file will not appear at all ;)
'            SetAttr strManifest, vbNormal
'            Kill (strManifest)
'        End If
'
'        XP_OBJxMainXPStyle = True
'
'    End If
'        m_bIsManifestActive = XP_OBJxMainXPStyle
'End Function
'
'Public Property Get XP_OBJbIsManifestActive() As Boolean
'    XP_OBJbIsManifestActive = m_bIsManifestActive
'End Property
'
Public Function XP_APIbIsWindowsXP() As Boolean
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

    If Val(Trim(tWinVersion)) >= 5.01 Then
        XP_APIbIsWindowsXP = True
    End If
    Exit Function
ErrHandle:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Function
'
'


Public Sub XP_FRMxLoadThemeStyle(ByVal pfForm As Form, ByVal posfForm As SkinFramework)
'-----------------------------------------------------------
'   Call :
'   Cmt :
'   Developer : SOMJAI
'   Date Create : 06-09-2007 :  17:22:58
'-----------------------------------------------------------
    Dim tFolder$
    Dim oFSO As FileSystemObject
    
    On Error GoTo ErrHandle
    
    Set oFSO = New FileSystemObject
    
    'get ค่าจาก AdaIni.ada
    tFolder = Left(App.Path, InStrRev(App.Path, "\") - 1) + "\6.AdaTools\X-Resource\Styles"
    If Not SP_FLDbManFolder(tFolder, nEN_Fld1Exist) Then
        tFolder = Replace(tFolder, "1.AdaPos\", "")       'for develop
        If Not SP_FLDbManFolder(tFolder, nEN_Fld1Exist) Then
            tFolder = oFSO.GetParentFolderName(oFSO.GetParentFolderName(App.Path))
            tFolder = tFolder & "\6.AdaTools\X-Resource\Styles"
        End If
    End If

    posfForm.LoadSkin tFolder & "\Office2007.cjstyles", ""
    'posfForm.LoadSkin tFolder & "\Vista.cjstyles", ""
    'Vista.cjstyles
    posfForm.ApplyWindow pfForm.hwnd
    posfForm.ApplyOptions = posfForm.ApplyOptions Or xtpSkinApplyMetrics
    Exit Sub
ErrHandle:
    Call SP_MSGShowError
End Sub

'ReadRegistry(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\", "DefaultUserName")
Public Function ReadRegistry(ByVal Group As Long, ByVal Section As String, ByVal Key As String) As String
    Dim lResult As Long, lKeyValue As Long, lDataTypeValue As Long, lValueLength As Long, sValue As String, td As Double
On Error Resume Next
    lResult = RegOpenKey(Group, Section, lKeyValue)
    sValue = Space$(2048)
    lValueLength = Len(sValue)
    lResult = RegQueryValueEx(lKeyValue, Key, 0&, lDataTypeValue, sValue, lValueLength)
    If (lResult = 0) And (Err.Number = 0) Then
       If lDataTypeValue = REG_DWORD Then
          td = Asc(Mid$(sValue, 1, 1)) + &H100& * Asc(Mid$(sValue, 2, 1)) + &H10000 * Asc(Mid$(sValue, 3, 1)) + &H1000000 * CDbl(Asc(Mid$(sValue, 4, 1)))
          sValue = Format$(td, "000")
       End If
       If lDataTypeValue = REG_BINARY Then
           ' Return a binary field as a hex string (2 chars per byte)
           tStr2 = ""
           For i = 1 To lValueLength
              tStr1 = Hex(Asc(Mid(sValue, i, 1)))
              If Len(tStr1) = 1 Then tStr1 = "0" & tStr1
              tStr2 = tStr2 + tStr1
           Next
           sValue = tStr2
       Else
          sValue = Left$(sValue, lValueLength - 1)
       End If
    Else
       sValue = "Not Found"
    End If
    lResult = RegCloseKey(lKeyValue)
    ReadRegistry = sValue
End Function








