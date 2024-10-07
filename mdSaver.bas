Attribute VB_Name = "Saver"
Option Explicit

Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function PwdChangePassword& Lib "mpr" Alias "PwdChangePasswordA" (ByVal lpcRegkeyname$, ByVal hwnd&, ByVal uiReserved1&, ByVal uiReserved2&)
Public Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long

Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Const HWND_TOP = 0
Public Const GWL_HWNDPARENT = (-8)
Public Const GWL_STYLE = (-16)
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_NOZORDER = &H4
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const SWP_SHOWWINDOW = &H40
Public Const WS_CHILD = &H40000000
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2

Enum RunModes
    Configure = 1
    ScreenSaver
    Preview
    Password
End Enum

Const MODES_STR = "CSPA"

Global RunMode As RunModes

Public PreviewWindow As Long
Public Version As String

Public Const APP_NAME = "Primes Screen Saver"
Const SECTION = "Feuchtersoft"
Const HD_FTNAME = "Font Name"
Const HD_FTBOLD = "Font Bold"
Const HD_FTITALIC = "Font Italic"
Const HD_FTUNDERLINE = "Font Underline"
Const HD_FTSIZE = "Font Size"
Const HD_TXTINT = "Print Interval"
Const HD_MOVLEN = "Movement Speed"

' Runtime variables
Public FtName As String
Public FtBold As Boolean
Public FtItalic As Boolean
Public FtUnderline As Boolean
Public FtSize As Integer
Public TxtInt As Integer
Public MovLen As Integer

Sub GetSettings()
    ' Load runtime variables from the system registry
    FtName = GetSetting(SECTION, APP_NAME, HD_FTNAME, "MS Sans Serif")
    FtBold = CBool(GetSetting(SECTION, APP_NAME, HD_FTBOLD, "True"))
    FtItalic = CBool(GetSetting(SECTION, APP_NAME, HD_FTITALIC, "False"))
    FtUnderline = CBool(GetSetting(SECTION, APP_NAME, HD_FTUNDERLINE, "False"))
    FtSize = CInt(GetSetting(SECTION, APP_NAME, HD_FTSIZE, "14"))
    TxtInt = CInt(GetSetting(SECTION, APP_NAME, HD_TXTINT, "1"))
    MovLen = CInt(GetSetting(SECTION, APP_NAME, HD_MOVLEN, "1"))
End Sub

Public Sub SaveSettings()
    ' Save runtime variables to the system registry
    SaveSetting SECTION, APP_NAME, HD_FTNAME, FtName
    SaveSetting SECTION, APP_NAME, HD_FTBOLD, CStr(FtBold)
    SaveSetting SECTION, APP_NAME, HD_FTITALIC, CStr(FtItalic)
    SaveSetting SECTION, APP_NAME, HD_FTUNDERLINE, CStr(FtUnderline)
    SaveSetting SECTION, APP_NAME, HD_FTSIZE, CStr(FtSize)
    SaveSetting SECTION, APP_NAME, HD_TXTINT, CStr(TxtInt)
    SaveSetting SECTION, APP_NAME, HD_MOVLEN, CStr(MovLen)
End Sub

Sub Main()
Dim rctPreview As RECT
Dim WindowStyle As Long, rc As Long
    
    ' Determine how the screen saver is being loaded.
    rc = InStr(1, MODES_STR, Right(Left(UCase(Trim(Command)) + "  ", 2), 1))
    ' If no expected command line is provided, by default, load the
    ' screen saver in configure mode.
    If rc = 0 Then rc = 1
    
    RunMode = rc
    
    GetSettings
    
    Randomize Timer
    
    ' If running in preview or password modes, the command line
    ' contains information on the mother containing window.
    If RunMode = Preview Or RunMode = Password Then
        PreviewWindow = CLng(Right(Command, Len(Command) - 3))
    End If
    
    With frmMain
        Select Case RunMode
            Case Configure
                ' Configure mode; allows user to modify settings
                frmSetup.Show
            Case Preview
                ' Preview mode. Docks the screensaver form inside
                ' a window whose handle is contained in the command line
                ' arguments.
                GetClientRect PreviewWindow, rctPreview
                Load frmMain
                WindowStyle = GetWindowLong(.hwnd, GWL_STYLE)
                WindowStyle = (WindowStyle Or WS_CHILD)
                SetWindowLong .hwnd, GWL_STYLE, WindowStyle
                SetParent .hwnd, PreviewWindow
                SetWindowLong .hwnd, GWL_HWNDPARENT, PreviewWindow
                SetWindowPos .hwnd, HWND_TOP, 0&, 0&, rctPreview.Right, rctPreview.Bottom, SWP_NOZORDER Or SWP_NOACTIVATE Or SWP_SHOWWINDOW
            Case Password
                ' Password mode. Prompt user for password.
                On Error GoTo Error
                rc = PwdChangePassword("SCRSAVE", PreviewWindow, 0, 0)
            Case ScreenSaver
                ' Screensaver mode.
                CheckShouldRun
                .Show
        End Select
    End With
    Exit Sub

Error:
    MsgBox "Password could not be changed", vbOKOnly
    End
End Sub

' Checks if there is already an instance of the screensaver loaded.
Private Sub CheckShouldRun()
    If Not App.PrevInstance Then Exit Sub
    If FindWindow(vbNullString, APP_NAME) Then End
    frmMain.Caption = APP_NAME
End Sub

Public Function UsePassword() As Boolean
' Check wether a password has been used or not by checking the
' registry.
Dim lHandle As Long
Dim lResult As Long
Dim lValue As Long
    ' Earlier versions of windows do not store password information
    ' the same way Windows 9x does.
    lResult = RegOpenKeyEx(&H80000001, "Control Panel\Desktop", 0, 1, lHandle)
    If lResult = 0 Then
        lResult = RegQueryValueEx(lHandle, "ScreenSaveUsePassword", 0, 4, lValue, 32)
        If lResult = 0 Then
            ' Value does exist.
            UsePassword = lValue
            lResult = RegCloseKey(lHandle)
        End If
    End If
End Function

Public Sub ShowCurs(Optional Show As Boolean = True)
' Hides/Shows the curser
Dim sCursor As Long, sShow As Integer
    sShow = -Show * 2 - 1
    Do
        sCursor = ShowCursor(Show)
    Loop Until Abs(sCursor) * sShow = sCursor And sCursor <> 0
End Sub
