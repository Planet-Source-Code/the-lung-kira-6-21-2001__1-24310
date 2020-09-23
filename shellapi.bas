Attribute VB_Name = "shellapi"
Option Explicit


Public Declare Function Shell_NotifyIcon Lib "shell32.dll" (ByVal dwMessage As Long, ByRef pnid As NOTIFYICONDATA) As Boolean
Public Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hwnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
Public Declare Function SHEmptyRecycleBin Lib "shell32.dll" Alias "SHEmptyRecycleBinA" (ByVal hwnd As Long, ByVal pszRootPath As String, ByVal dwFlags As Long) As Long


Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const NIM_SETFOCUS = &H3
Public Const NIM_SETVERSION = &H4

Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4
Public Const NIF_STATE = &H8
Public Const NIF_INFO = &H10

Public Const NIS_HIDDEN = &H1
Public Const NIS_SHAREDICON = &H2

Public Const NIIF_NONE = &H0
Public Const NIIF_INFO = &H1
Public Const NIIF_WARNING = &H2
Public Const NIIF_ERROR = &H3

Public Const SHERB_NOCONFIRMATION = &H1
Public Const SHERB_NOPROGRESSUI = &H2
Public Const SHERB_NOSOUND = &H4


Public Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 128       'Version 5.0
    dwState As Long             'Version 5.0
    dwStateMask As Long         'Version 5.0
    szInfo As String * 256      'Version 5.0
    uTimeout As Long            'Version 5.0
    uVersion As Long            'Version 5.0
    szInfoTitle As String * 64  'Version 5.0
    dwInfoFlags As Long         'Version 5.0
End Type
