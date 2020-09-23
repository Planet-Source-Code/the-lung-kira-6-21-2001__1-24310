Attribute VB_Name = "winuser"
Option Explicit


Public Declare Function ChangeDisplaySettings Lib "user32.dll" Alias "ChangeDisplaySettingsA" (ByRef lpDevMode As DEVMODE, ByVal dwFlags As Long) As Long
Public Declare Function DefWindowProc Lib "user32.dll" Alias "DefWindowProcA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function DestroyWindow Lib "user32.dll" (ByVal hwnd As Long) As Boolean
Public Declare Function EnumDisplayDevices Lib "user32.dll" Alias "EnumDisplayDevicesA" (ByVal lpDevice As Long, ByVal iDevNum As Long, ByRef lpDisplayDevice As DISPLAY_DEVICE, ByVal dwFlags As Long) As Boolean
Public Declare Function EnumDisplayMonitors Lib "user32.dll" (ByVal hdc As Long, ByRef lprcClip As Any, ByVal lpfnEnum As Long, ByVal dwData As Long) As Boolean
Public Declare Function EnumDisplaySettings Lib "user32.dll" Alias "EnumDisplaySettingsA" (ByRef lpszDeviceName As Any, ByVal iModeNum As Long, ByRef lpDevMode As DEVMODE) As Boolean
Public Declare Function EnumDisplaySettingsEx Lib "user32.dll" Alias "EnumDisplaySettingsExA" (ByRef lpszDeviceName As Any, ByVal iModeNum As Long, ByRef lpDevMode As DEVMODE, ByVal dwFlags As Long) As Boolean
Public Declare Function EnumWindows Lib "user32.dll" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Boolean
Public Declare Function ExitWindowsEx Lib "user32.dll" (ByVal uFlags As Long, ByVal dwReserved As Long) As Boolean
Public Declare Function FlashWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal bInvert As Long) As Boolean
Public Declare Function GetAncestor Lib "user32.dll" (ByVal hwnd As Long, ByVal gaFlags As Long) As Long
Public Declare Function GetCaretBlinkTime Lib "user32.dll" () As Long
Public Declare Function GetCursorPos Lib "user32.dll" (ByRef lpPoint As POINTAPI) As Boolean
Public Declare Function GetDoubleClickTime Lib "user32.dll" () As Long
Public Declare Function GetGuiResources Lib "user32.dll" (ByVal hProcess As Long, ByVal uiFlags As Long) As Long
Public Declare Function GetKeyboardLayout Lib "user32.dll" (ByVal idThread As Long) As Long
Public Declare Function GetKeyboardLayoutName Lib "user32.dll" Alias "GetKeyboardLayoutNameA" (ByVal pwszKLID As String) As Boolean
Public Declare Function GetKeyboardType Lib "user32.dll" (ByVal nTypeFlag As Long) As Long
Public Declare Function GetMonitorInfo Lib "user32.dll" Alias "GetMonitorInfoA" (ByVal hMonitor As Long, ByRef lpmi As MONITORINFOEX) As Boolean
Public Declare Function GetSystemMetrics Lib "user32.dll" (ByVal nIndex As Long) As Long
Public Declare Function GetWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal uCmd As Long) As Long
Public Declare Function GetWindowInfo Lib "user32.dll" (ByVal hwnd As Long, ByRef pwi As WINDOWINFO) As Boolean
Public Declare Function GetWindowModuleFileName Lib "user32.dll" Alias "GetWindowModuleFileNameA" (ByVal hwnd As Long, ByVal lpszFileName As String, ByVal cchFileNameMax As Long) As Long
Public Declare Function GetWindowPlacement Lib "user32.dll" (ByVal hwnd As Long, ByRef lpwndpl As WINDOWPLACEMENT) As Boolean
Public Declare Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32.dll" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32.dll" (ByVal hwnd As Long, ByRef lpdwProcessId As Long) As Long
Public Declare Function IsWindowUnicode Lib "user32.dll" (ByVal hwnd As Long) As Boolean
Public Declare Function MessageBoxEx Lib "user32.dll" Alias "MessageBoxExA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal uType As Long, ByVal wLanguageId As Integer) As Long
Public Declare Function SendMessage Lib "user32.dll" (ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SetCaretBlinkTime Lib "user32.dll" (ByVal wMSeconds As Long) As Boolean
Public Declare Function SetCursorPos Lib "user32.dll" (ByVal X As Long, ByVal Y As Long) As Boolean
Public Declare Function SetDoubleClickTime Lib "user32.dll" (ByVal wCount As Long) As Boolean
Public Declare Function SetForegroundWindow Lib "user32.dll" (ByVal hwnd As Long) As Boolean
Public Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetWindowPlacement Lib "user32.dll" (ByVal hwnd As Long, ByRef lpwndpl As WINDOWPLACEMENT) As Boolean
Public Declare Function SetWindowPos Lib "user32.dll" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal uFlags As Long) As Boolean
Public Declare Function SetWindowText Lib "user32.dll" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Boolean
Public Declare Function SystemParametersInfo Lib "user32.dll" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Boolean


Public Const ARW_BOTTOMLEFT = &H0
Public Const ARW_BOTTOMRIGHT = &H1
Public Const ARW_TOPLEFT = &H2
Public Const ARW_TOPRIGHT = &H3
Public Const ARW_STARTMASK = &H3
Public Const ARW_STARTRIGHT = &H1
Public Const ARW_STARTTOP = &H2

Public Const ARW_LEFT = &H0
Public Const ARW_RIGHT = &H0
Public Const ARW_UP = &H4
Public Const ARW_DOWN = &H4
Public Const ARW_HIDE = &H8

Public Const ATF_TIMEOUTON = &H1
Public Const ATF_ONOFFFEEDBACK = &H2

Public Const CDS_UPDATEREGISTRY = &H1
Public Const CDS_TEST = &H2
Public Const CDS_FULLSCREEN = &H4
Public Const CDS_GLOBAL = &H8
Public Const CDS_SET_PRIMARY = &H10
Public Const CDS_VIDEOPARAMETERS = &H20
Public Const CDS_RESET = &H40000000
Public Const CDS_NORESET = &H10000000

Public Const DISP_CHANGE_SUCCESSFUL = 0
Public Const DISP_CHANGE_RESTART = 1
Public Const DISP_CHANGE_FAILED = -1
Public Const DISP_CHANGE_BADMODE = -2
Public Const DISP_CHANGE_NOTUPDATED = -3
Public Const DISP_CHANGE_BADFLAGS = -4
Public Const DISP_CHANGE_BADPARAM = -5
Public Const DISP_CHANGE_BADDUALVIEW = -6

Public Const EDS_RAWMODE = &H2

Public Const EWX_LOGOFF = 0
Public Const EWX_SHUTDOWN = &H1
Public Const EWX_REBOOT = &H2
Public Const EWX_FORCE = &H4
Public Const EWX_POWEROFF = &H8
Public Const EWX_FORCEIFHUNG = &H10

Public Const FKF_FILTERKEYSON = &H1
Public Const FKF_AVAILABLE = &H2
Public Const FKF_HOTKEYACTIVE = &H4
Public Const FKF_CONFIRMHOTKEY = &H8
Public Const FKF_HOTKEYSOUND = &H10
Public Const FKF_INDICATOR = &H20
Public Const FKF_CLICKON = &H40

Public Const GA_PARENT = 1
Public Const GA_ROOT = 2
Public Const GA_ROOTOWNER = 3

Public Const GR_GDIOBJECTS = 0
Public Const GR_USEROBJECTS = 1

Public Const GW_HWNDFIRST = 0
Public Const GW_HWNDLAST = 1
Public Const GW_HWNDNEXT = 2
Public Const GW_HWNDPREV = 3
Public Const GW_OWNER = 4
Public Const GW_CHILD = 5
Public Const GW_ENABLEDPOPUP = 6
Public Const GW_MAX = 6

Public Const GWL_WNDPROC = -4
Public Const GWL_HINSTANCE = -6
Public Const GWL_HWNDPARENT = -8
Public Const GWL_STYLE = -16
Public Const GWL_EXSTYLE = -20
Public Const GWL_USERDATA = -21
Public Const GWL_ID = -12

Public Const HWND_TOP = 0
Public Const HWND_BOTTOM = 1
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2

Public Const MB_OK = &H0
Public Const MB_OKCANCEL = &H1
Public Const MB_ABORTRETRYIGNORE = &H2
Public Const MB_YESNOCANCEL = &H3
Public Const MB_YESNO = &H4
Public Const MB_RETRYCANCEL = &H5
Public Const MB_CANCELTRYCONTINUE = &H6
Public Const MB_ICONHAND = &H10
Public Const MB_ICONQUESTION = &H20
Public Const MB_ICONEXCLAMATION = &H30
Public Const MB_ICONASTERISK = &H40
Public Const MB_USERICON = &H80
Public Const MB_ICONWARNING = MB_ICONEXCLAMATION
Public Const MB_ICONERROR = MB_ICONHAND
Public Const MB_ICONINFORMATION = MB_ICONASTERISK
Public Const MB_ICONSTOP = MB_ICONHAND
Public Const MB_DEFBUTTON1 = &H0
Public Const MB_DEFBUTTON2 = &H100
Public Const MB_DEFBUTTON3 = &H200
Public Const MB_DEFBUTTON4 = &H300
Public Const MB_APPLMODAL = &H0
Public Const MB_SYSTEMMODAL = &H1000
Public Const MB_TASKMODAL = &H2000
Public Const MB_HELP = &H4000
Public Const MB_NOFOCUS = &H8000
Public Const MB_SETFOREGROUND = &H10000
Public Const MB_DEFAULT_DESKTOP_ONLY = &H20000
Public Const MB_TOPMOST = &H40000
Public Const MB_RIGHT = &H80000
Public Const MB_RTLREADING = &H100000
Public Const MB_SERVICE_NOTIFICATION = &H200000
Public Const MB_SERVICE_NOTIFICATION_NT3X = &H40000
Public Const MB_TYPEMASK = &HF
Public Const MB_ICONMASK = &HF0
Public Const MB_DEFMASK = &HF00
Public Const MB_MODEMASK = &H3000
Public Const MB_MISCMASK = &HC000

Public Const MK_LBUTTON = &H1
Public Const MK_RBUTTON = &H2
Public Const MK_SHIFT = &H4
Public Const MK_CONTROL = &H8
Public Const MK_MBUTTON = &H10
Public Const MK_XBUTTON1 = &H20
Public Const MK_XBUTTON2 = &H40

Public Const MKF_MOUSEKEYSON = &H1
Public Const MKF_AVAILABLE = &H2
Public Const MKF_HOTKEYACTIVE = &H4
Public Const MKF_CONFIRMHOTKEY = &H8
Public Const MKF_HOTKEYSOUND = &H10
Public Const MKF_INDICATOR = &H20
Public Const MKF_MODIFIERS = &H40
Public Const MKF_REPLACENUMBERS = &H80
Public Const MKF_LEFTBUTTONSEL = &H10000000
Public Const MKF_RIGHTBUTTONSEL = &H20000000
Public Const MKF_LEFTBUTTONDOWN = &H1000000
Public Const MKF_RIGHTBUTTONDOWN = &H2000000
Public Const MKF_MOUSEMODE = &H80000000

Public Const MONITORINFOF_PRIMARY = &H1

Public Const SERKF_SERIALKEYSON = &H1
Public Const SERKF_AVAILABLE = &H2
Public Const SERKF_INDICATOR = &H4

Public Const SKF_STICKYKEYSON = &H1
Public Const SKF_AVAILABLE = &H2
Public Const SKF_HOTKEYACTIVE = &H4
Public Const SKF_CONFIRMHOTKEY = &H8
Public Const SKF_HOTKEYSOUND = &H10
Public Const SKF_INDICATOR = &H20
Public Const SKF_AUDIBLEFEEDBACK = &H40
Public Const SKF_TRISTATE = &H80
Public Const SKF_TWOKEYSOFF = &H100
Public Const SKF_LALTLATCHED = &H10000000
Public Const SKF_LCTLLATCHED = &H4000000
Public Const SKF_LSHIFTLATCHED = &H1000000
Public Const SKF_RALTLATCHED = &H20000000
Public Const SKF_RCTLLATCHED = &H8000000
Public Const SKF_RSHIFTLATCHED = &H2000000
Public Const SKF_LWINLATCHED = &H40000000
Public Const SKF_RWINLATCHED = &H80000000
Public Const SKF_LALTLOCKED = &H100000
Public Const SKF_LCTLLOCKED = &H40000
Public Const SKF_LSHIFTLOCKED = &H10000
Public Const SKF_RALTLOCKED = &H200000
Public Const SKF_RCTLLOCKED = &H80000
Public Const SKF_RSHIFTLOCKED = &H20000
Public Const SKF_LWINLOCKED = &H400000
Public Const SKF_RWINLOCKED = &H800000

Public Const SM_CXSCREEN = 0
Public Const SM_CYSCREEN = 1
Public Const SM_CXVSCROLL = 2
Public Const SM_CYHSCROLL = 3
Public Const SM_CYCAPTION = 4
Public Const SM_CXBORDER = 5
Public Const SM_CYBORDER = 6
Public Const SM_CXDLGFRAME = 7
Public Const SM_CYDLGFRAME = 8
Public Const SM_CYVTHUMB = 9
Public Const SM_CXHTHUMB = 10
Public Const SM_CXICON = 11
Public Const SM_CYICON = 12
Public Const SM_CXCURSOR = 13
Public Const SM_CYCURSOR = 14
Public Const SM_CYMENU = 15
Public Const SM_CXFULLSCREEN = 16
Public Const SM_CYFULLSCREEN = 17
Public Const SM_CYKANJIWINDOW = 18
Public Const SM_MOUSEPRESENT = 19
Public Const SM_CYVSCROLL = 20
Public Const SM_CXHSCROLL = 21
Public Const SM_DEBUG = 22
Public Const SM_SWAPBUTTON = 23
Public Const SM_RESERVED1 = 24
Public Const SM_RESERVED2 = 25
Public Const SM_RESERVED3 = 26
Public Const SM_RESERVED4 = 27
Public Const SM_CXMIN = 28
Public Const SM_CYMIN = 29
Public Const SM_CXSIZE = 30
Public Const SM_CYSIZE = 31
Public Const SM_CXFRAME = 32
Public Const SM_CYFRAME = 33
Public Const SM_CXMINTRACK = 34
Public Const SM_CYMINTRACK = 35
Public Const SM_CXDOUBLECLK = 36
Public Const SM_CYDOUBLECLK = 37
Public Const SM_CXICONSPACING = 38
Public Const SM_CYICONSPACING = 39
Public Const SM_MENUDROPALIGNMENT = 40
Public Const SM_PENWINDOWS = 41
Public Const SM_DBCSENABLED = 42
Public Const SM_CMOUSEBUTTONS = 43
Public Const SM_CXFIXEDFRAME = SM_CXDLGFRAME
Public Const SM_CYFIXEDFRAME = SM_CYDLGFRAME
Public Const SM_CXSIZEFRAME = SM_CXFRAME
Public Const SM_CYSIZEFRAME = SM_CYFRAME
Public Const SM_SECURE = 44
Public Const SM_CXEDGE = 45
Public Const SM_CYEDGE = 46
Public Const SM_CXMINSPACING = 47
Public Const SM_CYMINSPACING = 48
Public Const SM_CXSMICON = 49
Public Const SM_CYSMICON = 50
Public Const SM_CYSMCAPTION = 51
Public Const SM_CXSMSIZE = 52
Public Const SM_CYSMSIZE = 53
Public Const SM_CXMENUSIZE = 54
Public Const SM_CYMENUSIZE = 55
Public Const SM_ARRANGE = 56
Public Const SM_CXMINIMIZED = 57
Public Const SM_CYMINIMIZED = 58
Public Const SM_CXMAXTRACK = 59
Public Const SM_CYMAXTRACK = 60
Public Const SM_CXMAXIMIZED = 61
Public Const SM_CYMAXIMIZED = 62
Public Const SM_NETWORK = 63
Public Const SM_CLEANBOOT = 67
Public Const SM_CXDRAG = 68
Public Const SM_CYDRAG = 69
Public Const SM_SHOWSOUNDS = 70
Public Const SM_CXMENUCHECK = 71
Public Const SM_CYMENUCHECK = 72
Public Const SM_SLOWMACHINE = 73
Public Const SM_MIDEASTENABLED = 74
Public Const SM_MOUSEWHEELPRESENT = 75
Public Const SM_XVIRTUALSCREEN = 76
Public Const SM_YVIRTUALSCREEN = 77
Public Const SM_CXVIRTUALSCREEN = 78
Public Const SM_CYVIRTUALSCREEN = 79
Public Const SM_CMONITORS = 80
Public Const SM_SAMEDISPLAYFORMAT = 81
Public Const SM_IMMENABLED = 82
Public Const SM_CXFOCUSBORDER = 83
Public Const SM_CYFOCUSBORDER = 84
Public Const SM_CMETRICS = 86
Public Const SM_REMOTESESSION = &H1000
Public Const SM_SHUTTINGDOWN = &H2000

Public Const SPI_GETBEEP = 1
Public Const SPI_SETBEEP = 2
Public Const SPI_GETMOUSE = 3
Public Const SPI_SETMOUSE = 4
Public Const SPI_GETBORDER = 5
Public Const SPI_SETBORDER = 6
Public Const SPI_GETKEYBOARDSPEED = 10
Public Const SPI_SETKEYBOARDSPEED = 11
Public Const SPI_LANGDRIVER = 12
Public Const SPI_ICONHORIZONTALSPACING = 13
Public Const SPI_GETSCREENSAVETIMEOUT = 14
Public Const SPI_SETSCREENSAVETIMEOUT = 15
Public Const SPI_GETSCREENSAVEACTIVE = 16
Public Const SPI_SETSCREENSAVEACTIVE = 17
Public Const SPI_GETGRIDGRANULARITY = 18
Public Const SPI_SETGRIDGRANULARITY = 19
Public Const SPI_SETDESKWALLPAPER = 20
Public Const SPI_SETDESKPATTERN = 21
Public Const SPI_GETKEYBOARDDELAY = 22
Public Const SPI_SETKEYBOARDDELAY = 23
Public Const SPI_ICONVERTICALSPACING = 24
Public Const SPI_GETICONTITLEWRAP = 25
Public Const SPI_SETICONTITLEWRAP = 26
Public Const SPI_GETMENUDROPALIGNMENT = 27
Public Const SPI_SETMENUDROPALIGNMENT = 28
Public Const SPI_SETDOUBLECLKWIDTH = 29
Public Const SPI_SETDOUBLECLKHEIGHT = 30
Public Const SPI_GETICONTITLELOGFONT = 31
Public Const SPI_SETDOUBLECLICKTIME = 32
Public Const SPI_SETMOUSEBUTTONSWAP = 33
Public Const SPI_SETICONTITLELOGFONT = 34
Public Const SPI_GETFASTTASKSWITCH = 35
Public Const SPI_SETFASTTASKSWITCH = 36
Public Const SPI_SETDRAGFULLWINDOWS = 37
Public Const SPI_GETDRAGFULLWINDOWS = 38
Public Const SPI_GETNONCLIENTMETRICS = 41
Public Const SPI_SETNONCLIENTMETRICS = 42
Public Const SPI_GETMINIMIZEDMETRICS = 43
Public Const SPI_SETMINIMIZEDMETRICS = 44
Public Const SPI_GETICONMETRICS = 45
Public Const SPI_SETICONMETRICS = 46
Public Const SPI_SETWORKAREA = 47
Public Const SPI_GETWORKAREA = 48
Public Const SPI_SETPENWINDOWS = 49
Public Const SPI_GETHIGHCONTRAST = 66
Public Const SPI_SETHIGHCONTRAST = 67
Public Const SPI_GETKEYBOARDPREF = 68
Public Const SPI_SETKEYBOARDPREF = 69
Public Const SPI_GETSCREENREADER = 70
Public Const SPI_SETSCREENREADER = 71
Public Const SPI_GETANIMATION = 72
Public Const SPI_SETANIMATION = 73
Public Const SPI_GETFONTSMOOTHING = 74
Public Const SPI_SETFONTSMOOTHING = 75
Public Const SPI_SETDRAGWIDTH = 76
Public Const SPI_SETDRAGHEIGHT = 77
Public Const SPI_SETHANDHELD = 78
Public Const SPI_GETLOWPOWERTIMEOUT = 79
Public Const SPI_GETPOWEROFFTIMEOUT = 80
Public Const SPI_SETLOWPOWERTIMEOUT = 81
Public Const SPI_SETPOWEROFFTIMEOUT = 82
Public Const SPI_GETLOWPOWERACTIVE = 83
Public Const SPI_GETPOWEROFFACTIVE = 84
Public Const SPI_SETLOWPOWERACTIVE = 85
Public Const SPI_SETPOWEROFFACTIVE = 86
Public Const SPI_SETCURSORS = 87
Public Const SPI_SETICONS = 88
Public Const SPI_GETDEFAULTINPUTLANG = 89
Public Const SPI_SETDEFAULTINPUTLANG = 90
Public Const SPI_SETLANGTOGGLE = 91
Public Const SPI_GETWINDOWSEXTENSION = 92
Public Const SPI_SETMOUSETRAILS = 93
Public Const SPI_GETMOUSETRAILS = 94
Public Const SPI_SETSCREENSAVERRUNNING = 97
Public Const SPI_SCREENSAVERRUNNING = SPI_SETSCREENSAVERRUNNING
Public Const SPI_GETFILTERKEYS = 50
Public Const SPI_SETFILTERKEYS = 51
Public Const SPI_GETTOGGLEKEYS = 52
Public Const SPI_SETTOGGLEKEYS = 53
Public Const SPI_GETMOUSEKEYS = 54
Public Const SPI_SETMOUSEKEYS = 55
Public Const SPI_GETSHOWSOUNDS = 56
Public Const SPI_SETSHOWSOUNDS = 57
Public Const SPI_GETSTICKYKEYS = 58
Public Const SPI_SETSTICKYKEYS = 59
Public Const SPI_GETACCESSTIMEOUT = 60
Public Const SPI_SETACCESSTIMEOUT = 61
Public Const SPI_GETSERIALKEYS = 62
Public Const SPI_SETSERIALKEYS = 63
Public Const SPI_GETSOUNDSENTRY = 64
Public Const SPI_SETSOUNDSENTRY = 65
Public Const SPI_GETSNAPTODEFBUTTON = 95
Public Const SPI_SETSNAPTODEFBUTTON = 96
Public Const SPI_GETMOUSEHOVERWIDTH = 98
Public Const SPI_SETMOUSEHOVERWIDTH = 99
Public Const SPI_GETMOUSEHOVERHEIGHT = 100
Public Const SPI_SETMOUSEHOVERHEIGHT = 101
Public Const SPI_GETMOUSEHOVERTIME = 102
Public Const SPI_SETMOUSEHOVERTIME = 103
Public Const SPI_GETWHEELSCROLLLINES = 104
Public Const SPI_SETWHEELSCROLLLINES = 105
Public Const SPI_GETMENUSHOWDELAY = 106
Public Const SPI_SETMENUSHOWDELAY = 107
Public Const SPI_GETSHOWIMEUI = 110
Public Const SPI_SETSHOWIMEUI = 111
Public Const SPI_GETMOUSESPEED = 112
Public Const SPI_SETMOUSESPEED = 113
Public Const SPI_GETSCREENSAVERRUNNING = 114
Public Const SPI_GETDESKWALLPAPER = 115
Public Const SPI_GETACTIVEWINDOWTRACKING = &H1000
Public Const SPI_SETACTIVEWINDOWTRACKING = &H1001
Public Const SPI_GETMENUANIMATION = &H1002
Public Const SPI_SETMENUANIMATION = &H1003
Public Const SPI_GETCOMBOBOXANIMATION = &H1004
Public Const SPI_SETCOMBOBOXANIMATION = &H1005
Public Const SPI_GETLISTBOXSMOOTHSCROLLING = &H1006
Public Const SPI_SETLISTBOXSMOOTHSCROLLING = &H1007
Public Const SPI_GETGRADIENTCAPTIONS = &H1008
Public Const SPI_SETGRADIENTCAPTIONS = &H1009
Public Const SPI_GETKEYBOARDCUES = &H100A
Public Const SPI_SETKEYBOARDCUES = &H100B
Public Const SPI_GETMENUUNDERLINES = SPI_GETKEYBOARDCUES
Public Const SPI_SETMENUUNDERLINES = SPI_SETKEYBOARDCUES
Public Const SPI_GETACTIVEWNDTRKZORDER = &H100C
Public Const SPI_SETACTIVEWNDTRKZORDER = &H100D
Public Const SPI_GETHOTTRACKING = &H100E
Public Const SPI_SETHOTTRACKING = &H100F
Public Const SPI_GETMENUFADE = &H1012
Public Const SPI_SETMENUFADE = &H1013
Public Const SPI_GETSELECTIONFADE = &H1014
Public Const SPI_SETSELECTIONFADE = &H1015
Public Const SPI_GETTOOLTIPANIMATION = &H1016
Public Const SPI_SETTOOLTIPANIMATION = &H1017
Public Const SPI_GETTOOLTIPFADE = &H1018
Public Const SPI_SETTOOLTIPFADE = &H1019
Public Const SPI_GETCURSORSHADOW = &H101A
Public Const SPI_SETCURSORSHADOW = &H101B
Public Const SPI_GETUIEFFECTS = &H103E
Public Const SPI_SETUIEFFECTS = &H103F
Public Const SPI_GETFOREGROUNDLOCKTIMEOUT = &H2000
Public Const SPI_SETFOREGROUNDLOCKTIMEOUT = &H2001
Public Const SPI_GETACTIVEWNDTRKTIMEOUT = &H2002
Public Const SPI_SETACTIVEWNDTRKTIMEOUT = &H2003
Public Const SPI_GETFOREGROUNDFLASHCOUNT = &H2004
Public Const SPI_SETFOREGROUNDFLASHCOUNT = &H2005
Public Const SPI_GETCARETWIDTH = &H2006
Public Const SPI_SETCARETWIDTH = &H2007
Public Const SPI_GETMOUSECLICKLOCKTIME = &H2008
Public Const SPI_SETMOUSECLICKLOCKTIME = &H2009
Public Const SPI_GETFONTSMOOTHINGTYPE = &H200A
Public Const SPI_SETFONTSMOOTHINGTYPE = &H200B

Public Const SPIF_UPDATEINIFILE = &H1
Public Const SPIF_SENDWININICHANGE = &H2
Public Const SPIF_SENDCHANGE = SPIF_SENDWININICHANGE

Public Const SSF_SOUNDSENTRYON = &H1
Public Const SSF_AVAILABLE = &H2
Public Const SSF_INDICATOR = &H4

Public Const SSGF_NONE = 0
Public Const SSGF_DISPLAY = 3

Public Const SSTF_NONE = 0
Public Const SSTF_CHARS = 1
Public Const SSTF_BORDER = 2
Public Const SSTF_DISPLAY = 3

Public Const SSWF_NONE = 0
Public Const SSWF_TITLE = 1
Public Const SSWF_WINDOW = 2
Public Const SSWF_DISPLAY = 3
Public Const SSWF_CUSTOM = 4

Public Const SW_HIDE = 0
Public Const SW_SHOWNORMAL = 1
Public Const SW_NORMAL = 1
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_MAXIMIZE = 3
Public Const SW_SHOWNOACTIVATE = 4
Public Const SW_SHOW = 5
Public Const SW_MINIMIZE = 6
Public Const SW_SHOWMINNOACTIVE = 7
Public Const SW_SHOWNA = 8
Public Const SW_RESTORE = 9
Public Const SW_SHOWDEFAULT = 10
Public Const SW_FORCEMINIMIZE = 11
Public Const SW_MAX = 11

Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOZORDER = &H4
Public Const SWP_NOREDRAW = &H8
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_FRAMECHANGED = &H20
Public Const SWP_SHOWWINDOW = &H40
Public Const SWP_HIDEWINDOW = &H80
Public Const SWP_NOCOPYBITS = &H100
Public Const SWP_NOOWNERZORDER = &H200
Public Const SWP_NOSENDCHANGING = &H400
Public Const SWP_DRAWFRAME = SWP_FRAMECHANGED
Public Const SWP_NOREPOSITION = SWP_NOOWNERZORDER
Public Const SWP_DEFERERASE = &H2000
Public Const SWP_ASYNCWINDOWPOS = &H4000

Public Const TKF_TOGGLEKEYSON = &H1
Public Const TKF_AVAILABLE = &H2
Public Const TKF_HOTKEYACTIVE = &H4
Public Const TKF_CONFIRMHOTKEY = &H8
Public Const TKF_HOTKEYSOUND = &H10
Public Const TKF_INDICATOR = &H20

Public Const WM_NULL = &H0
Public Const WM_CREATE = &H1
Public Const WM_DESTROY = &H2
Public Const WM_MOVE = &H3
Public Const WM_SIZE = &H5
Public Const WM_ACTIVATE = &H6
Public Const WM_SETFOCUS = &H7
Public Const WM_KILLFOCUS = &H8
Public Const WM_ENABLE = &HA
Public Const WM_SETREDRAW = &HB
Public Const WM_SETTEXT = &HC
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_PAINT = &HF
Public Const WM_CLOSE = &H10
Public Const WM_QUERYENDSESSION = &H11
Public Const WM_QUERYOPEN = &H13
Public Const WM_ENDSESSION = &H16
Public Const WM_QUIT = &H12
Public Const WM_ERASEBKGND = &H14
Public Const WM_SYSCOLORCHANGE = &H15
Public Const WM_SHOWWINDOW = &H18
Public Const WM_WININICHANGE = &H1A
Public Const WM_SETTINGCHANGE = WM_WININICHANGE
Public Const WM_DEVMODECHANGE = &H1B
Public Const WM_ACTIVATEAPP = &H1C
Public Const WM_FONTCHANGE = &H1D
Public Const WM_TIMECHANGE = &H1E
Public Const WM_CANCELMODE = &H1F
Public Const WM_SETCURSOR = &H20
Public Const WM_MOUSEACTIVATE = &H21
Public Const WM_CHILDACTIVATE = &H22
Public Const WM_QUEUESYNC = &H23
Public Const WM_GETMINMAXINFO = &H24
Public Const WM_PAINTICON = &H26
Public Const WM_ICONERASEBKGND = &H27
Public Const WM_NEXTDLGCTL = &H28
Public Const WM_SPOOLERSTATUS = &H2A
Public Const WM_DRAWITEM = &H2B
Public Const WM_MEASUREITEM = &H2C
Public Const WM_DELETEITEM = &H2D
Public Const WM_VKEYTOITEM = &H2E
Public Const WM_CHARTOITEM = &H2F
Public Const WM_SETFONT = &H30
Public Const WM_GETFONT = &H31
Public Const WM_SETHOTKEY = &H32
Public Const WM_GETHOTKEY = &H33
Public Const WM_QUERYDRAGICON = &H37
Public Const WM_COMPAREITEM = &H39
Public Const WM_GETOBJECT = &H3D
Public Const WM_COMPACTING = &H41
Public Const WM_COMMNOTIFY = &H44
Public Const WM_WINDOWPOSCHANGING = &H46
Public Const WM_WINDOWPOSCHANGED = &H47
Public Const WM_POWER = &H48
Public Const WM_COPYDATA = &H4A
Public Const WM_CANCELJOURNAL = &H4B
Public Const WM_NOTIFY = &H4E
Public Const WM_INPUTLANGCHANGEREQUEST = &H50
Public Const WM_INPUTLANGCHANGE = &H51
Public Const WM_TCARD = &H52
Public Const WM_HELP = &H53
Public Const WM_USERCHANGED = &H54
Public Const WM_NOTIFYFORMAT = &H55
Public Const WM_CONTEXTMENU = &H7B
Public Const WM_STYLECHANGING = &H7C
Public Const WM_STYLECHANGED = &H7D
Public Const WM_DISPLAYCHANGE = &H7E
Public Const WM_GETICON = &H7F
Public Const WM_SETICON = &H80
Public Const WM_NCCREATE = &H81
Public Const WM_NCDESTROY = &H82
Public Const WM_NCCALCSIZE = &H83
Public Const WM_NCHITTEST = &H84
Public Const WM_NCPAINT = &H85
Public Const WM_NCACTIVATE = &H86
Public Const WM_GETDLGCODE = &H87
Public Const WM_SYNCPAINT = &H88
Public Const WM_NCMOUSEMOVE = &HA0
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const WM_NCLBUTTONUP = &HA2
Public Const WM_NCLBUTTONDBLCLK = &HA3
Public Const WM_NCRBUTTONDOWN = &HA4
Public Const WM_NCRBUTTONUP = &HA5
Public Const WM_NCRBUTTONDBLCLK = &HA6
Public Const WM_NCMBUTTONDOWN = &HA7
Public Const WM_NCMBUTTONUP = &HA8
Public Const WM_NCMBUTTONDBLCLK = &HA9
Public Const WM_NCXBUTTONDOWN = &HAB
Public Const WM_NCXBUTTONUP = &HAC
Public Const WM_NCXBUTTONDBLCLK = &HAD
Public Const WM_INPUT = &HFF
Public Const WM_KEYFIRST = &H100
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_CHAR = &H102
Public Const WM_DEADCHAR = &H103
Public Const WM_SYSKEYDOWN = &H104
Public Const WM_SYSKEYUP = &H105
Public Const WM_SYSCHAR = &H106
Public Const WM_SYSDEADCHAR = &H107
Public Const WM_KEYLAST = &H108
Public Const WM_IME_STARTCOMPOSITION = &H10D
Public Const WM_IME_ENDCOMPOSITION = &H10E
Public Const WM_IME_COMPOSITION = &H10F
Public Const WM_IME_KEYLAST = &H10F
Public Const WM_INITDIALOG = &H110
Public Const WM_COMMAND = &H111
Public Const WM_SYSCOMMAND = &H112
Public Const WM_TIMER = &H113
Public Const WM_HSCROLL = &H114
Public Const WM_VSCROLL = &H115
Public Const WM_INITMENU = &H116
Public Const WM_INITMENUPOPUP = &H117
Public Const WM_MENUSELECT = &H11F
Public Const WM_MENUCHAR = &H120
Public Const WM_ENTERIDLE = &H121
Public Const WM_MENURBUTTONUP = &H122
Public Const WM_MENUDRAG = &H123
Public Const WM_MENUGETOBJECT = &H124
Public Const WM_UNINITMENUPOPUP = &H125
Public Const WM_MENUCOMMAND = &H126
Public Const WM_CHANGEUISTATE = &H127
Public Const WM_UPDATEUISTATE = &H128
Public Const WM_QUERYUISTATE = &H129
Public Const WM_CTLCOLORMSGBOX = &H132
Public Const WM_CTLCOLOREDIT = &H133
Public Const WM_CTLCOLORLISTBOX = &H134
Public Const WM_CTLCOLORBTN = &H135
Public Const WM_CTLCOLORDLG = &H136
Public Const WM_CTLCOLORSCROLLBAR = &H137
Public Const WM_CTLCOLORSTATIC = &H138
Public Const WM_MOUSEFIRST = &H200
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_RBUTTONDBLCLK = &H206
Public Const WM_MBUTTONDOWN = &H207
Public Const WM_MBUTTONUP = &H208
Public Const WM_MBUTTONDBLCLK = &H209
Public Const WM_MOUSEWHEEL = &H20A
Public Const WM_XBUTTONDOWN = &H20B
Public Const WM_XBUTTONUP = &H20C
Public Const WM_XBUTTONDBLCLK = &H20D

'If (_WIN32_WINNT >= &H0500)
'Public Const  WM_MOUSELAST           =         &H020D
'ElseIf (WIN32_WINNT >= &H0400) Or (WIN32_WINDOWS > &H0400)
'Public Const  WM_MOUSELAST           =         &H020A
'Else
'Public Const  WM_MOUSELAST           =         &H0209
'End If

Public Const WM_PARENTNOTIFY = &H210
Public Const WM_ENTERMENULOOP = &H211
Public Const WM_EXITMENULOOP = &H212
Public Const WM_NEXTMENU = &H213
Public Const WM_SIZING = &H214
Public Const WM_CAPTURECHANGED = &H215
Public Const WM_MOVING = &H216
Public Const WM_POWERBROADCAST = &H218
Public Const WM_DEVICECHANGE = &H219
Public Const WM_MDICREATE = &H220
Public Const WM_MDIDESTROY = &H221
Public Const WM_MDIACTIVATE = &H222
Public Const WM_MDIRESTORE = &H223
Public Const WM_MDINEXT = &H224
Public Const WM_MDIMAXIMIZE = &H225
Public Const WM_MDITILE = &H226
Public Const WM_MDICASCADE = &H227
Public Const WM_MDIICONARRANGE = &H228
Public Const WM_MDIGETACTIVE = &H229
Public Const WM_MDISETMENU = &H230
Public Const WM_ENTERSIZEMOVE = &H231
Public Const WM_EXITSIZEMOVE = &H232
Public Const WM_DROPFILES = &H233
Public Const WM_MDIREFRESHMENU = &H234
Public Const WM_IME_SETCONTEXT = &H281
Public Const WM_IME_NOTIFY = &H282
Public Const WM_IME_CONTROL = &H283
Public Const WM_IME_COMPOSITIONFULL = &H284
Public Const WM_IME_SELECT = &H285
Public Const WM_IME_CHAR = &H286
Public Const WM_IME_REQUEST = &H288
Public Const WM_IME_KEYDOWN = &H290
Public Const WM_IME_KEYUP = &H291
Public Const WM_MOUSEHOVER = &H2A1
Public Const WM_MOUSELEAVE = &H2A3
Public Const WM_NCMOUSEHOVER = &H2A0
Public Const WM_NCMOUSELEAVE = &H2A2
Public Const WM_WTSSESSION_CHANGE = &H2B1
Public Const WM_TABLET_FIRST = &H2C0
Public Const WM_TABLET_LAST = &H2DF
Public Const WM_CUT = &H300
Public Const WM_COPY = &H301
Public Const WM_PASTE = &H302
Public Const WM_CLEAR = &H303
Public Const WM_UNDO = &H304
Public Const WM_RENDERFORMAT = &H305
Public Const WM_RENDERALLFORMATS = &H306
Public Const WM_DESTROYCLIPBOARD = &H307
Public Const WM_DRAWCLIPBOARD = &H308
Public Const WM_PAINTCLIPBOARD = &H309
Public Const WM_VSCROLLCLIPBOARD = &H30A
Public Const WM_SIZECLIPBOARD = &H30B
Public Const WM_ASKCBFORMATNAME = &H30C
Public Const WM_CHANGECBCHAIN = &H30D
Public Const WM_HSCROLLCLIPBOARD = &H30E
Public Const WM_QUERYNEWPALETTE = &H30F
Public Const WM_PALETTEISCHANGING = &H310
Public Const WM_PALETTECHANGED = &H311
Public Const WM_HOTKEY = &H312
Public Const WM_PRINT = &H317
Public Const WM_PRINTCLIENT = &H318
Public Const WM_APPCOMMAND = &H319
Public Const WM_THEMECHANGED = &H31A
Public Const WM_HANDHELDFIRST = &H358
Public Const WM_HANDHELDLAST = &H35F
Public Const WM_AFXFIRST = &H360
Public Const WM_AFXLAST = &H37F
Public Const WM_PENWINFIRST = &H380
Public Const WM_PENWINLAST = &H38F
Public Const WM_APP = &H8000
Public Const WM_USER = &H400

Public Const WPF_SETMINPOSITION = &H1
Public Const WPF_RESTORETOMAXIMIZED = &H2
Public Const WPF_ASYNCWINDOWPLACEMENT = &H4

Public Const WS_OVERLAPPED = &H0
Public Const WS_POPUP = &H80000000
Public Const WS_CHILD = &H40000000
Public Const WS_MINIMIZE = &H20000000
Public Const WS_VISIBLE = &H10000000
Public Const WS_DISABLED = &H8000000
Public Const WS_CLIPSIBLINGS = &H4000000
Public Const WS_CLIPCHILDREN = &H2000000
Public Const WS_MAXIMIZE = &H1000000
Public Const WS_CAPTION = &HC00000
Public Const WS_BORDER = &H800000
Public Const WS_DLGFRAME = &H400000
Public Const WS_VSCROLL = &H200000
Public Const WS_HSCROLL = &H100000
Public Const WS_SYSMENU = &H80000
Public Const WS_THICKFRAME = &H40000
Public Const WS_GROUP = &H20000
Public Const WS_TABSTOP = &H10000
Public Const WS_MINIMIZEBOX = &H20000
Public Const WS_MAXIMIZEBOX = &H10000
Public Const WS_TILED = WS_OVERLAPPED
Public Const WS_ICONIC = WS_MINIMIZE
Public Const WS_SIZEBOX = WS_THICKFRAME
Public Const WS_OVERLAPPEDWINDOW = (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)
Public Const WS_POPUPWINDOW = (WS_POPUP Or WS_BORDER Or WS_SYSMENU)
Public Const WS_CHILDWINDOW = WS_CHILD
Public Const WS_TILEDWINDOW = WS_OVERLAPPEDWINDOW

Public Const WS_EX_DLGMODALFRAME = &H1
Public Const WS_EX_NOPARENTNOTIFY = &H4
Public Const WS_EX_TOPMOST = &H8
Public Const WS_EX_ACCEPTFILES = &H10
Public Const WS_EX_TRANSPARENT = &H20
Public Const WS_EX_MDICHILD = &H40
Public Const WS_EX_TOOLWINDOW = &H80
Public Const WS_EX_WINDOWEDGE = &H100
Public Const WS_EX_CLIENTEDGE = &H200
Public Const WS_EX_CONTEXTHELP = &H400
Public Const WS_EX_RIGHT = &H1000
Public Const WS_EX_LEFT = &H0
Public Const WS_EX_RTLREADING = &H2000
Public Const WS_EX_LTRREADING = &H0
Public Const WS_EX_LEFTSCROLLBAR = &H4000
Public Const WS_EX_RIGHTSCROLLBAR = &H0
Public Const WS_EX_CONTROLPARENT = &H10000
Public Const WS_EX_STATICEDGE = &H20000
Public Const WS_EX_APPWINDOW = &H40000
Public Const WS_EX_OVERLAPPEDWINDOW = (WS_EX_WINDOWEDGE Or WS_EX_CLIENTEDGE)
Public Const WS_EX_PALETTEWINDOW = (WS_EX_WINDOWEDGE Or WS_EX_TOOLWINDOW Or WS_EX_TOPMOST)
Public Const WS_EX_LAYERED = &H80000
Public Const WS_EX_NOINHERITLAYOUT = &H100000
Public Const WS_EX_LAYOUTRTL = &H400000
Public Const WS_EX_COMPOSITED = &H2000000
Public Const WS_EX_NOACTIVATE = &H8000000

Public Const XBUTTON1 = &H1
Public Const XBUTTON2 = &H2


Public Type ACCESSTIMEOUT
    cbSize As Long
    dwFlags As Long
    iTimeOutMSec As Long
End Type

Public Type FILTERKEYS
    cbSize As Long
    dwFlags As Long
    iWaitMSec As Long
    iDelayMSec As Long
    iRepeatMSec As Long
    iBounceMSec As Long
End Type

Public Type ICONMETRICS
    cbSize As Long
    iHorzSpacing As Long
    iVertSpacing As Long
    iTitleWrap As Long
    lfFont As LOGFONT
End Type

Public Type MONITORINFOEX
    cbSize As Long
    rcMonitor As RECT
    rcWork As RECT
    dwFlags As Long
    szDevice As String * CCHDEVICENAME
End Type

Public Type MOUSEKEYS
    cbSize As Long
    dwFlags As Long
    iMaxSpeed As Long
    iTimeToMaxSpeed As Long
    iCtrlSpeed As Long
    dwReserved1 As Long
    dwReserved2 As Long
End Type

Public Type SERIALKEYS
    cbSize As Long
    dwFlags As Long
    lpszActivePort As String
    lpszPort As String
    iBaudRate As Long
    iPortState As Long
    iActive As Long
End Type

Public Type SOUNDSENTRY
    cbSize As Long
    dwFlags As Long
    iFSTextEffect As Long
    iFSTextEffectMSec As Long
    iFSTextEffectColorBits As Long
    iFSGrafEffect As Long
    iFSGrafEffectMSec As Long
    iFSGrafEffectColor As Long
    iWindowsEffect As Long
    iWindowsEffectMSec As Long
    lpszWindowsEffectDLL As String
    iWindowsEffectOrdinal As Long
End Type

Public Type STICKYKEYS
    cbSize As Long
    dwFlags As Long
End Type

Public Type TOGGLEKEYS
    cbSize As Long
    dwFlags As Long
End Type

Public Type WINDOWINFO
    cbSize As Long
    rcWindow As RECT
    rcClient As RECT
    dwStyle As Long
    dwExStyle As Long
    dwWindowStatus As Long
    cxWindowBorders As Long
    cyWindowBorders As Long
    atomWindowType As Long
    wCreatorVersion As Integer
End Type

Public Type WINDOWPLACEMENT
    Length As Long
    flags As Long
    showCmd As Long
    ptMinPosition As POINTAPI
    ptMaxPosition As POINTAPI
    rcNormalPosition As RECT
End Type


Public Function EnumWindowsProc(ByVal hwnd As Long, ByVal lParam As Long) As Long
    WindowListNum = WindowListNum + 1
    ReDim Preserve WindowList(WindowListNum)
    WindowList(WindowListNum) = hwnd
    
    EnumWindowsProc = 1
End Function

Public Function Get_WindowText(ByVal hwnd As Long) As String
    Dim strWindowTitle As String
    Dim lngRetValue As Long
    
    strWindowTitle = String$(GetWindowTextLength(hwnd) + 1, &H0)
    lngRetValue = GetWindowText(hwnd, strWindowTitle, Len(strWindowTitle))
    
    If lngRetValue = 0 Then
        If Err.LastDllError <> 0 Then
            Failed "GetWindowText"
        End If
    Else
        Get_WindowText = Fix_NullTermStr(Left$(strWindowTitle, lngRetValue))
    End If
End Function

Public Function MonitorEnumProc(ByVal hMonitor As Long, ByVal hdcMonitor As Long, ByRef lprcMonitor As RECT, ByVal dwData As Long) As Boolean
    frmDisplayMonitors.cboDisplayMonitors.AddItem CStr(hMonitor)
    
    MonitorEnumProc = 1
End Function
