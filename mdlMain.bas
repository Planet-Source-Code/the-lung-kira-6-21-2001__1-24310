Attribute VB_Name = "mdlMain"
Option Explicit

Public Const WM_PROJECT_TRAY = WM_USER - 1
Public Const WM_PROJECT_WS = WM_USER - 2

Public MouseHook As Long
'Public ShellHook As Long

Public MouseMonitor As MouseMonitor
Public Type MouseMonitor
    TotalXMovement As Double
    TotalYMovement As Double
    TotalWheelMovement As Double
    LastCoordinate As POINTAPI
    TotalLClicks As Double
    TotalMClicks As Double
    TotalRClicks As Double
    TotalX1Clicks As Double
    TotalX2Clicks As Double
    TotalWarp As Double
End Type

Public dblCounterFrequency As Double
Public MouseHook_OldProc As Long
'Public ShellHook_OldProc As Long
Public Tray_OldProc As Long
Public WinID As Long
Public WinVer As Long

Public errMsg As Boolean
Public apiError As Long
Public WS2 As Boolean

Public LocaleList() As String
Public LocaleListNum As Long
Public WindowList() As Long
Public WindowListNum As Long


Public Sub App_Startup()
    If App.PrevInstance = True Then End
    errMsg = False
    
    
    Tray_OldProc = SetWindowLong(frmMain.txtTray.hwnd, GWL_WNDPROC, AddressOf Tray_Proc): If Tray_OldProc = 0 Then Failed "SetWindowLong"
    
    
    Dim NOTIFYICONDATA As NOTIFYICONDATA
    With NOTIFYICONDATA
        .cbSize = Len(NOTIFYICONDATA)
        .hwnd = frmMain.txtTray.hwnd
        .uID = 0
        .uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
        .uCallbackMessage = WM_PROJECT_TRAY
        .hIcon = frmMain.Icon
        .szTip = "Kira" & Chr$(0)
    End With
    
    If Shell_NotifyIcon(NIM_ADD, NOTIFYICONDATA) = False Then
        If NOTIFYICONDATA.uFlags And NIM_SETVERSION Then
            Failed "Shell_NotifyIcon"
        End If
    End If
    
    
    Dim OSVERSIONINFO As OSVERSIONINFO
    OSVERSIONINFO.dwOSVersionInfoSize = Len(OSVERSIONINFO)
    If GetVersionEx(OSVERSIONINFO) = 0 Then Failed "GetVersionEx"
    
    WinID = OSVERSIONINFO.dwPlatformId
    If WinID = VER_PLATFORM_WIN32_WINDOWS Then
        WinVer = CLng(Right$("0" & OSVERSIONINFO.dwMajorVersion, 1) & _
                 Right$("00" & OSVERSIONINFO.dwMinorVersion, 2) & _
                 Right$("0000" & (LO_WORD(OSVERSIONINFO.dwBuildNumber)), 4))
    Else
        WinVer = CLng(Right$("0" & OSVERSIONINFO.dwMajorVersion, 1) & _
                 Right$("00" & OSVERSIONINFO.dwMinorVersion, 2) & _
                 Right$("0000" & (OSVERSIONINFO.dwBuildNumber), 4))
    End If
    
    
    If GetRegSetting(HKEY_CURRENT_USER, "Software\Kira", "Kira") <> "6-21-2001" Then
        DefaultSettings
        frmExtra.Show
    End If
    frmMain.mnuMouseMonitorOO.Checked = CBool(GetRegSetting(HKEY_CURRENT_USER, "Software\Kira", "MouseMonitorOO"))
    frmMain.mnuMouseWarpOO.Checked = CBool(GetRegSetting(HKEY_CURRENT_USER, "Software\Kira", "MouseWarpOO"))
    With MouseMonitor
        .TotalXMovement = CDbl(GetRegSetting(HKEY_CURRENT_USER, "Software\Kira\MouseMonitor", "TotalXMovement"))
        .TotalYMovement = CDbl(GetRegSetting(HKEY_CURRENT_USER, "Software\Kira\MouseMonitor", "TotalYMovement"))
        .TotalWheelMovement = CDbl(GetRegSetting(HKEY_CURRENT_USER, "Software\Kira\MouseMonitor", "TotalWheelMovement"))
        .TotalLClicks = CDbl(GetRegSetting(HKEY_CURRENT_USER, "Software\Kira\MouseMonitor", "TotalLClicks"))
        .TotalMClicks = CDbl(GetRegSetting(HKEY_CURRENT_USER, "Software\Kira\MouseMonitor", "TotalMClicks"))
        .TotalRClicks = CDbl(GetRegSetting(HKEY_CURRENT_USER, "Software\Kira\MouseMonitor", "TotalRClicks"))
        .TotalX1Clicks = CDbl(GetRegSetting(HKEY_CURRENT_USER, "Software\Kira\MouseMonitor", "TotalX1Clicks"))
        .TotalX2Clicks = CDbl(GetRegSetting(HKEY_CURRENT_USER, "Software\Kira\MouseMonitor", "TotalX2Clicks"))
        .TotalWarp = CDbl(GetRegSetting(HKEY_CURRENT_USER, "Software\Kira\MouseWarp", "TotalWarp"))
    End With
    
    
    If frmMain.mnuMouseMonitorOO.Checked = True Then MouseHookInstall
    If frmMain.mnuMouseWarpOO.Checked = True Then MouseHookInstall
    
    
    Winsock_Startup
    
    
    Dim LARGE_INTEGER  As LARGE_INTEGER
    If QueryPerformanceFrequency(LARGE_INTEGER) = False Then Failed "QueryPerformanceFrequency"
    
    dblCounterFrequency = CLargeInt(LARGE_INTEGER.LowPart, LARGE_INTEGER.HighPart)
    If dblCounterFrequency = 0 Then dblCounterFrequency = 1
    
    
    Dim POINTAPI As POINTAPI
    If GetCursorPos(POINTAPI) = 0 Then Failed "GetCursorPos"

    MouseMonitor.LastCoordinate.X = POINTAPI.X
    MouseMonitor.LastCoordinate.Y = POINTAPI.Y
End Sub

Public Sub App_Shutdown()
    CloseAll
    
    
    Winsock_Shutdown
    
    
    If MouseHook > 0 Then
        MouseHook = 1
        MouseHookRemove
    End If
    'If ShellHook > 0 Then
    '    ShellHook = 1
    '    ShellHookRemove
    'End If
    
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira", "MouseMonitorOO", CLng(frmMain.mnuMouseMonitorOO.Checked), REG_DWORD
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira", "MouseWarpOO", CLng(frmMain.mnuMouseWarpOO.Checked), REG_DWORD
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\MouseMonitor", "TotalXMovement", CStr(MouseMonitor.TotalXMovement), REG_SZ
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\MouseMonitor", "TotalYMovement", CStr(MouseMonitor.TotalYMovement), REG_SZ
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\MouseMonitor", "TotalWheelMovement", CStr(MouseMonitor.TotalWheelMovement), REG_SZ
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\MouseMonitor", "TotalLClicks", CStr(MouseMonitor.TotalLClicks), REG_SZ
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\MouseMonitor", "TotalMClicks", CStr(MouseMonitor.TotalMClicks), REG_SZ
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\MouseMonitor", "TotalRClicks", CStr(MouseMonitor.TotalRClicks), REG_SZ
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\MouseMonitor", "TotalX1Clicks", CStr(MouseMonitor.TotalX1Clicks), REG_SZ
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\MouseMonitor", "TotalX2Clicks", CStr(MouseMonitor.TotalX2Clicks), REG_SZ
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\MouseWarp", "TotalWarp", CStr(MouseMonitor.TotalWarp), REG_SZ
    

    Dim NOTIFYICONDATA As NOTIFYICONDATA
    With NOTIFYICONDATA
        .cbSize = Len(NOTIFYICONDATA)
        .hwnd = frmMain.txtTray.hwnd
        .uID = 0
        .uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
        .uCallbackMessage = WM_PROJECT_TRAY
        .hIcon = frmMain.Icon
        .szTip = Chr$(0)
    End With
    
    If Shell_NotifyIcon(NIM_DELETE, NOTIFYICONDATA) = False Then
        If NOTIFYICONDATA.uFlags And NIM_SETVERSION Then
            Failed "Shell_NotifyIcon"
        End If
    End If
    
    
    If SetWindowLong(frmMain.txtTray.hwnd, GWL_WNDPROC, Tray_OldProc) = 0 Then Failed "SetWindowLong"
    
    Unload frmMain
End Sub


Public Function Tray_Proc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Select Case uMsg
        Case WM_PROJECT_TRAY
            Select Case lParam
                Case WM_LBUTTONUP
                    If SetForegroundWindow(frmMain.hwnd) = False Then Failed "SetForegroundWindow"
                    frmMain.PopupMenu frmMain.mnuMain
                Case WM_MBUTTONUP
                    If SetForegroundWindow(frmMain.hwnd) = False Then Failed "SetForegroundWindow"
                    frmMain.PopupMenu frmMain.mnuMain
                Case WM_RBUTTONUP
                    If SetForegroundWindow(frmMain.hwnd) = False Then Failed "SetForegroundWindow"
                    frmMain.PopupMenu frmMain.mnuMain
            End Select
            
        Case Else
            Tray_Proc = DefWindowProc(frmMain.txtTray.hwnd, uMsg, wParam, lParam)
    End Select
End Function


Public Sub CloseAll()
    Unload frmAccessTimeout
    Unload frmAdaptersInfo
    Unload frmCachedPasswords
    Unload frmCPUID_00000000
    Unload frmCPUID_00000001
    Unload frmCPUID_00000002
    Unload frmCPUID_80000000
    Unload frmCPUID_80000001
    Unload frmCPUID_80000002_4
    Unload frmCPUID_80000005
    Unload frmCPUID_80000006
    Unload frmCPUID_Other
    Unload frmCPUInfo
    Unload frmDayTime
    Unload frmDirectories
    Unload frmDisplayDevices
    Unload frmDisplayMonitors
    Unload frmDisplaySettings
    Unload frmDriveInfo
    Unload frmDriveSpace
    Unload frmEcho
    Unload frmErrors
    Unload frmExitWindows
    Unload frmExtra
    Unload frmFileAttributes
    Unload frmFileChecksum
    Unload frmFileTime
    Unload frmFilterKeys
    Unload frmGetIPHost
    Unload frmGIF
    Unload frmHeaps
    Unload frmICMP_Echo
    Unload frmICMPStatistics
    Unload frmIconMetrics
    Unload frmIconSettings
    Unload frmIEHistory
    Unload frmIESettings
    Unload frmIPAddressTable
    Unload frmIPForwardTable
    Unload frmIPNetTable
    Unload frmIPStatistics
    Unload frmKeyboardInfo
    Unload frmLocalesCurrency
    Unload frmLocalesDate
    Unload frmLocalesGeneral
    Unload frmLocalesNumber
    Unload frmLocalesTime
    Unload frmMemoryInfo
    Unload frmMemoryStatus
    Unload frmMenuSettings
    Unload frmMIB2IFTable
    Unload frmModules
    Unload frmMouseInfo
    Unload frmMouseKeys
    Unload frmMouseMonitor
    Unload frmMouseSettings
    Unload frmMouseWarp
    Unload frmMZ
    Unload frmNetworkInfo
    Unload frmNicname_Whois
    Unload frmName_Finger
    Unload frmOperatingTime
    Unload frmPowerStatus
    Unload frmProcesses
    Unload frmProcessTimes
    Unload frmQOTD
    Unload frmRecycleBin
    Unload frmSerialKeys
    Unload frmSharedFiles
    Unload frmSoundSentry
    Unload frmStartMenu
    Unload frmStickyKeys
    Unload frmTCPStatistics
    Unload frmTCPTable
    Unload frmThreads
    Unload frmThreadTimes
    Unload frmTime
    Unload frmToggleKeys
    Unload frmUDPSender
    Unload frmUDPStatistics
    Unload frmUDPTable
    Unload frmUser
    Unload frmWallpaper
    Unload frmWFPProtectedFiles
    Unload frmWFPSettings
    Unload frmWindowMetrics
    Unload frmWindowInfo
    Unload frmWindowPlacement
    Unload frmWindowSettings
    Unload frmWindowsInfo
    Unload frmWinsockInfo
End Sub


Public Sub DefaultSettings()
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira", "Kira", "6-21-2001", REG_SZ
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira", "MouseMonitorOO", 0, REG_DWORD
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira", "MouseWarpOO", 0, REG_DWORD
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\CPUID_Other", "Level", "0", REG_SZ
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\DayTime", "HostIP", "127.0.0.1", REG_SZ
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\DayTime", "Method", 0, REG_DWORD
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\DayTime", "Port", 13, REG_DWORD
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\DisplaySettings", "GlobalChange", 1, REG_DWORD
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\DisplaySettings", "Test", 1, REG_DWORD
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\DriveSpace", "Output", 3, REG_DWORD
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\DriveSpace", "Round", 0, REG_DWORD
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\Echo", "DataSize", 0, REG_DWORD
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\Echo", "HostIP", "127.0.0.1", REG_SZ
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\Echo", "Method", 0, REG_DWORD
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\Echo", "Port", 7, REG_DWORD
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\Errors", "Number", 0, REG_DWORD
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\Errors", "Type", 0, REG_DWORD
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\ExitWindows", "Force", 0, REG_DWORD
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\ExitWindows", "ForceIfHung", 0, REG_DWORD
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\ExitWindows", "Method", 0, REG_DWORD
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\GetIPHost", "Host", "localhost", REG_SZ
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\GetIPHost", "IP", "127.0.0.1", REG_SZ
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\ICMP_Echo", "HostIP", "127.0.0.1", REG_SZ
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\ICMP_Echo", "Number", 1, REG_DWORD
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\ICMP_Echo", "Timeout", 5000, REG_DWORD
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\IPAddressTable", "Sorted", 1, REG_DWORD
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\IPForwardTable", "Sorted", 1, REG_DWORD
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\IPNetTable", "Sorted", 1, REG_DWORD
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\MemoryStatus", "Output", 2, REG_DWORD
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\MemoryStatus", "Round", 0, REG_DWORD
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\MIB2IFTable", "Sorted", 1, REG_DWORD
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\MouseMonitor", "TotalXMovement", "0", REG_SZ
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\MouseMonitor", "TotalYMovement", "0", REG_SZ
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\MouseMonitor", "TotalWheelMovement", "0", REG_SZ
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\MouseMonitor", "TotalLClicks", "0", REG_SZ
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\MouseMonitor", "TotalMClicks", "0", REG_SZ
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\MouseMonitor", "TotalRClicks", "0", REG_SZ
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\MouseMonitor", "TotalX1Clicks", "0", REG_SZ
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\MouseMonitor", "TotalX2Clicks", "0", REG_SZ
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\MouseWarp", "TotalWarp", "0", REG_SZ
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\Name_Finger", "HostIP", "127.0.0.1", REG_SZ
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\Name_Finger", "Port", 79, REG_DWORD
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\Name_Finger", "Send", "", REG_SZ
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\Nicname_Whois", "HostIP", "127.0.0.1", REG_SZ
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\Nicname_Whois", "Port", 43, REG_DWORD
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\Nicname_Whois", "Send", "", REG_SZ
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\QOTD", "HostIP", "127.0.0.1", REG_SZ
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\QOTD", "Method", 0, REG_DWORD
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\QOTD", "Port", 17, REG_DWORD
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\RecycleBin", "Confirmation", 1, REG_DWORD
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\RecycleBin", "ProgressUI", 1, REG_DWORD
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\RecycleBin", "Sound", 1, REG_DWORD
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\TCPTable", "Sorted", 1, REG_DWORD
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\Time", "DaylightSavings", 0, REG_DWORD
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\Time", "HostIP", "127.0.0.1", REG_SZ
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\Time", "Method", 0, REG_DWORD
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\Time", "Port", 37, REG_DWORD
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\UDPSender", "DataSize", 0, REG_DWORD
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\UDPSender", "HostIP", "127.0.0.1", REG_SZ
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\UDPSender", "Number", 0, REG_DWORD
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\UDPSender", "Port", 0, REG_DWORD
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\UDPTable", "Sorted", 1, REG_DWORD
End Sub
