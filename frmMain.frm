VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   255
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   1590
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   255
   ScaleWidth      =   1590
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.TextBox txtShellHook 
      Height          =   195
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtMouseHook 
      Height          =   195
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtTray 
      Height          =   195
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Menu mnuMain 
      Caption         =   ""
      Begin VB.Menu mnuHardware 
         Caption         =   "Hardware"
         Begin VB.Menu mnuDrive 
            Caption         =   "Drive"
            Begin VB.Menu mnuCDPlayer 
               Caption         =   "CD Player"
            End
            Begin VB.Menu mnuDriveInfo 
               Caption         =   "Drive Info"
            End
            Begin VB.Menu mnuDriveSpace 
               Caption         =   "Drive Space"
            End
            Begin VB.Menu mnuFiles 
               Caption         =   "Files"
               Begin VB.Menu mnuFileAttributes 
                  Caption         =   "File Attributes"
               End
               Begin VB.Menu mnuFileChecksum 
                  Caption         =   "File Checksum"
               End
               Begin VB.Menu mnuFileTime 
                  Caption         =   "File Time"
               End
               Begin VB.Menu mnuFormats 
                  Caption         =   "Formats"
                  Begin VB.Menu mnuGIF 
                     Caption         =   "GIF"
                  End
                  Begin VB.Menu mnuMZ 
                     Caption         =   "MZ"
                  End
               End
            End
         End
         Begin VB.Menu mnuCPU 
            Caption         =   "CPU"
            Begin VB.Menu mnuCPUID 
               Caption         =   "CPUID"
               Begin VB.Menu mnuCPUID_00000000 
                  Caption         =   "00000000"
               End
               Begin VB.Menu mnuCPUID_00000001 
                  Caption         =   "00000001"
               End
               Begin VB.Menu mnuCPUID_00000002 
                  Caption         =   "00000002"
               End
               Begin VB.Menu mnuCPUID_80000000 
                  Caption         =   "80000000"
               End
               Begin VB.Menu mnuCPUID_80000001 
                  Caption         =   "80000001"
               End
               Begin VB.Menu mnuCPUID_80000002_4 
                  Caption         =   "80000002-4"
               End
               Begin VB.Menu mnuCPUID_80000005 
                  Caption         =   "80000005"
               End
               Begin VB.Menu mnuCPUID_80000006 
                  Caption         =   "80000006"
               End
               Begin VB.Menu mnuCPUID_Other 
                  Caption         =   "Other"
               End
            End
            Begin VB.Menu mnuCPUInfo 
               Caption         =   "CPU Info"
            End
         End
         Begin VB.Menu mnuMemory 
            Caption         =   "Memory"
            Begin VB.Menu mnuMemoryInfo 
               Caption         =   "Memory Info"
            End
            Begin VB.Menu mnuMemoryStatus 
               Caption         =   "Memory Status"
            End
         End
         Begin VB.Menu mnuPowerStatus 
            Caption         =   "Power Status"
         End
      End
      Begin VB.Menu mnuMonitor 
         Caption         =   "Monitor"
         Begin VB.Menu mnuMouseMonitor 
            Caption         =   "Mouse Monitor"
         End
         Begin VB.Menu mnuMouseWarp 
            Caption         =   "Mouse Warp"
         End
         Begin VB.Menu mnuOnOff 
            Caption         =   "On/Off"
            Begin VB.Menu mnuMouseMonitorOO 
               Caption         =   "Mouse Monitor"
            End
            Begin VB.Menu mnuMouseWarpOO 
               Caption         =   "Mouse Warp"
            End
         End
      End
      Begin VB.Menu mnuNetwork 
         Caption         =   "Network"
         Begin VB.Menu mnuGetIPHost 
            Caption         =   "Get IP / Host"
         End
         Begin VB.Menu mnuICMP 
            Caption         =   "ICMP"
            Begin VB.Menu mnuICMP_Echo 
               Caption         =   "Echo"
            End
         End
         Begin VB.Menu mnuNetwork_Info 
            Caption         =   "Info"
            Begin VB.Menu mnuAdaptersInfo 
               Caption         =   "Adapters Info"
            End
            Begin VB.Menu mnuICMPStatistics 
               Caption         =   "ICMP Statistics"
            End
            Begin VB.Menu mnuIPAddressTable 
               Caption         =   "IP Address Table"
            End
            Begin VB.Menu mnuIPForwardTable 
               Caption         =   "IP Forward Table"
            End
            Begin VB.Menu mnuIPNetTable 
               Caption         =   "IP Net Table"
            End
            Begin VB.Menu mnuIPStatistics 
               Caption         =   "IP Statistics"
            End
            Begin VB.Menu mnuMIB2IFTable 
               Caption         =   "MIB-II Interface Table"
            End
            Begin VB.Menu mnuNetworkInfo 
               Caption         =   "Network Info"
            End
            Begin VB.Menu mnuTCPStatistics 
               Caption         =   "TCP Statistics"
            End
            Begin VB.Menu mnuTCPTable 
               Caption         =   "TCP Table"
            End
            Begin VB.Menu mnuUDPStatistics 
               Caption         =   "UDP Statistics"
            End
            Begin VB.Menu mnuUDPTable 
               Caption         =   "UDP Table"
            End
            Begin VB.Menu mnuWinsockInfo 
               Caption         =   "Winsock Info"
            End
         End
         Begin VB.Menu mnuService 
            Caption         =   "Service"
            Begin VB.Menu mnuDayTime 
               Caption         =   "DayTime"
            End
            Begin VB.Menu mnuEcho 
               Caption         =   "Echo"
            End
            Begin VB.Menu mnuName_Finger 
               Caption         =   "Name/Finger"
            End
            Begin VB.Menu mnuNicname_Whois 
               Caption         =   "Nicname/Whois"
            End
            Begin VB.Menu mnuQOTD 
               Caption         =   "Quote Of The Day"
            End
            Begin VB.Menu mnuTime 
               Caption         =   "Time"
            End
         End
         Begin VB.Menu mnuUDPSender 
            Caption         =   "UDP Sender"
         End
      End
      Begin VB.Menu mnuPeripherals 
         Caption         =   "Peripherals"
         Begin VB.Menu mnuDisplay 
            Caption         =   "Display"
            Begin VB.Menu mnuDisplayDevices 
               Caption         =   "Display Devices"
            End
            Begin VB.Menu mnuDisplayMonitors 
               Caption         =   "Display Monitors"
            End
            Begin VB.Menu mnuDisplaySettings 
               Caption         =   "Display Settings"
            End
         End
         Begin VB.Menu mnuKeyboard 
            Caption         =   "Keyboard"
            Begin VB.Menu mnuKeyboardInfo 
               Caption         =   "Keyboard Info"
            End
            Begin VB.Menu mnuKeyboardSettings 
               Caption         =   "Keyboard Settings"
            End
         End
         Begin VB.Menu mnuMouse 
            Caption         =   "Mouse"
            Begin VB.Menu mnuMouseInfo 
               Caption         =   "Mouse Info"
            End
            Begin VB.Menu mnuMouseSettings 
               Caption         =   "Mouse Settings"
            End
         End
      End
      Begin VB.Menu mnuWindowS 
         Caption         =   "Windows"
         Begin VB.Menu mnuShellAbout 
            Caption         =   "About Shell"
         End
         Begin VB.Menu mnuAccessibility 
            Caption         =   "Accessibility"
            Begin VB.Menu mnuAccessTimeout 
               Caption         =   "Access Timeout"
            End
            Begin VB.Menu mnuFilterKeys 
               Caption         =   "Filter Keys"
            End
            Begin VB.Menu mnuMouseKeys 
               Caption         =   "Mouse Keys"
            End
            Begin VB.Menu mnuSerialKeys 
               Caption         =   "Serial Keys"
            End
            Begin VB.Menu mnuSoundSentry 
               Caption         =   "Sound Sentry"
            End
            Begin VB.Menu mnuStickyKeys 
               Caption         =   "Sticky Keys"
            End
            Begin VB.Menu mnuToggleKeys 
               Caption         =   "Toggle Keys"
            End
         End
         Begin VB.Menu mnuCachedPasswords 
            Caption         =   "Cached Passwords"
         End
         Begin VB.Menu mnuDirectory 
            Caption         =   "Directory"
            Begin VB.Menu mnuDirectories 
               Caption         =   "Directories"
            End
            Begin VB.Menu mnuRecycleBin 
               Caption         =   "Recycle Bin"
            End
         End
         Begin VB.Menu mnuErrors 
            Caption         =   "Errors"
         End
         Begin VB.Menu mnuExitWindows 
            Caption         =   "Exit Windows"
         End
         Begin VB.Menu mnuWindowsFiles 
            Caption         =   "Files"
            Begin VB.Menu mnuSharedFiles 
               Caption         =   "Shared Files"
            End
            Begin VB.Menu mnuWindowsFileProtection 
               Caption         =   "Windows File Protection"
               Begin VB.Menu mnuWFPProtectedFiles 
                  Caption         =   "Protected Files"
               End
               Begin VB.Menu mnuWFPSettings 
                  Caption         =   "Settings"
               End
            End
         End
         Begin VB.Menu mnuIconS 
            Caption         =   "Icons"
            Begin VB.Menu mnuIconMetrics 
               Caption         =   "Icon Metrics"
            End
            Begin VB.Menu mnuIconSettings 
               Caption         =   "Icon Settings"
            End
         End
         Begin VB.Menu mnuIE 
            Caption         =   "Internet Explorer"
            Begin VB.Menu mnuIEHistory 
               Caption         =   "IE History"
            End
            Begin VB.Menu mnuIESettings 
               Caption         =   "IE Settings"
            End
         End
         Begin VB.Menu mnuLocales 
            Caption         =   "Locales"
            Begin VB.Menu mnuLocalesCurrency 
               Caption         =   "Currency"
            End
            Begin VB.Menu mnuLocalesDate 
               Caption         =   "Date"
            End
            Begin VB.Menu mnuLocalesGeneral 
               Caption         =   "General"
            End
            Begin VB.Menu mnuLocalesNumber 
               Caption         =   "Number"
            End
            Begin VB.Menu mnuLocalesTime 
               Caption         =   "Time"
            End
         End
         Begin VB.Menu mnuMenu 
            Caption         =   "Menu"
            Begin VB.Menu mnuMenuSettings 
               Caption         =   "Menu Settings"
            End
            Begin VB.Menu mnuStartMenu 
               Caption         =   "Start Menu"
            End
         End
         Begin VB.Menu mnuOperatingTime 
            Caption         =   "Operating Time"
         End
         Begin VB.Menu mnuProcess 
            Caption         =   "Processes"
            Begin VB.Menu mnuHeaps 
               Caption         =   "Heaps"
            End
            Begin VB.Menu mnuModules 
               Caption         =   "Modules"
            End
            Begin VB.Menu mnuProcesses 
               Caption         =   "Processes"
            End
            Begin VB.Menu mnuProcessTimes 
               Caption         =   "Process Times"
            End
            Begin VB.Menu mnuThread 
               Caption         =   "Threads"
               Begin VB.Menu mnuThreads 
                  Caption         =   "Threads"
               End
               Begin VB.Menu mnuThreadTimes 
                  Caption         =   "Thread Times"
               End
            End
         End
         Begin VB.Menu mnuUser 
            Caption         =   "User"
         End
         Begin VB.Menu mnuWallpaper 
            Caption         =   "Wallpaper"
         End
         Begin VB.Menu mnuWindow 
            Caption         =   "Windows"
            Begin VB.Menu mnuWindowInfo 
               Caption         =   "Window Info"
            End
            Begin VB.Menu mnuWindowMetrics 
               Caption         =   "Window Metrics"
            End
            Begin VB.Menu mnuWindowPlacement 
               Caption         =   "Window Placement"
            End
            Begin VB.Menu mnuWindowSettings 
               Caption         =   "Window Settings"
            End
         End
         Begin VB.Menu mnuWindowsInfo 
            Caption         =   "Windows Info"
         End
      End
      Begin VB.Menu mnuBreak2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuApplication 
         Caption         =   "Application"
         Begin VB.Menu mnuCloseAll 
            Caption         =   "Close All"
         End
         Begin VB.Menu mnuDefaultSettings 
            Caption         =   "Default Settings"
         End
      End
      Begin VB.Menu mnuBreak1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExtra 
         Caption         =   "Kira"
      End
      Begin VB.Menu mnuBreak0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    App_Startup
End Sub

Private Sub mnuAccessTimeout_Click()
    frmAccessTimeout.Show
End Sub

Private Sub mnuAdaptersInfo_Click()
    frmAdaptersInfo.Show
End Sub

Private Sub mnuCachedPasswords_Click()
    frmCachedPasswords.Show
End Sub

Private Sub mnuCloseAll_Click()
    CloseAll
End Sub

Private Sub mnuCPUID_00000000_Click()
    frmCPUID_00000000.Show
End Sub

Private Sub mnuCPUID_00000001_Click()
    frmCPUID_00000001.Show
End Sub

Private Sub mnuCPUID_00000002_Click()
    frmCPUID_00000002.Show
End Sub

Private Sub mnuCPUID_80000000_Click()
    frmCPUID_80000000.Show
End Sub

Private Sub mnuCPUID_80000001_Click()
    frmCPUID_80000001.Show
End Sub

Private Sub mnuCPUID_80000002_4_Click()
    frmCPUID_80000002_4.Show
End Sub

Private Sub mnuCPUID_80000005_Click()
    frmCPUID_80000005.Show
End Sub

Private Sub mnuCPUID_80000006_Click()
    frmCPUID_80000006.Show
End Sub

Private Sub mnuCPUID_Other_Click()
    frmCPUID_Other.Show
End Sub

Private Sub mnuCPUInfo_Click()
    frmCPUInfo.Show
End Sub

Private Sub mnuDayTime_Click()
    frmDayTime.Show
End Sub

Private Sub mnuDefaultSettings_Click()
    DefaultSettings
End Sub

Private Sub mnuDirectories_Click()
    frmDirectories.Show
End Sub

Private Sub mnuDisplayDevices_Click()
    frmDisplayDevices.Show
End Sub

Private Sub mnuDisplayMonitors_Click()
    frmDisplayMonitors.Show
End Sub

Private Sub mnuDisplaySettings_Click()
    frmDisplaySettings.Show
End Sub

Private Sub mnuDriveInfo_Click()
    frmDriveInfo.Show
End Sub

Private Sub mnuDriveSpace_Click()
    frmDriveSpace.Show
End Sub

Private Sub mnuEcho_Click()
    frmEcho.Show
End Sub

Private Sub mnuErrors_Click()
    frmErrors.Show
End Sub

Private Sub mnuExit_Click()
    App_Shutdown
End Sub

Private Sub mnuExitWindows_Click()
    frmExitWindows.Show
End Sub

Private Sub mnuExtra_Click()
    frmExtra.Show
End Sub

Private Sub mnuFileAttributes_Click()
    frmFileAttributes.Show
End Sub

Private Sub mnuFileChecksum_Click()
    frmFileChecksum.Show
End Sub

Private Sub mnuFileTime_Click()
    frmFileTime.Show
End Sub

Private Sub mnuFilterKeys_Click()
    frmFilterKeys.Show
End Sub

Private Sub mnuGetIPHost_Click()
    frmGetIPHost.Show
End Sub

Private Sub mnuGIF_Click()
    frmGIF.Show
End Sub

Private Sub mnuHeaps_Click()
    frmHeaps.Show
End Sub

Private Sub mnuICMP_Echo_Click()
    frmICMP_Echo.Show
End Sub

Private Sub mnuICMPStatistics_Click()
    frmICMPStatistics.Show
End Sub

Private Sub mnuIconMetrics_Click()
    frmIconMetrics.Show
End Sub

Private Sub mnuIconSettings_Click()
    frmIconSettings.Show
End Sub

Private Sub mnuIEHistory_Click()
    frmIEHistory.Show
End Sub

Private Sub mnuIESettings_Click()
    frmIESettings.Show
End Sub

Private Sub mnuIPAddressTable_Click()
    frmIPAddressTable.Show
End Sub

Private Sub mnuIPForwardTable_Click()
    frmIPForwardTable.Show
End Sub

Private Sub mnuIPNetTable_Click()
    frmIPNetTable.Show
End Sub

Private Sub mnuIPStatistics_Click()
    frmIPStatistics.Show
End Sub

Private Sub mnuKeyboardInfo_Click()
    frmKeyboardInfo.Show
End Sub

Private Sub mnuKeyboardSettings_Click()
    frmKeyboardSettings.Show
End Sub

Private Sub mnuLocalesCurrency_Click()
    frmLocalesCurrency.Show
End Sub

Private Sub mnuLocalesDate_Click()
    frmLocalesDate.Show
End Sub

Private Sub mnuLocalesGeneral_Click()
    frmLocalesGeneral.Show
End Sub

Private Sub mnuLocalesNumber_Click()
    frmLocalesNumber.Show
End Sub

Private Sub mnuLocalesTime_Click()
    frmLocalesTime.Show
End Sub

Private Sub mnuMemoryInfo_Click()
    frmMemoryInfo.Show
End Sub

Private Sub mnuMemoryStatus_Click()
    frmMemoryStatus.Show
End Sub

Private Sub mnuMenuSettings_Click()
    frmMenuSettings.Show
End Sub

Private Sub mnuMIB2IFTable_Click()
    frmMIB2IFTable.Show
End Sub

Private Sub mnuModules_Click()
    frmModules.Show
End Sub

Private Sub mnuMouseInfo_Click()
    frmMouseInfo.Show
End Sub

Private Sub mnuMouseKeys_Click()
    frmMouseKeys.Show
End Sub

Private Sub mnuMouseMonitor_Click()
    frmMouseMonitor.Show
End Sub

Private Sub mnuMouseMonitorOO_Click()
    If mnuMouseMonitorOO.Checked = False Then
        mnuMouseMonitorOO.Checked = True
        
        Dim POINTAPI As POINTAPI
        If GetCursorPos(POINTAPI) = False Then Failed "GetCursorPos"
        
        MouseMonitor.LastCoordinate.X = POINTAPI.X
        MouseMonitor.LastCoordinate.Y = POINTAPI.Y
        
        MouseHookInstall
    Else
        mnuMouseMonitorOO.Checked = False
        MouseHookRemove
    End If
End Sub

Private Sub mnuMouseSettings_Click()
    frmMouseSettings.Show
End Sub

Private Sub mnuMouseWarp_Click()
    frmMouseWarp.Show
End Sub

Private Sub mnuMouseWarpOO_Click()
    If mnuMouseWarpOO.Checked = False Then
        mnuMouseWarpOO.Checked = True
        
        Dim POINTAPI As POINTAPI
        If GetCursorPos(POINTAPI) = False Then Failed "GetCursorPos"
        
        MouseMonitor.LastCoordinate.X = POINTAPI.X
        MouseMonitor.LastCoordinate.Y = POINTAPI.Y
        
        MouseHookInstall
    Else
        mnuMouseWarpOO.Checked = False
        MouseHookRemove
    End If
End Sub

Private Sub mnuMZ_Click()
    frmMZ.Show
End Sub

Private Sub mnuName_Finger_Click()
    frmName_Finger.Show
End Sub

Private Sub mnuNetworkInfo_Click()
    frmNetworkInfo.Show
End Sub

Private Sub mnuNicname_Whois_Click()
    frmNicname_Whois.Show
End Sub

Private Sub mnuOperatingTime_Click()
    frmOperatingTime.Show
End Sub

Private Sub mnuPowerStatus_Click()
    frmPowerStatus.Show
End Sub

Private Sub mnuProcesses_Click()
    frmProcesses.Show
End Sub

Private Sub mnuProcessTimes_Click()
    frmProcessTimes.Show
End Sub

Private Sub mnuQOTD_Click()
    frmQOTD.Show
End Sub

Private Sub mnuRecycleBin_Click()
    frmRecycleBin.Show
End Sub

Private Sub mnuSerialKeys_Click()
    frmSerialKeys.Show
End Sub

Private Sub mnuSharedFiles_Click()
    frmSharedFiles.Show
End Sub

Private Sub mnuShellAbout_Click()
    If ShellAbout(&H0, "", "", frmMain.Icon) = False Then Failed "ShellAbout"
End Sub

Private Sub mnuSoundSentry_Click()
    frmSoundSentry.Show
End Sub

Private Sub mnuStartMenu_Click()
    frmStartMenu.Show
End Sub

Private Sub mnuStickyKeys_Click()
    frmStickyKeys.Show
End Sub

Private Sub mnuTCPStatistics_Click()
    frmTCPStatistics.Show
End Sub

Private Sub mnuTCPTable_Click()
    frmTCPTable.Show
End Sub

Private Sub mnuThreads_Click()
    frmThreads.Show
End Sub

Private Sub mnuThreadTimes_Click()
    frmThreadTimes.Show
End Sub

Private Sub mnuTime_Click()
    frmTime.Show
End Sub

Private Sub mnuToggleKeys_Click()
    frmToggleKeys.Show
End Sub

Private Sub mnuUDPSender_Click()
    frmUDPSender.Show
End Sub

Private Sub mnuUDPStatistics_Click()
    frmUDPStatistics.Show
End Sub

Private Sub mnuUDPTable_Click()
    frmUDPTable.Show
End Sub

Private Sub mnuUser_Click()
    frmUser.Show
End Sub

Private Sub mnuWallpaper_Click()
    frmWallpaper.Show
End Sub

Private Sub mnuWFPProtectedFiles_Click()
    frmWFPProtectedFiles.Show
End Sub

Private Sub mnuWFPSettings_Click()
    frmWFPSettings.Show
End Sub

Private Sub mnuWindowInfo_Click()
    frmWindowInfo.Show
End Sub

Private Sub mnuWindowMetrics_Click()
    frmWindowMetrics.Show
End Sub

Private Sub mnuWindowPlacement_Click()
    frmWindowPlacement.Show
End Sub

Private Sub mnuWindowSettings_Click()
    frmWindowSettings.Show
End Sub

Private Sub mnuWindowsInfo_Click()
    frmWindowsInfo.Show
End Sub

Private Sub mnuWinsockInfo_Click()
    frmWinsockInfo.Show
End Sub
