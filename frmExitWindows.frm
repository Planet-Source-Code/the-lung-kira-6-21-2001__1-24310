VERSION 5.00
Begin VB.Form frmExitWindows 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Exit Windows"
   ClientHeight    =   1650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4095
   Icon            =   "frmExitWindows.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1650
   ScaleWidth      =   4095
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkForceIfHung 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   3720
      TabIndex        =   5
      Top             =   840
      Width           =   255
   End
   Begin VB.CheckBox chkForce 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   3720
      TabIndex        =   3
      Top             =   600
      Width           =   255
   End
   Begin VB.ComboBox cboMethod 
      Height          =   315
      Left            =   1560
      TabIndex        =   1
      Top             =   120
      Width           =   2415
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   350
      Left            =   3000
      TabIndex        =   6
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label lblForceIfHung 
      Caption         =   "Force If Hung"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label lblForce 
      Caption         =   "Force"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label lblMethod 
      Caption         =   "Method"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmExitWindows"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOk_Click()
    Dim flags As Long
    
    Select Case cboMethod.ListIndex
        Case 0: flags = EWX_LOGOFF
        Case 1: flags = EWX_POWEROFF
        Case 2: flags = EWX_REBOOT
        Case 3: flags = EWX_SHUTDOWN
    End Select
    
    If chkForce.value = 1 Then flags = flags Or EWX_FORCE
    If chkForceIfHung.value = 1 Then flags = flags Or EWX_FORCEIFHUNG
    
    
    If WinVersion(0, -1, True) = True Then
        If ExitWindowsEx(flags, &H0) = False Then Failed "ExitWindowsEx"
    Else
        If cboMethod.ListIndex = 0 Then
            If ExitWindowsEx(flags, &H0) = False Then Failed "ExitWindowsEx"
        End If
        
        
        Dim hTokenHandle As Long
        Dim tmpLuid As LUID
        Dim tkpNewState As TOKEN_PRIVILEGES
        Dim tkpPreviousState As TOKEN_PRIVILEGES
        Dim lngBufferLen As Long
        
        
        If OpenProcessToken(GetCurrentProcess(), (TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY), hTokenHandle) = 0 Then Failed "OpenProcessToken"
        If LookupPrivilegeValue(&H0, SE_SHUTDOWN_NAME, tmpLuid) = 0 Then Failed "LookupPrivilegeValue"
        
        With tkpNewState
            .PrivilegeCount = 1
            .Privileges(0).Attributes = SE_PRIVILEGE_ENABLED
            .Privileges(0).pLuid = tmpLuid
        End With
        
        If AdjustTokenPrivileges(hTokenHandle, False, tkpNewState, Len(tkpPreviousState), tkpPreviousState, lngBufferLen) = False Then Failed "LookupPrivilegeValue"
        If CloseHandle(hTokenHandle) = False Then Failed "CloseHandle"
        If ExitWindowsEx(flags, &H0) = False Then Failed "ExitWindowsEx"
    End If
End Sub

Private Sub Form_Load()
    With cboMethod
        .AddItem "Logoff"
        .AddItem "Poweroff"
        .AddItem "Reboot"
        .AddItem "Shutdown"
    End With
    
    If WinVersion(-1, 5000000, True) = False Then
        lblForceIfHung.Enabled = False
        chkForceIfHung.Enabled = False
    End If
    
    chkForce.value = GetRegSetting(HKEY_CURRENT_USER, "Software\Kira\ExitWindows", "Force")
    chkForceIfHung.value = GetRegSetting(HKEY_CURRENT_USER, "Software\Kira\ExitWindows", "ForceIfHung")
    cboMethod.ListIndex = GetRegSetting(HKEY_CURRENT_USER, "Software\Kira\ExitWindows", "Method")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\ExitWindows", "Force", chkForce.value, REG_DWORD
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\ExitWindows", "ForceIfHung", chkForceIfHung.value, REG_DWORD
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\ExitWindows", "Method", cboMethod.ListIndex, REG_DWORD
End Sub
