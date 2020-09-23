VERSION 5.00
Begin VB.Form frmDisplaySettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Display Settings"
   ClientHeight    =   2130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4935
   Icon            =   "frmDisplaySettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2130
   ScaleWidth      =   4935
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkAllModes 
      Height          =   255
      Left            =   4560
      TabIndex        =   7
      Top             =   840
      Width           =   255
   End
   Begin VB.CheckBox chkTest 
      Height          =   255
      Left            =   4560
      TabIndex        =   11
      Top             =   1320
      Width           =   255
   End
   Begin VB.ComboBox cboModes 
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1560
      TabIndex        =   5
      Top             =   360
      Width           =   3255
   End
   Begin VB.CheckBox chkGlobal 
      Height          =   255
      Left            =   4560
      TabIndex        =   9
      Top             =   1080
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   350
      Left            =   3840
      TabIndex        =   13
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   350
      Left            =   2760
      TabIndex        =   12
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label lblAllModes 
      Caption         =   "All Modes"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label lblTest 
      Caption         =   "Test"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label lblRefresh 
      Caption         =   "Refresh"
      Height          =   255
      Left            =   3750
      TabIndex        =   3
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lblBPP 
      Caption         =   "BPP"
      Height          =   255
      Left            =   3030
      TabIndex        =   2
      Top             =   120
      Width           =   375
   End
   Begin VB.Label lblHeight 
      Caption         =   "Height"
      Height          =   255
      Left            =   2310
      TabIndex        =   1
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lblWidth 
      Caption         =   "Width"
      Height          =   255
      Left            =   1560
      TabIndex        =   0
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lblGlobal 
      Caption         =   "Global Change"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label lblAvailable 
      Caption         =   "Modes Available"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "frmDisplaySettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdApply_Click()
    Dim DEVMODE As DEVMODE
    
    With DEVMODE
        .dmSize = Len(DEVMODE)
        .dmBitsPerPel = CLng(Trim$(Mid$(cboModes.List(cboModes.ListIndex), 12, 6)))
        .dmPelsWidth = CLng(Trim$(Mid$(cboModes.List(cboModes.ListIndex), 1, 6)))
        .dmPelsHeight = CLng(Trim$(Mid$(cboModes.List(cboModes.ListIndex), 6, 6)))
        .dmFields = DM_BITSPERPEL Or DM_PELSWIDTH Or DM_PELSHEIGHT Or DM_DISPLAYFREQUENCY
        .dmDisplayFrequency = CLng(Trim$(Mid$(cboModes.List(cboModes.ListIndex), 18, 6)))
        '.dmPosition 'Multimonitor
    End With
    
    
    If chkTest.value = 1 Then
        If ChangeDisplaySettings(DEVMODE, CDS_TEST) <> 0 Then
            If MessageBoxEx(&H0, "Display test failed. Mode was not set.", "Error", MB_OK Or MB_ICONWARNING Or MB_SETFOREGROUND, 0) = 0 Then Failed "MessageBoxEx"
            Exit Sub
        End If
    End If
    If chkGlobal.value = 1 Then
        apiError = ChangeDisplaySettings(DEVMODE, CDS_UPDATEREGISTRY Or CDS_GLOBAL)
    Else
        apiError = ChangeDisplaySettings(DEVMODE, CDS_UPDATEREGISTRY)
    End If
    
    Select Case apiError
        Case DISP_CHANGE_RESTART: If MessageBoxEx(&H0, "Must restart Windows for changes to be implemented.", "Restart", MB_OK Or MB_ICONWARNING Or MB_SETFOREGROUND, 0) = 0 Then Failed "MessageBoxEx"
        Case DISP_CHANGE_BADFLAGS: Failed "ChangeDisplaySettings"
        Case DISP_CHANGE_BADPARAM: Failed "ChangeDisplaySettings"
        Case DISP_CHANGE_FAILED: Failed "ChangeDisplaySettings"
        Case DISP_CHANGE_BADMODE: Failed "ChangeDisplaySettings"
        Case DISP_CHANGE_NOTUPDATED: Failed "ChangeDisplaySettings"
        Case DISP_CHANGE_BADDUALVIEW: Failed "ChangeDisplaySettings"
    End Select
End Sub

Private Sub cmdRefresh_Click()
    Dim DEVMODE As DEVMODE
    Dim lngIncrement As Long
    
    Dim curBPP As Integer
    Dim curWidth As Integer
    Dim curHeight As Integer
    Dim curVRefresh As Integer
    
    DEVMODE.dmSize = Len(DEVMODE)
    
    curBPP = GetDeviceCaps(frmMain.hdc, BITSPIXEL)
    curWidth = Screen.Width \ Screen.TwipsPerPixelX
    curHeight = Screen.Height \ Screen.TwipsPerPixelY
    If WinID = VER_PLATFORM_WIN32_NT Then
        curVRefresh = GetDeviceCaps(frmDisplaySettings.hdc, VREFRESH)
    Else
        curVRefresh = GetRegSetting(HKEY_CURRENT_CONFIG, "Display\Settings", "RefreshRate")
    End If
    
    
    Dim Ex As Boolean
    Dim lngFlags As Long
    
    Ex = WinVersion(4010000, 5000000, True)
    If chkAllModes.value = 1 Then lngFlags = EDS_RAWMODE
    
    Do
        If Ex = True Then
            If EnumDisplaySettingsEx(ByVal 0, lngIncrement, DEVMODE, lngFlags) = 0 Then
                Failed "EnumDisplaySettings"
                Exit Do
            End If
        Else
            If EnumDisplaySettings(ByVal 0, lngIncrement, DEVMODE) = 0 Then
                Failed "EnumDisplaySettings"
                Exit Do
            End If
        End If
        
        lngIncrement = lngIncrement + 1
        
        With DEVMODE
            cboModes.AddItem Left$(.dmPelsWidth & Space$(6), 6) & _
                             Left$(.dmPelsHeight & Space$(6), 6) & _
                             Left$(.dmBitsPerPel & Space$(6), 6) & _
                             Left$(.dmDisplayFrequency & Space$(6), 6)
            
            If curBPP = .dmBitsPerPel And _
            curWidth = .dmPelsWidth And _
            curHeight = .dmPelsHeight And _
            curVRefresh = .dmDisplayFrequency Then
                cboModes.ListIndex = cboModes.NewIndex
            End If
        End With
    Loop
End Sub

Private Sub Form_Load()
    If WinVersion(4010000, 5000000, True) = False Then chkAllModes.Enabled = False
    cmdRefresh_Click
    
    chkGlobal.value = GetRegSetting(HKEY_CURRENT_USER, "Software\Kira\DisplaySettings", "GlobalChange")
    chkTest.value = GetRegSetting(HKEY_CURRENT_USER, "Software\Kira\DisplaySettings", "Test")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\DisplaySettings", "GlobalChange", chkGlobal.value, REG_DWORD
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\DisplaySettings", "Test", chkTest.value, REG_DWORD
End Sub
