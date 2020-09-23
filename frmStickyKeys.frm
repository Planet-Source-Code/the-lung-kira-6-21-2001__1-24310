VERSION 5.00
Begin VB.Form frmStickyKeys 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sticky Keys"
   ClientHeight    =   2655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7575
   Icon            =   "frmStickyKeys.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   7575
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkRWinLocked 
      Enabled         =   0   'False
      Height          =   255
      Left            =   7200
      TabIndex        =   49
      Top             =   1800
      Width           =   255
   End
   Begin VB.CheckBox chkLWinLocked 
      Enabled         =   0   'False
      Height          =   255
      Left            =   7200
      TabIndex        =   41
      Top             =   840
      Width           =   255
   End
   Begin VB.CheckBox chkRWinLatched 
      Enabled         =   0   'False
      Height          =   255
      Left            =   4680
      TabIndex        =   33
      Top             =   1800
      Width           =   255
   End
   Begin VB.CheckBox chkLWinLatched 
      Enabled         =   0   'False
      Height          =   255
      Left            =   4680
      TabIndex        =   25
      Top             =   840
      Width           =   255
   End
   Begin VB.CheckBox chkRShiftLocked 
      Enabled         =   0   'False
      Height          =   255
      Left            =   7200
      TabIndex        =   47
      Top             =   1560
      Width           =   255
   End
   Begin VB.CheckBox chkRCtrlLocked 
      Enabled         =   0   'False
      Height          =   255
      Left            =   7200
      TabIndex        =   45
      Top             =   1320
      Width           =   255
   End
   Begin VB.CheckBox chkRAltLocked 
      Enabled         =   0   'False
      Height          =   255
      Left            =   7200
      TabIndex        =   43
      Top             =   1080
      Width           =   255
   End
   Begin VB.CheckBox chkLShiftLocked 
      Enabled         =   0   'False
      Height          =   255
      Left            =   7200
      TabIndex        =   39
      Top             =   600
      Width           =   255
   End
   Begin VB.CheckBox chkLCtrlLocked 
      Enabled         =   0   'False
      Height          =   255
      Left            =   7200
      TabIndex        =   37
      Top             =   360
      Width           =   255
   End
   Begin VB.CheckBox chkLAltLocked 
      Enabled         =   0   'False
      Height          =   255
      Left            =   7200
      TabIndex        =   35
      Top             =   120
      Width           =   255
   End
   Begin VB.CheckBox chkRShiftLatched 
      Enabled         =   0   'False
      Height          =   255
      Left            =   4680
      TabIndex        =   31
      Top             =   1560
      Width           =   255
   End
   Begin VB.CheckBox chkRCtrlLatched 
      Enabled         =   0   'False
      Height          =   255
      Left            =   4680
      TabIndex        =   29
      Top             =   1320
      Width           =   255
   End
   Begin VB.CheckBox chkRAltLatched 
      Enabled         =   0   'False
      Height          =   255
      Left            =   4680
      TabIndex        =   27
      Top             =   1080
      Width           =   255
   End
   Begin VB.CheckBox chkLShiftLatched 
      Enabled         =   0   'False
      Height          =   255
      Left            =   4680
      TabIndex        =   23
      Top             =   600
      Width           =   255
   End
   Begin VB.CheckBox chkLCtrlLatched 
      Enabled         =   0   'False
      Height          =   255
      Left            =   4680
      TabIndex        =   21
      Top             =   360
      Width           =   255
   End
   Begin VB.CheckBox chkLAltLatched 
      Enabled         =   0   'False
      Height          =   255
      Left            =   4680
      TabIndex        =   19
      Top             =   120
      Width           =   255
   End
   Begin VB.CheckBox chkTwoKeysOff 
      Height          =   255
      Left            =   2160
      TabIndex        =   13
      Top             =   1680
      Width           =   255
   End
   Begin VB.CheckBox chkTristate 
      Height          =   255
      Left            =   2160
      TabIndex        =   11
      Top             =   1440
      Width           =   255
   End
   Begin VB.CheckBox chkStickyKeysOn 
      Height          =   255
      Left            =   2160
      TabIndex        =   9
      Top             =   1200
      Width           =   255
   End
   Begin VB.CheckBox chkIndicator 
      Height          =   255
      Left            =   2160
      TabIndex        =   17
      Top             =   2280
      Width           =   255
   End
   Begin VB.CheckBox chkHotKeySound 
      Height          =   255
      Left            =   2160
      TabIndex        =   7
      Top             =   960
      Width           =   255
   End
   Begin VB.CheckBox chkHotKeyActive 
      Height          =   255
      Left            =   2160
      TabIndex        =   5
      Top             =   720
      Width           =   255
   End
   Begin VB.CheckBox chkConfirmHotKey 
      Height          =   255
      Left            =   2160
      TabIndex        =   15
      Top             =   2040
      Width           =   255
   End
   Begin VB.CheckBox chkAvailable 
      Enabled         =   0   'False
      Height          =   255
      Left            =   2160
      TabIndex        =   1
      Top             =   120
      Width           =   255
   End
   Begin VB.CheckBox chkAudibleFeedBack 
      Height          =   255
      Left            =   2160
      TabIndex        =   3
      Top             =   480
      Width           =   255
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   350
      Left            =   6480
      TabIndex        =   50
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label lblRWinLocked 
      Caption         =   "Right Winkey Locked"
      Height          =   255
      Left            =   5160
      TabIndex        =   48
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label lblLWinLocked 
      Caption         =   "Left Winkey Locked"
      Height          =   255
      Left            =   5160
      TabIndex        =   40
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label lblRWinLatched 
      Caption         =   "Right WinKey Latched"
      Height          =   255
      Left            =   2640
      TabIndex        =   32
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label lblLWinLatched 
      Caption         =   "Left WinKey Latched"
      Height          =   255
      Left            =   2640
      TabIndex        =   24
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label lblRShiftLocked 
      Caption         =   "Right Shift Locked"
      Height          =   255
      Left            =   5160
      TabIndex        =   46
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label lblRCtrlLocked 
      Caption         =   "Right Ctrl Locked"
      Height          =   255
      Left            =   5160
      TabIndex        =   44
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label lblRAltLocked 
      Caption         =   "Right Alt Locked"
      Height          =   255
      Left            =   5160
      TabIndex        =   42
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label lblLShiftLocked 
      Caption         =   "Left Shift Locked"
      Height          =   255
      Left            =   5160
      TabIndex        =   38
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label lblLCtrlLocked 
      Caption         =   "Left Ctrl Locked"
      Height          =   255
      Left            =   5160
      TabIndex        =   36
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label lblLAltLocked 
      Caption         =   "Left Alt Locked"
      Height          =   255
      Left            =   5160
      TabIndex        =   34
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label lblRShiftLatched 
      Caption         =   "Right Shift Latched"
      Height          =   255
      Left            =   2640
      TabIndex        =   30
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label lblRCtrlLatched 
      Caption         =   "Right Ctrl Latched"
      Height          =   255
      Left            =   2640
      TabIndex        =   28
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label lblRAltLatched 
      Caption         =   "Right Alt Latched"
      Height          =   255
      Left            =   2640
      TabIndex        =   26
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label lblLShiftLatched 
      Caption         =   "Left Shift Latched"
      Height          =   255
      Left            =   2640
      TabIndex        =   22
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label lblLCtrlLatched 
      Caption         =   "Left Ctrl Latched"
      Height          =   255
      Left            =   2640
      TabIndex        =   20
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label lblLAltLatched 
      Caption         =   "Left Alt Latched"
      Height          =   255
      Left            =   2640
      TabIndex        =   18
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label lblTwoKeysOff 
      Caption         =   "Two Keys Off"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label lblTristate 
      Caption         =   "Tristate"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label lblStickyKeysOn 
      Caption         =   "Sticky Keys On"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label lblIndicator 
      Caption         =   "Indicator"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label lblHotKeySound 
      Caption         =   "Hot Key Sound"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label lblHotKeyActive 
      Caption         =   "Hot Key Active"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label lblConfirmHotKey 
      Caption         =   "Confirm Hotkey"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label lblAvailable 
      Caption         =   "Available"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label lblAudibleFeedBack 
      Caption         =   "Audible Feed Back"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1815
   End
End
Attribute VB_Name = "frmStickyKeys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdApply_Click()
    Dim STICKYKEYS As STICKYKEYS
    STICKYKEYS.cbSize = Len(STICKYKEYS)
        
    Dim AudibleFeedBack As Long
    Dim ConfirmHotKey As Long
    Dim HotKeyActive As Long
    Dim HotKeySound As Long
    Dim Indicator As Long
    Dim StickyKeysOn As Long
    Dim TriState As Long
    Dim TwoKeysOff As Long
    
    
    If chkAudibleFeedBack.value = 1 Then AudibleFeedBack = SKF_AUDIBLEFEEDBACK
    If chkConfirmHotKey.value = 1 Then ConfirmHotKey = SKF_CONFIRMHOTKEY
    If chkHotKeyActive.value = 1 Then HotKeyActive = SKF_HOTKEYACTIVE
    If chkHotKeySound.value = 1 Then HotKeySound = SKF_HOTKEYSOUND
    If chkIndicator.value = 1 Then Indicator = SKF_INDICATOR
    If chkStickyKeysOn.value = 1 Then StickyKeysOn = SKF_STICKYKEYSON
    If chkTristate.value = 1 Then TriState = SKF_TRISTATE
    If chkTwoKeysOff.value = 1 Then TwoKeysOff = SKF_TWOKEYSOFF
    
    STICKYKEYS.dwFlags = AudibleFeedBack Or ConfirmHotKey Or HotKeyActive Or HotKeySound Or Indicator Or StickyKeysOn Or TriState Or TwoKeysOff
    
    If SystemParametersInfo(SPI_SETSTICKYKEYS, STICKYKEYS.cbSize, STICKYKEYS, SPIF_UPDATEINIFILE) = False Then Failed "SystemParametersInfo"
End Sub

Private Sub Form_Load()
    Dim STICKYKEYS As STICKYKEYS
    STICKYKEYS.cbSize = Len(STICKYKEYS)
    
    If SystemParametersInfo(SPI_GETSTICKYKEYS, STICKYKEYS.cbSize, STICKYKEYS, 0) = False Then Failed "SystemParametersInfo"
    
    If STICKYKEYS.dwFlags And SKF_AVAILABLE Then
        With STICKYKEYS
            If .dwFlags And SKF_AVAILABLE Then chkAvailable.value = 1
            If .dwFlags And SKF_AUDIBLEFEEDBACK Then chkAudibleFeedBack.value = 1
            If .dwFlags And SKF_HOTKEYACTIVE Then chkHotKeyActive.value = 1
            If .dwFlags And SKF_HOTKEYSOUND Then chkHotKeySound.value = 1
            If .dwFlags And SKF_STICKYKEYSON Then chkStickyKeysOn.value = 1
            If .dwFlags And SKF_TRISTATE Then chkTristate.value = 1
            If .dwFlags And SKF_TWOKEYSOFF Then chkTwoKeysOff.value = 1
            
            
            If WinVersion(4000000, 5000000, True) = True Then
                If .dwFlags And SKF_CONFIRMHOTKEY Then chkConfirmHotKey.value = 1
                If .dwFlags And SKF_INDICATOR Then chkIndicator.value = 1
            Else
                chkConfirmHotKey.Enabled = False
                chkIndicator.Enabled = False
            End If
        
            If WinVersion(4010000, 5000000, True) = True Then
                If .dwFlags And SKF_LALTLATCHED Then chkLAltLatched.value = 1
                If .dwFlags And SKF_LCTLLATCHED Then chkLCtrlLatched.value = 1
                If .dwFlags And SKF_LSHIFTLATCHED Then chkLShiftLatched.value = 1
                If .dwFlags And SKF_RALTLATCHED Then chkRAltLatched.value = 1
                If .dwFlags And SKF_RCTLLATCHED Then chkRCtrlLatched.value = 1
                If .dwFlags And SKF_RSHIFTLATCHED Then chkRShiftLatched.value = 1
                
                If .dwFlags And SKF_LALTLOCKED Then chkLAltLocked.value = 1
                If .dwFlags And SKF_LCTLLOCKED Then chkLCtrlLocked.value = 1
                If .dwFlags And SKF_LSHIFTLOCKED Then chkLShiftLocked.value = 1
                If .dwFlags And SKF_RALTLOCKED Then chkRAltLocked.value = 1
                If .dwFlags And SKF_RCTLLOCKED Then chkRCtrlLocked.value = 1
                If .dwFlags And SKF_RSHIFTLOCKED Then chkRShiftLocked.value = 1
                
                If .dwFlags And SKF_LWINLATCHED Then chkLWinLatched.value = 1
                If .dwFlags And SKF_RWINLATCHED Then chkRWinLatched.value = 1
                If .dwFlags And SKF_LWINLOCKED Then chkLWinLocked.value = 1
                If .dwFlags And SKF_RWINLOCKED Then chkRWinLocked.value = 1
            Else
                lblLAltLatched.Enabled = False
                lblLCtrlLatched.Enabled = False
                lblLShiftLatched.Enabled = False
                lblRAltLatched.Enabled = False
                lblRCtrlLatched.Enabled = False
                lblRShiftLatched.Enabled = False
                lblLWinLatched.Enabled = False
                lblRWinLatched.Enabled = False
                
                lblLAltLocked.Enabled = False
                lblLCtrlLocked.Enabled = False
                lblLShiftLocked.Enabled = False
                lblRAltLocked.Enabled = False
                lblRCtrlLocked.Enabled = False
                lblRShiftLocked.Enabled = False
                lblLWinLocked.Enabled = False
                lblRWinLocked.Enabled = False
            End If
        End With
    Else
        lblAudibleFeedBack.Enabled = False
        chkAudibleFeedBack.Enabled = False
        lblHotKeyActive.Enabled = False
        chkHotKeyActive.Enabled = False
        lblHotKeySound.Enabled = False
        chkHotKeySound.Enabled = False
        lblStickyKeysOn.Enabled = False
        chkStickyKeysOn.Enabled = False
        lblTristate.Enabled = False
        chkTristate.Enabled = False
        lblTwoKeysOff.Enabled = False
        chkTwoKeysOff.Enabled = False
        lblConfirmHotKey.Enabled = False
        chkConfirmHotKey.Enabled = False
        lblIndicator.Enabled = False
        chkIndicator.Enabled = False
        cmdApply.Enabled = False
        
        lblLAltLatched.Enabled = False
        lblLCtrlLatched.Enabled = False
        lblLShiftLatched.Enabled = False
        lblRAltLatched.Enabled = False
        lblRCtrlLatched.Enabled = False
        lblRShiftLatched.Enabled = False
        lblLWinLatched.Enabled = False
        lblRWinLatched.Enabled = False
        
        lblLAltLocked.Enabled = False
        lblLCtrlLocked.Enabled = False
        lblLShiftLocked.Enabled = False
        lblRAltLocked.Enabled = False
        lblRCtrlLocked.Enabled = False
        lblRShiftLocked.Enabled = False
        lblLWinLocked.Enabled = False
        lblRWinLocked.Enabled = False
    End If
End Sub
