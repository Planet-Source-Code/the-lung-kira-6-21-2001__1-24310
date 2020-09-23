VERSION 5.00
Begin VB.Form frmKeyboardSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Keyboard Settings"
   ClientHeight    =   3555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3615
   Icon            =   "frmKeyboardSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   3615
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cboLanguageToggle 
      Height          =   315
      Left            =   1920
      TabIndex        =   7
      Top             =   1080
      Width           =   1575
   End
   Begin VB.CheckBox chkCues 
      Height          =   255
      Left            =   3240
      TabIndex        =   3
      Top             =   480
      Width           =   255
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   350
      Left            =   2520
      TabIndex        =   16
      Top             =   3120
      Width           =   975
   End
   Begin VB.HScrollBar hsRepeatDelay 
      Height          =   255
      LargeChange     =   5
      Left            =   720
      Max             =   3
      TabIndex        =   10
      Top             =   1920
      Width           =   2175
   End
   Begin VB.HScrollBar hsRepeatRate 
      Height          =   255
      LargeChange     =   5
      Left            =   720
      Max             =   31
      TabIndex        =   14
      Top             =   2760
      Width           =   2175
   End
   Begin VB.CheckBox chkPref 
      Height          =   255
      Left            =   3240
      TabIndex        =   5
      Top             =   720
      Width           =   255
   End
   Begin VB.TextBox txtBlinkRate 
      Height          =   285
      Left            =   1920
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label lblLanguageToggle 
      Caption         =   "Language Toggle"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label lblCues 
      Caption         =   "Cues Underlined"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label lblPref 
      Caption         =   "Keyboard Preference"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label lblFast 
      Caption         =   "Fast"
      Height          =   255
      Left            =   3120
      TabIndex        =   15
      Top             =   2760
      Width           =   375
   End
   Begin VB.Label lblSlow 
      Caption         =   "Slow"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   2760
      Width           =   375
   End
   Begin VB.Label lblLong 
      Caption         =   "Long"
      Height          =   255
      Left            =   3120
      TabIndex        =   11
      Top             =   1920
      Width           =   375
   End
   Begin VB.Label lblShort 
      Caption         =   "Short"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1920
      Width           =   375
   End
   Begin VB.Label lblRepeatRate 
      Caption         =   "Repeat Rate"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label lblRepeatDelay 
      Caption         =   "Repeat Delay"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label lblBlinkRate 
      Caption         =   "Caret Blink Rate"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmKeyboardSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdApply_Click()
    If SetCaretBlinkTime(Val(txtBlinkRate.Text)) = False Then Failed "SetCaretBlinkTime"
    
    If WinVersion(4010000, 5000000, True) = True Then
        If SystemParametersInfo(SPI_SETKEYBOARDCUES, 0, chkCues.value, SPIF_UPDATEINIFILE) = 0 Then Failed "SystemParametersInfo"
    End If
    If WinVersion(0, 5000000, True) = True Then
        If SystemParametersInfo(SPI_SETKEYBOARDPREF, 0, chkPref.value, SPIF_UPDATEINIFILE) = 0 Then Failed "SystemParametersInfo"
    End If
    
    If SystemParametersInfo(SPI_SETKEYBOARDDELAY, hsRepeatDelay.value, 0, SPIF_UPDATEINIFILE) = 0 Then Failed "SystemParametersInfo"
    If SystemParametersInfo(SPI_SETKEYBOARDSPEED, hsRepeatRate.value, 0, SPIF_UPDATEINIFILE) = 0 Then Failed "SystemParametersInfo"
    
    SaveRegSetting HKEY_CURRENT_USER, "Keyboard Layout\Toggle", "Hotkey", CStr(cboLanguageToggle.ListIndex + 1), REG_SZ
    If SystemParametersInfo(SPI_SETLANGTOGGLE, 0, 0, SPIF_UPDATEINIFILE) = 0 Then Failed "SystemParametersInfo"
End Sub

Private Sub Form_Load()
    With cboLanguageToggle
        .AddItem "ALT+SHIFT"
        .AddItem "CTRL+SHIFT"
        .AddItem "None"
    End With
    
    
    txtBlinkRate.Text = CStr(GetCaretBlinkTime)
    
    
    If WinVersion(4010000, 5000000, True) = True Then
        Dim boolKBCues As Boolean
        If SystemParametersInfo(SPI_GETKEYBOARDCUES, 0, boolKBCues, 0) = 0 Then Failed "SystemParametersInfo"
        chkCues.value = boolKBCues
    Else
        lblCues.Enabled = False
        chkCues.Enabled = False
    End If
    
    
    Dim intDelay As Integer
    If SystemParametersInfo(SPI_GETKEYBOARDDELAY, 0, intDelay, 0) = 0 Then Failed "SystemParametersInfo"
    hsRepeatDelay.value = intDelay
    
    
    If WinVersion(0, 5000000, True) = True Then
        Dim boolPref As Boolean
        If SystemParametersInfo(SPI_GETKEYBOARDPREF, 0, boolPref, 0) = 0 Then Failed "SystemParametersInfo"
        chkPref.value = boolPref
    Else
        lblPref.Enabled = False
        chkPref.Enabled = False
    End If
    
    
    Dim lngSpeed As Long
    If SystemParametersInfo(SPI_GETKEYBOARDSPEED, 0, lngSpeed, 0) = 0 Then Failed "SystemParametersInfo"
    hsRepeatRate.value = lngSpeed
    
    
    Select Case GetRegSetting(HKEY_CURRENT_USER, "Keyboard Layout\Toggle", "Hotkey")
        Case "1": cboLanguageToggle.ListIndex = 0
        Case "2": cboLanguageToggle.ListIndex = 1
        Case "3": cboLanguageToggle.ListIndex = 2
        Case Else: cboLanguageToggle.ListIndex = 2
    End Select
End Sub

Private Sub txtBlinkRate_Change()
    txtBlinkRate.Text = CStr(Val(Rem_NonNumeric_Chr(txtBlinkRate.Text)))
    If Val(txtBlinkRate.Text) < 0 Then txtBlinkRate.Text = "0"
    If Val(txtBlinkRate.Text) > 5000 Then txtBlinkRate.Text = "5000"
End Sub
