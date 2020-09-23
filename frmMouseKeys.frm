VERSION 5.00
Begin VB.Form frmMouseKeys 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mouse Keys"
   ClientHeight    =   2955
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6015
   Icon            =   "frmMouseKeys.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   6015
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtCtrlSpeed 
      Height          =   285
      Left            =   4800
      TabIndex        =   27
      Top             =   1440
      Width           =   1095
   End
   Begin VB.TextBox txtTimeToMaxSpeed 
      Height          =   285
      Left            =   4800
      TabIndex        =   31
      Top             =   2160
      Width           =   1095
   End
   Begin VB.TextBox txtMaxSpeed 
      Height          =   285
      Left            =   4800
      TabIndex        =   29
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CheckBox chkLeftButtonDown 
      Enabled         =   0   'False
      Height          =   255
      Left            =   5640
      TabIndex        =   23
      Top             =   840
      Width           =   255
   End
   Begin VB.CheckBox chkRightButtonDown 
      Enabled         =   0   'False
      Height          =   255
      Left            =   5640
      TabIndex        =   25
      Top             =   1080
      Width           =   255
   End
   Begin VB.CheckBox chkRightButtonSelect 
      Enabled         =   0   'False
      Height          =   255
      Left            =   5640
      TabIndex        =   21
      Top             =   600
      Width           =   255
   End
   Begin VB.CheckBox chkLeftButtonSelect 
      Enabled         =   0   'False
      Height          =   255
      Left            =   5640
      TabIndex        =   19
      Top             =   360
      Width           =   255
   End
   Begin VB.CheckBox chkReplaceNumbers 
      Height          =   255
      Left            =   2640
      TabIndex        =   15
      Top             =   2040
      Width           =   255
   End
   Begin VB.CheckBox chkMouseMode 
      Enabled         =   0   'False
      Height          =   255
      Left            =   5640
      TabIndex        =   17
      Top             =   120
      Width           =   255
   End
   Begin VB.CheckBox chkModifiers 
      Height          =   255
      Left            =   2640
      TabIndex        =   13
      Top             =   1800
      Width           =   255
   End
   Begin VB.CheckBox chkMouseKeysOn 
      Height          =   255
      Left            =   2640
      TabIndex        =   7
      Top             =   960
      Width           =   255
   End
   Begin VB.CheckBox chkIndicator 
      Height          =   255
      Left            =   2640
      TabIndex        =   11
      Top             =   1560
      Width           =   255
   End
   Begin VB.CheckBox chkHotKeySound 
      Height          =   255
      Left            =   2640
      TabIndex        =   5
      Top             =   720
      Width           =   255
   End
   Begin VB.CheckBox chkHotKeyActive 
      Height          =   255
      Left            =   2640
      TabIndex        =   3
      Top             =   480
      Width           =   255
   End
   Begin VB.CheckBox chkConfirmHotKey 
      Height          =   255
      Left            =   2640
      TabIndex        =   9
      Top             =   1320
      Width           =   255
   End
   Begin VB.CheckBox chkAvailable 
      Enabled         =   0   'False
      Height          =   255
      Left            =   2640
      TabIndex        =   1
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   350
      Left            =   4920
      TabIndex        =   32
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label lblCtrlSpeed 
      Caption         =   "Ctrl Speed"
      Height          =   255
      Left            =   3120
      TabIndex        =   26
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label lblTimeToMaxSpeed 
      Caption         =   "Time To Max Speed"
      Height          =   255
      Left            =   3120
      TabIndex        =   30
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label lblMaxSpeed 
      Caption         =   "Max Speed"
      Height          =   255
      Left            =   3120
      TabIndex        =   28
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label lblLeftButtonDown 
      Caption         =   "Left Button Down"
      Height          =   255
      Left            =   3120
      TabIndex        =   22
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label lblRightButtonDown 
      Caption         =   "Right Button Down"
      Height          =   255
      Left            =   3120
      TabIndex        =   24
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label lblRightButtonSelect 
      Caption         =   "Right Button Select"
      Height          =   255
      Left            =   3120
      TabIndex        =   20
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label lblLeftButtonSelect 
      Caption         =   "Left Button Select"
      Height          =   255
      Left            =   3120
      TabIndex        =   18
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label lblReplaceNumbers 
      Caption         =   "Replace Numbers"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label lblMouseMode 
      Caption         =   "Mouse Mode"
      Height          =   255
      Left            =   3120
      TabIndex        =   16
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label lblModifiers 
      Caption         =   "Modifiers"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label lblMouseKeysOn 
      Caption         =   "Mouse Keys On"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label lblIndicator 
      Caption         =   "Indicator"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label lblHotKeySound 
      Caption         =   "Hot Key Sound"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label lblHotKeyActive 
      Caption         =   "Hot Key Active"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label lblConfirmHotKey 
      Caption         =   "Confirm Hot Key"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label lblAvailable 
      Caption         =   "Available"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmMouseKeys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdApply_Click()
    Dim MOUSEKEYS As MOUSEKEYS
    MOUSEKEYS.cbSize = Len(MOUSEKEYS)
        
    Dim ConfirmHotKey As Long
    Dim HotKeyActive As Long
    Dim HotKeySound As Long
    Dim Indicator As Long
    Dim MouseKeysOn As Long
    Dim Modifiers As Long
    Dim ReplaceNumbers As Long
    
    
    If chkConfirmHotKey.value = 1 Then ConfirmHotKey = MKF_CONFIRMHOTKEY
    If chkHotKeyActive.value = 1 Then HotKeyActive = MKF_HOTKEYACTIVE
    If chkHotKeySound.value = 1 Then HotKeySound = MKF_HOTKEYSOUND
    If chkIndicator.value = 1 Then Indicator = MKF_INDICATOR
    If chkMouseKeysOn.value = 1 Then MouseKeysOn = MKF_MOUSEKEYSON
    If chkModifiers.value = 1 Then Modifiers = MKF_MODIFIERS
    If chkReplaceNumbers.value = 1 Then ReplaceNumbers = MKF_REPLACENUMBERS
    
    With MOUSEKEYS
        .dwFlags = ConfirmHotKey Or HotKeyActive Or HotKeySound Or Indicator Or MouseKeysOn Or Modifiers Or ReplaceNumbers
        
        .iCtrlSpeed = Val(txtCtrlSpeed.Text)
        .iMaxSpeed = Val(txtMaxSpeed.Text)
        .iTimeToMaxSpeed = Val(txtTimeToMaxSpeed.Text)
    End With
    
    If SystemParametersInfo(SPI_SETMOUSEKEYS, MOUSEKEYS.cbSize, MOUSEKEYS, SPIF_UPDATEINIFILE) = False Then Failed "SystemParametersInfo"
End Sub

Private Sub Form_Load()
    Dim MOUSEKEYS As MOUSEKEYS
    MOUSEKEYS.cbSize = Len(MOUSEKEYS)
    
    If SystemParametersInfo(SPI_GETMOUSEKEYS, MOUSEKEYS.cbSize, MOUSEKEYS, 0) = False Then Failed "SystemParametersInfo"
    
    If MOUSEKEYS.dwFlags And MKF_AVAILABLE Then
        With MOUSEKEYS
            If .dwFlags And MKF_AVAILABLE Then chkAvailable.value = 1
            If .dwFlags And MKF_HOTKEYACTIVE Then chkHotKeyActive.value = 1
            If .dwFlags And MKF_HOTKEYSOUND Then chkHotKeySound.value = 1
            
            If WinVersion(4000000, 5000000, True) = True Then
                If .dwFlags And MKF_CONFIRMHOTKEY Then chkConfirmHotKey.value = 1
                If .dwFlags And MKF_INDICATOR Then chkIndicator.value = 1
                If .dwFlags And MKF_MODIFIERS Then chkModifiers.value = 1
                If .dwFlags And MKF_REPLACENUMBERS Then chkReplaceNumbers.value = 1
            Else
                lblConfirmHotKey.Enabled = False
                chkConfirmHotKey.Enabled = False
                lblIndicator.Enabled = False
                chkIndicator.Enabled = False
                lblModifiers.Enabled = False
                chkModifiers.Enabled = False
                lblReplaceNumbers.Enabled = False
                chkReplaceNumbers.Enabled = False
            End If
            If WinVersion(4010000, 5000000, True) = True Then
                If .dwFlags And MKF_MOUSEMODE Then chkMouseMode.value = 1
                If .dwFlags And MKF_LEFTBUTTONSEL Then chkLeftButtonSelect.value = 1
                If .dwFlags And MKF_RIGHTBUTTONSEL Then chkRightButtonSelect.value = 1
                If .dwFlags And MKF_LEFTBUTTONDOWN Then chkLeftButtonDown.value = 1
                If .dwFlags And MKF_RIGHTBUTTONDOWN Then chkRightButtonDown.value = 1
            Else
                lblMouseMode.Enabled = False
                lblLeftButtonSelect.Enabled = False
                lblRightButtonSelect.Enabled = False
                lblLeftButtonDown.Enabled = False
                lblRightButtonDown.Enabled = False
            End If
            
            txtCtrlSpeed.Text = CStr(.iCtrlSpeed)
            txtMaxSpeed.Text = CStr(.iMaxSpeed)
            txtTimeToMaxSpeed.Text = CStr(.iTimeToMaxSpeed)
        End With
    Else
        lblHotKeyActive.Enabled = False
        chkHotKeyActive.Enabled = False
        lblHotKeySound.Enabled = False
        chkHotKeySound.Enabled = False
        lblMouseKeysOn.Enabled = False
        chkMouseKeysOn.Enabled = False
        
        lblConfirmHotKey.Enabled = False
        chkConfirmHotKey.Enabled = False
        lblIndicator.Enabled = False
        chkIndicator.Enabled = False
        lblModifiers.Enabled = False
        chkModifiers.Enabled = False
        lblReplaceNumbers.Enabled = False
        chkReplaceNumbers.Enabled = False
                
        lblMouseMode.Enabled = False
        lblLeftButtonSelect.Enabled = False
        lblRightButtonSelect.Enabled = False
        lblLeftButtonDown.Enabled = False
        lblRightButtonDown.Enabled = False
                
        cmdApply.Enabled = False
    End If
End Sub

Private Sub txtCtrlSpeed_Change()
    txtCtrlSpeed.Text = CStr(Val(Rem_NonNumeric_Chr(txtCtrlSpeed.Text)))
    
    If WinVersion(-1, 0, True) = True Then
        If Val(txtCtrlSpeed.Text) < 10 Then txtCtrlSpeed.Text = "10"
        If Val(txtCtrlSpeed.Text) > 360 Then txtCtrlSpeed.Text = "360"
    Else
        If Val(txtCtrlSpeed.Text) < 0 Then txtCtrlSpeed.Text = "0"
        If Val(txtCtrlSpeed.Text) > 2147483647 Then txtCtrlSpeed.Text = "2147483647"
    End If
End Sub

Private Sub txtMaxSpeed_Change()
    txtMaxSpeed.Text = CStr(Val(Rem_NonNumeric_Chr(txtMaxSpeed.Text)))
    If Val(txtMaxSpeed.Text) < 0 Then txtMaxSpeed.Text = "0"
    If Val(txtMaxSpeed.Text) > 2147483647 Then txtMaxSpeed.Text = "2147483647"
End Sub

Private Sub txtTimeToMaxSpeed_Change()
    txtTimeToMaxSpeed.Text = CStr(Val(Rem_NonNumeric_Chr(txtTimeToMaxSpeed.Text)))
    If Val(txtTimeToMaxSpeed.Text) < 1000 Then txtTimeToMaxSpeed.Text = "1000"
    If Val(txtTimeToMaxSpeed.Text) > 5000 Then txtTimeToMaxSpeed.Text = "5000"
End Sub
