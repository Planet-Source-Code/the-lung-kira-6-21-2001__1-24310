VERSION 5.00
Begin VB.Form frmFilterKeys 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Filter Keys"
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5535
   Icon            =   "frmFilterKeys.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   5535
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtBounce 
      Height          =   285
      Left            =   4200
      TabIndex        =   15
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox txtRepeat 
      Height          =   285
      Left            =   4200
      TabIndex        =   19
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox txtDelay 
      Height          =   285
      Left            =   4200
      TabIndex        =   17
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox txtWait 
      Height          =   285
      Left            =   4200
      TabIndex        =   21
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CheckBox chkIndicator 
      Height          =   255
      Left            =   2160
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1800
      Width           =   255
   End
   Begin VB.CheckBox chkConfirmHotKey 
      Height          =   255
      Left            =   2160
      TabIndex        =   11
      Top             =   1560
      Width           =   255
   End
   Begin VB.CheckBox chkHotKeySound 
      Height          =   255
      Left            =   2160
      TabIndex        =   9
      Top             =   1200
      Width           =   255
   End
   Begin VB.CheckBox chkHotKeyActive 
      Height          =   255
      Left            =   2160
      TabIndex        =   7
      Top             =   960
      Width           =   255
   End
   Begin VB.CheckBox chkFilterKeysOn 
      Height          =   255
      Left            =   2160
      TabIndex        =   5
      Top             =   720
      Width           =   255
   End
   Begin VB.CheckBox chkClickOn 
      Height          =   255
      Left            =   2160
      TabIndex        =   3
      Top             =   480
      Width           =   255
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   350
      Left            =   4440
      TabIndex        =   22
      Top             =   1680
      Width           =   975
   End
   Begin VB.CheckBox chkAvailable 
      Enabled         =   0   'False
      Height          =   255
      Left            =   2160
      TabIndex        =   1
      Top             =   120
      Width           =   255
   End
   Begin VB.Label lblBounce 
      Caption         =   "Bounce"
      Height          =   255
      Left            =   2640
      TabIndex        =   14
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label lblRepeat 
      Caption         =   "Repeat"
      Height          =   255
      Left            =   2640
      TabIndex        =   18
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label lblDelay 
      Caption         =   "Delay"
      Height          =   255
      Left            =   2640
      TabIndex        =   16
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label lblWait 
      Caption         =   "Wait"
      Height          =   255
      Left            =   2640
      TabIndex        =   20
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label lblIndicator 
      Caption         =   "Indicator"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label lblConfirmHotKey 
      Caption         =   "Confirm Hot Key"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label lblHotKeySound 
      Caption         =   "Hot Key Sound"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label lblHotKeyActive 
      Caption         =   "Hot Key Active"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label lblFilterKeysOn 
      Caption         =   "Filter Keys On"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label lblClickOn 
      Caption         =   "Click On"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label lblAvailable 
      Caption         =   "Available"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmFilterKeys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdApply_Click()
    Dim FILTERKEYS As FILTERKEYS
    FILTERKEYS.cbSize = Len(FILTERKEYS)
        
    Dim ClickOn As Long
    Dim FilterKeysOn As Long
    Dim HotKeyActive As Long
    Dim HotKeySound As Long
    Dim ConfirmHotKey As Long
    Dim Indicator As Long
    
    
    If chkClickOn.value = 1 Then ClickOn = FKF_CLICKON
    If chkFilterKeysOn.value = 1 Then FilterKeysOn = FKF_FILTERKEYSON
    If chkHotKeyActive.value = 1 Then HotKeyActive = FKF_HOTKEYACTIVE
    If chkHotKeySound.value = 1 Then HotKeySound = FKF_HOTKEYSOUND
    If chkConfirmHotKey.value = 1 Then ConfirmHotKey = FKF_CONFIRMHOTKEY
    If chkIndicator.value = 1 Then Indicator = FKF_INDICATOR
    
    With FILTERKEYS
        .dwFlags = ClickOn Or FilterKeysOn Or HotKeyActive Or HotKeySound Or ConfirmHotKey Or Indicator
        
        .iBounceMSec = Val(txtBounce.Text)
        .iDelayMSec = Val(txtDelay.Text)
        .iRepeatMSec = Val(txtRepeat.Text)
        .iWaitMSec = Val(txtWait.Text)
    End With
    
    If SystemParametersInfo(SPI_SETFILTERKEYS, FILTERKEYS.cbSize, FILTERKEYS, SPIF_UPDATEINIFILE) = False Then Failed "SystemParametersInfo"
End Sub

Private Sub Form_Load()
    Dim FILTERKEYS As FILTERKEYS
    FILTERKEYS.cbSize = Len(FILTERKEYS)
    
    If SystemParametersInfo(SPI_GETFILTERKEYS, FILTERKEYS.cbSize, FILTERKEYS, 0) = False Then Failed "SystemParametersInfo"
    
    If FILTERKEYS.dwFlags And FKF_AVAILABLE Then
        With FILTERKEYS
            If .dwFlags And FKF_AVAILABLE Then chkAvailable.value = 1
            If .dwFlags And FKF_CLICKON Then chkClickOn.value = 1
            If .dwFlags And FKF_FILTERKEYSON Then chkFilterKeysOn.value = 1
            If .dwFlags And FKF_HOTKEYACTIVE Then chkHotKeyActive.value = 1
            If .dwFlags And FKF_HOTKEYSOUND Then chkHotKeySound.value = 1
            
            If WinVersion(4000000, 5000000, True) = True Then
                If .dwFlags And FKF_CONFIRMHOTKEY Then chkConfirmHotKey.value = 1
                If .dwFlags And FKF_INDICATOR Then chkIndicator.value = 1
            Else
                chkConfirmHotKey.Enabled = False
                chkIndicator.Enabled = False
            End If
            
            txtWait.Text = CStr(.iWaitMSec)
            txtDelay.Text = CStr(.iDelayMSec)
            txtRepeat.Text = CStr(.iRepeatMSec)
            txtBounce.Text = CStr(.iBounceMSec)
        End With
    Else
        lblClickOn.Enabled = False
        chkClickOn.Enabled = False
        lblFilterKeysOn.Enabled = False
        chkFilterKeysOn.Enabled = False
        lblHotKeyActive.Enabled = False
        chkHotKeyActive.Enabled = False
        lblHotKeySound.Enabled = False
        chkHotKeySound.Enabled = False
        lblConfirmHotKey.Enabled = False
        chkConfirmHotKey.Enabled = False
        lblIndicator.Enabled = False
        chkIndicator.Enabled = False
        lblWait.Enabled = False
        txtWait.Enabled = False
        lblDelay.Enabled = False
        txtDelay.Enabled = False
        lblRepeat.Enabled = False
        txtRepeat.Enabled = False
        lblBounce.Enabled = False
        txtBounce.Enabled = False
        cmdApply.Enabled = False
    End If
End Sub

Private Sub txtBounce_Change()
    txtBounce.Text = CStr(Val(Rem_NonNumeric_Chr(txtBounce.Text)))
    If Val(txtBounce.Text) < 0 Then txtBounce.Text = "0"
    If Val(txtBounce.Text) > 2147483647 Then txtBounce.Text = "2147483647"
End Sub

Private Sub txtDelay_Change()
    txtDelay.Text = CStr(Val(Rem_NonNumeric_Chr(txtDelay.Text)))
    If Val(txtDelay.Text) < 0 Then txtDelay.Text = "0"
    If Val(txtDelay.Text) > 2147483647 Then txtDelay.Text = "2147483647"
End Sub

Private Sub txtRepeat_Change()
    txtRepeat.Text = CStr(Val(Rem_NonNumeric_Chr(txtRepeat.Text)))
    If Val(txtRepeat.Text) < 0 Then txtRepeat.Text = "0"
    If Val(txtRepeat.Text) > 2147483647 Then txtRepeat.Text = "2147483647"
End Sub

Private Sub txtWait_Change()
    txtWait.Text = CStr(Val(Rem_NonNumeric_Chr(txtWait.Text)))
    If Val(txtWait.Text) < 0 Then txtWait.Text = "0"
    If Val(txtWait.Text) > 2147483647 Then txtWait.Text = "2147483647"
End Sub
