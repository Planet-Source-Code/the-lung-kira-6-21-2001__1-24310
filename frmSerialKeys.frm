VERSION 5.00
Begin VB.Form frmSerialKeys 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Serial Keys"
   ClientHeight    =   2940
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3015
   Icon            =   "frmSerialKeys.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   3015
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cboPortState 
      Height          =   315
      Left            =   1680
      TabIndex        =   13
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox txtActive 
      Height          =   285
      Left            =   1680
      TabIndex        =   11
      Top             =   1440
      Width           =   1215
   End
   Begin VB.ComboBox cboBaudRate 
      Height          =   315
      Left            =   1680
      TabIndex        =   9
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CheckBox chkAvailable 
      Enabled         =   0   'False
      Height          =   255
      Left            =   2640
      TabIndex        =   4
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   350
      Left            =   1920
      TabIndex        =   3
      Top             =   2520
      Width           =   975
   End
   Begin VB.CheckBox chkSerialKeysOn 
      Height          =   255
      Left            =   2640
      TabIndex        =   2
      Top             =   720
      Width           =   255
   End
   Begin VB.CheckBox chkIndicator 
      Height          =   255
      Left            =   2640
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   480
      Width           =   255
   End
   Begin VB.TextBox txtActivePort 
      Height          =   285
      Left            =   1680
      TabIndex        =   0
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label lblPortState 
      Caption         =   "Port State"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label lblActive 
      Caption         =   "Active"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label lblBaudRate 
      Caption         =   "Baud Rate"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label lblAvailable 
      Caption         =   "Available"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label lblSerialKeysOn 
      Caption         =   "Serial Keys On"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label lblIndicator 
      Caption         =   "Indicator"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label lblActivePort 
      Caption         =   "Active Port"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   1335
   End
End
Attribute VB_Name = "frmSerialKeys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdApply_Click()
    Dim SERIALKEYS As SERIALKEYS
    SERIALKEYS.cbSize = Len(SERIALKEYS)
    
    Dim Indicator As Long
    Dim SerialKeysOn As Long
    
    
    If chkIndicator.value = 1 Then Indicator = SERKF_INDICATOR
    If chkSerialKeysOn.value = 1 Then SerialKeysOn = SERKF_SERIALKEYSON
    
    With SERIALKEYS
        .dwFlags = Indicator Or SerialKeysOn
        
        .lpszActivePort = txtActivePort.Text
        .iActive = Val(txtActive.Text)
        .iBaudRate = cboBaudRate.List(cboBaudRate.ListIndex)
        .iPortState = cboPortState.ListIndex
    End With
    
    If SystemParametersInfo(SPI_SETSERIALKEYS, SERIALKEYS.cbSize, SERIALKEYS, SPIF_UPDATEINIFILE) = False Then Failed "SystemParametersInfo"
End Sub

Private Sub Form_Load()
    With cboBaudRate
        .AddItem "110"
        .AddItem "300"
        .AddItem "600"
        .AddItem "1200"
        .AddItem "2400"
        .AddItem "4800"
        .AddItem "9600"
        .AddItem "14400"
        .AddItem "19200"
        .AddItem "38400"
        .AddItem "56000"
        .AddItem "57600"
        .AddItem "115200"
        .AddItem "128000"
        .AddItem "256000"
    End With
    With cboPortState
        .AddItem "0 Ignore"
        .AddItem "1 Watch"
        .AddItem "2 Input"
    End With


    Dim SERIALKEYS As SERIALKEYS
    SERIALKEYS.cbSize = Len(SERIALKEYS)
    
    If SystemParametersInfo(SPI_GETSERIALKEYS, SERIALKEYS.cbSize, SERIALKEYS, 0) = False Then Failed "SystemParametersInfo"
    
    If SERIALKEYS.dwFlags And SERKF_AVAILABLE Then
        With SERIALKEYS
            If .dwFlags And SERKF_AVAILABLE Then chkAvailable.value = 1
            If .dwFlags And SERKF_INDICATOR Then chkIndicator.value = 1
            If .dwFlags And SERKF_SERIALKEYSON Then chkSerialKeysOn.value = 1
            
            
            txtActivePort.Text = .lpszActivePort
            txtActive.Text = CStr(.iActive)
            
            cboBaudRate.ListIndex = .iBaudRate
            cboPortState.ListIndex = .iPortState
        End With
    Else
        lblIndicator.Enabled = False
        chkIndicator.Enabled = False
        lblSerialKeysOn.Enabled = False
        chkSerialKeysOn.Enabled = False
        lblActivePort.Enabled = False
        txtActivePort.Enabled = False
        lblActive.Enabled = False
        txtActive.Enabled = False
        lblBaudRate.Enabled = False
        cboBaudRate.Enabled = False
        lblPortState.Enabled = False
        cboPortState.Enabled = False
        cmdApply.Enabled = False
    End If
End Sub

Private Sub txtActive_Change()
    txtActive.Text = CStr(Val(Rem_NonNumeric_Chr(txtActive.Text)))
    If Val(txtActive.Text) < 0 Then txtActive.Text = "0"
    If Val(txtActive.Text) > 2147483647 Then txtActive.Text = "2147483647"
End Sub
