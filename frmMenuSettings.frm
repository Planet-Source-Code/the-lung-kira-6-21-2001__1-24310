VERSION 5.00
Begin VB.Form frmMenuSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Menu Settings"
   ClientHeight    =   1995
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2775
   Icon            =   "frmMenuSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1995
   ScaleWidth      =   2775
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkMenuAnimation 
      Alignment       =   1  'Right Justify
      Caption         =   "Check1"
      Height          =   255
      Left            =   2400
      TabIndex        =   3
      Top             =   480
      Width           =   255
   End
   Begin VB.ComboBox cboDropAlignment 
      Height          =   315
      Left            =   1560
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.CheckBox chkMenuFade 
      Alignment       =   1  'Right Justify
      Caption         =   "Check1"
      Height          =   255
      Left            =   2400
      TabIndex        =   5
      Top             =   840
      Width           =   255
   End
   Begin VB.TextBox txtShowDelay 
      Height          =   285
      Left            =   1560
      TabIndex        =   7
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   350
      Left            =   1680
      TabIndex        =   8
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label lblMenuAnimation 
      Caption         =   "Menu Animation"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label lblMenuFade 
      Caption         =   "Menu Fade"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label lblDropAlignment 
      Caption         =   "Drop Alignment"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblShowDelay 
      Caption         =   "Show Delay"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   1215
   End
End
Attribute VB_Name = "frmMenuSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdApply_Click()
    Dim boolAlign As Boolean
    Select Case cboDropAlignment.ListIndex
        Case 0: boolAlign = False
        Case 1: boolAlign = True
    End Select
    If SystemParametersInfo(SPI_SETMENUDROPALIGNMENT, boolAlign, 0, SPIF_UPDATEINIFILE) = 0 Then Failed "SystemParametersInfo"
    
    'If WinVersion(-1, 5010000, True) = True Then
    '    If SystemParametersInfo(SPI_sETFLATMENU, 0, ByVal CBool(chkFlatMenu.value), SPIF_UPDATEINIFILE) = 0 Then Failed "SystemParametersInfo"
    'End If
    If WinVersion(4010000, 5000000, True) = True Then
        If SystemParametersInfo(SPI_SETMENUANIMATION, 0, ByVal CBool(chkMenuAnimation.value), SPIF_UPDATEINIFILE) = 0 Then Failed "SystemParametersInfo"
    End If
    If WinVersion(-1, 5000000, True) = True Then
        If SystemParametersInfo(SPI_SETMENUFADE, 0, ByVal CBool(chkMenuFade.value), SPIF_UPDATEINIFILE) = 0 Then Failed "SystemParametersInfo"
        'If SystemParametersInfo(SPI_SETSELECTIONFADE, 0, CBool(chkSelectionFade.value), SPIF_UPDATEINIFILE) = 0 Then Failed "SystemParametersInfo"
    End If
    If WinVersion(4010000, 0, True) = True Then
        SaveRegSetting HKEY_CURRENT_USER, "Control Panel\Desktop", "MenuShowDelay", txtShowDelay.Text, REG_SZ
        If SystemParametersInfo(SPI_SETMENUSHOWDELAY, Val(txtShowDelay.Text), 0, SPIF_UPDATEINIFILE) = 0 Then Failed "SystemParametersInfo"
    End If
End Sub

Private Sub Form_Load()
    With cboDropAlignment
        .AddItem "Left"
        .AddItem "Right"
    End With
    

    Dim boolValue As Boolean
    
    If SystemParametersInfo(SPI_GETMENUDROPALIGNMENT, 0, boolValue, 0) = 0 Then Failed "SystemParametersInfo"
    If boolValue = True Then 'True = left
        cboDropAlignment.ListIndex = 0
    Else 'False = right
        cboDropAlignment.ListIndex = 1
    End If

    'If WinVersion(-1, 5010000, True) = True Then
    '    If SystemParametersInfo(SPI_gETFLATMENU, 0, boolValue, 0) = 0 Then Failed "SystemParametersInfo"
    '    chkFlatMenu.value = boolValue
    'Else
    '    lblFlatMenu.Enabled = False
    '    chkFlatMenu.Enabled = False
    'End If
    If WinVersion(4010000, 5000000, True) = True Then
        If SystemParametersInfo(SPI_GETMENUANIMATION, 0, boolValue, 0) = 0 Then Failed "SystemParametersInfo"
        chkMenuAnimation.value = boolValue
    Else
        lblMenuAnimation.Enabled = False
        chkMenuAnimation.Enabled = False
    End If
    If WinVersion(-1, 5000000, True) = True Then
        If SystemParametersInfo(SPI_GETMENUFADE, 0, boolValue, 0) = 0 Then Failed "SystemParametersInfo"
        chkMenuFade.value = boolValue
        
        'If SystemParametersInfo(SPI_GETSELECTIONFADE, 0, boolValue, 0) = 0 Then Failed "SystemParametersInfo"
        'chkSelectionFade.value = boolValue
    Else
        lblMenuFade.Enabled = False
        chkMenuFade.Enabled = False
        'lblSelectionFade.Enabled = False
        'chkSelectionFade.Enabled = False
    End If
    If WinVersion(4010000, 0, True) = True Then
        Dim lngDelay As Long
        If SystemParametersInfo(SPI_GETMENUSHOWDELAY, 0, lngDelay, 0) = 0 Then Failed "SystemParametersInfo"
        txtShowDelay.Text = CStr(lngDelay)
    Else
        lblShowDelay.Enabled = False
        txtShowDelay.Enabled = False
    End If
End Sub

Private Sub txtShowDelay_Change()
    txtShowDelay.Text = CStr(Val(Rem_NonNumeric_Chr(txtShowDelay.Text)))
    If Val(txtShowDelay.Text) < 0 Then txtShowDelay.Text = "0"
    If Val(txtShowDelay.Text) > 999 Then txtShowDelay.Text = "999"
End Sub
