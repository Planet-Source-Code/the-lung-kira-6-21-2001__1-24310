VERSION 5.00
Begin VB.Form frmGetIPHost 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Get IP / Host"
   ClientHeight    =   1290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3975
   Icon            =   "frmGetIPHost.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1290
   ScaleWidth      =   3975
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGetHost 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Get Host"
      Height          =   350
      Left            =   2880
      TabIndex        =   5
      Top             =   840
      Width           =   975
   End
   Begin VB.TextBox txtHost 
      Height          =   285
      Left            =   1200
      TabIndex        =   3
      Top             =   480
      Width           =   2655
   End
   Begin VB.CommandButton cmdGetIP 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Get IP"
      Height          =   350
      Left            =   1920
      TabIndex        =   4
      Top             =   840
      Width           =   975
   End
   Begin VB.ComboBox cboIP 
      Height          =   315
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label lblHost 
      Caption         =   "Host Name"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   975
   End
   Begin VB.Label lblIP 
      Caption         =   "IP"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmGetIPHost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdGetHost_Click()
    cmdGetHost.Enabled = False
    cmdGetIP.Enabled = False
    
    txtHost.Text = GetHostByIP(cboIP.Text)
    
    cmdGetIP.Enabled = True
    cmdGetHost.Enabled = True
End Sub

Private Sub cmdGetIP_Click()
    cmdGetIP.Enabled = False
    cmdGetHost.Enabled = False
    
    With cboIP
        .Clear
        
        Dim aryIP() As String
        Dim lngIP As Long
        lngIP = GetIPByHost(txtHost.Text, aryIP())
        
        Dim lngIncrement As Long
        For lngIncrement = 1 To lngIP
            .AddItem aryIP(lngIncrement)
        Next lngIncrement
        
        If .ListCount > 0 Then
            .ListIndex = 0
        End If
    End With
    
    cmdGetHost.Enabled = True
    cmdGetIP.Enabled = True
End Sub

Private Sub Form_Load()
    txtHost.Text = GetRegSetting(HKEY_CURRENT_USER, "Software\Kira\GetIPHost", "Host")
    
    If GetRegSetting(HKEY_CURRENT_USER, "Software\Kira\GetIPHost", "IP") <> "" Then
        cboIP.AddItem GetRegSetting(HKEY_CURRENT_USER, "Software\Kira\GetIPHost", "IP")
        cboIP.ListIndex = 0
    End If
    
    
    If WS2 = False Then
        lblIP.Enabled = False
        cboIP.Enabled = False
        lblHost.Enabled = False
        txtHost.Enabled = False
        cmdGetIP.Enabled = False
        cmdGetHost.Enabled = False
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\GetIPHost", "Host", txtHost.Text, REG_SZ
    
    If cboIP.ListIndex > 0 Then
        SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\GetIPHost", "IP", cboIP.List(cboIP.ListIndex), REG_SZ
    End If
End Sub
