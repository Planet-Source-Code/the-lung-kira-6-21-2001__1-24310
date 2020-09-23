VERSION 5.00
Begin VB.Form frmStartMenu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Start Menu"
   ClientHeight    =   2595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3495
   Icon            =   "frmStartMenu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2595
   ScaleWidth      =   3495
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkHelp 
      Height          =   255
      Left            =   3120
      TabIndex        =   5
      Top             =   600
      Width           =   255
   End
   Begin VB.CheckBox chkNetworkConnections 
      Height          =   255
      Left            =   3120
      TabIndex        =   9
      Top             =   1080
      Width           =   255
   End
   Begin VB.CheckBox chkRun 
      Height          =   255
      Left            =   3120
      TabIndex        =   15
      Top             =   1800
      Width           =   255
   End
   Begin VB.CheckBox chkFind 
      Height          =   255
      Left            =   3120
      TabIndex        =   3
      Top             =   360
      Width           =   255
   End
   Begin VB.CheckBox chkRecentDocsMenu 
      Height          =   255
      Left            =   3120
      TabIndex        =   13
      Top             =   1560
      Width           =   255
   End
   Begin VB.CheckBox chkRecentDocsHistory 
      Height          =   255
      Left            =   3120
      TabIndex        =   11
      Top             =   1320
      Width           =   255
   End
   Begin VB.CheckBox chkLogoff 
      Height          =   255
      Left            =   3120
      TabIndex        =   7
      Top             =   840
      Width           =   255
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   350
      Left            =   2400
      TabIndex        =   16
      Top             =   2160
      Width           =   975
   End
   Begin VB.CheckBox chkFavoritesMenu 
      Height          =   255
      Left            =   3120
      TabIndex        =   1
      Top             =   120
      Width           =   255
   End
   Begin VB.Label lblHelp 
      Caption         =   "Help"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   1965
   End
   Begin VB.Label lblNetworkConnections 
      Caption         =   "Network Connections"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1080
      Width           =   1965
   End
   Begin VB.Label lblRun 
      Caption         =   "Run"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   1800
      Width           =   1965
   End
   Begin VB.Label lblFind 
      Caption         =   "Find"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   1965
   End
   Begin VB.Label lblRecentDocsMenu 
      Caption         =   "Recent Docs Menu"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1560
      Width           =   1965
   End
   Begin VB.Label lblRecentDocsHistory 
      Caption         =   "Recent Docs History"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1320
      Width           =   1965
   End
   Begin VB.Label lblLogoff 
      Caption         =   "Logoff"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   1965
   End
   Begin VB.Label lblFavoritesMenu 
      Caption         =   "Favorites Menu"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1965
   End
End
Attribute VB_Name = "frmStartMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdApply_Click()
    If chkFavoritesMenu.value = 0 Then
        SaveRegSetting HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFavoritesMenu", 1, REG_BINARY
    Else
        SaveRegSetting HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFavoritesMenu", 0, REG_BINARY
    End If
    If chkFind.value = 0 Then
        SaveRegSetting HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFind", 1, REG_BINARY
    Else
        SaveRegSetting HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFind", 0, REG_BINARY
    End If
    If chkHelp.value = 0 Then
        SaveRegSetting HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoSMHelp", 1, REG_BINARY
    Else
        SaveRegSetting HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoSMHelp", 0, REG_BINARY
    End If
    If chkLogoff.value = 0 Then
        SaveRegSetting HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoLogoff", 1, REG_BINARY
    Else
        SaveRegSetting HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoLogoff", 0, REG_BINARY
    End If
    If chkNetworkConnections.value = 0 Then
        SaveRegSetting HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoNetworkConnections", 1, REG_BINARY
    Else
        SaveRegSetting HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoNetworkConnections", 0, REG_BINARY
    End If
    If chkRecentDocsHistory.value = 0 Then
        SaveRegSetting HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoRecentDocsHistory", 1, REG_BINARY
    Else
        SaveRegSetting HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoRecentDocsHistory", 0, REG_BINARY
    End If
    If chkRecentDocsMenu.value = 0 Then
        SaveRegSetting HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoRecentDocsMenu", 1, REG_BINARY
    Else
        SaveRegSetting HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoRecentDocsMenu", 0, REG_BINARY
    End If
    If chkRun.value = 0 Then
        SaveRegSetting HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoRun", 1, REG_BINARY
    Else
        SaveRegSetting HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoRun", 0, REG_BINARY
    End If
End Sub

Private Sub Form_Load()
    Select Case GetRegSetting(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFavoritesMenu")
        Case 0: chkFavoritesMenu.value = 1
        Case 1: chkFavoritesMenu.value = 0
    End Select
    Select Case GetRegSetting(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFind")
        Case 0: chkFind.value = 1
        Case 1: chkFind.value = 0
    End Select
    Select Case GetRegSetting(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoSMHelp")
        Case 0: chkHelp.value = 1
        Case 1: chkHelp.value = 0
    End Select
    Select Case GetRegSetting(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoLogoff")
        Case 0: chkLogoff.value = 1
        Case 1: chkLogoff.value = 0
    End Select
    Select Case GetRegSetting(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoNetworkConnections")
        Case 0: chkNetworkConnections.value = 1
        Case 1: chkNetworkConnections.value = 0
    End Select
    Select Case GetRegSetting(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoRecentDocsHistory")
        Case 0: chkRecentDocsHistory.value = 1
        Case 1: chkRecentDocsHistory.value = 0
    End Select
    Select Case GetRegSetting(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoRecentDocsMenu")
        Case 0: chkRecentDocsMenu.value = 1
        Case 1: chkRecentDocsMenu.value = 0
    End Select
    Select Case GetRegSetting(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoRun")
        Case 0: chkRun.value = 1
        Case 1: chkRun.value = 0
    End Select
End Sub
