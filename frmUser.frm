VERSION 5.00
Begin VB.Form frmUser 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User"
   ClientHeight    =   1995
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4455
   Icon            =   "frmUser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1995
   ScaleWidth      =   4455
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtComputerName 
      Height          =   285
      Left            =   1680
      TabIndex        =   1
      Top             =   120
      Width           =   2655
   End
   Begin VB.TextBox txtUserName 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1200
      Width           =   2655
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   350
      Left            =   3360
      TabIndex        =   8
      Top             =   1560
      Width           =   975
   End
   Begin VB.TextBox txtOrginization 
      Height          =   285
      Left            =   1680
      TabIndex        =   5
      Top             =   840
      Width           =   2655
   End
   Begin VB.TextBox txtOwner 
      Height          =   285
      Left            =   1680
      TabIndex        =   3
      Top             =   480
      Width           =   2655
   End
   Begin VB.Label lblComputerName 
      Caption         =   "Computer Name"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label lblUserName 
      Caption         =   "User Name"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label lblOrginization 
      Caption         =   "Orginization"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label lblOwner 
      Caption         =   "Owner"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1335
   End
End
Attribute VB_Name = "frmUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdApply_Click()
    Set_ComputerName txtComputerName.Text
    
    If WinID = VER_PLATFORM_WIN32_WINDOWS Then
        SaveRegSetting HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion", "RegisteredOwner", txtOwner.Text, REG_SZ
        SaveRegSetting HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion", "RegisteredOrganization", txtOrginization.Text, REG_SZ
    Else
        SaveRegSetting HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows NT\CurrentVersion", "RegisteredOwner", txtOwner.Text, REG_SZ
        SaveRegSetting HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows NT\CurrentVersion", "RegisteredOrganization", txtOrginization.Text, REG_SZ
    End If
End Sub

Private Sub Form_Load()
    txtComputerName.Text = Get_ComputerName
    
    If WinID = VER_PLATFORM_WIN32_WINDOWS Then
        txtOwner.Text = GetRegSetting(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion", "RegisteredOwner")
        txtOrginization.Text = GetRegSetting(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion", "RegisteredOrganization")
    Else
        txtOwner.Text = GetRegSetting(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows NT\CurrentVersion", "RegisteredOwner")
        txtOrginization.Text = GetRegSetting(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows NT\CurrentVersion", "RegisteredOrganization")
    End If
    
    txtUserName.Text = Get_UserName
End Sub

Private Sub txtComputerName_Change()
    txtComputerName.Text = Rem_NonStd_Chr(txtComputerName.Text)
End Sub
