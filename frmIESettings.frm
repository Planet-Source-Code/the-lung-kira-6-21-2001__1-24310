VERSION 5.00
Begin VB.Form frmIESettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "IE Settings"
   ClientHeight    =   2730
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6135
   Icon            =   "frmIESettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   6135
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtSearchPage 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2160
      TabIndex        =   9
      Top             =   1560
      Width           =   3855
   End
   Begin VB.TextBox txtDefaultURL 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2160
      TabIndex        =   1
      Top             =   120
      Width           =   3855
   End
   Begin VB.TextBox txtDefaultSearch 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2160
      TabIndex        =   3
      Top             =   480
      Width           =   3855
   End
   Begin VB.TextBox txtWindowTitle 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2160
      TabIndex        =   11
      Top             =   1920
      Width           =   3855
   End
   Begin VB.CheckBox chkPersistantLinksFolder 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   5760
      TabIndex        =   5
      Top             =   840
      Width           =   255
   End
   Begin VB.CommandButton cmdApply 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Apply"
      Height          =   350
      Left            =   5040
      TabIndex        =   12
      Top             =   2280
      Width           =   975
   End
   Begin VB.CheckBox chkRatings 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   5760
      TabIndex        =   7
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label lblSearchPage 
      Caption         =   "Search Page"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label lblDefaultURL 
      Caption         =   "Default URL"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label lblDefaultSearch 
      Caption         =   "Default Search"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label lblWindowTitle 
      Caption         =   "Window Title"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label lblPersistantLinksFolder 
      Caption         =   "Persistant Links Folder"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label lblRatings 
      Caption         =   "Ratings"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   1815
   End
End
Attribute VB_Name = "frmIESettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdApply_Click()
    SaveRegSetting HKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\Main", "Default_Page_URL", txtDefaultURL.Text, REG_SZ
    SaveRegSetting HKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\Main", "Default_Search_URL", txtDefaultSearch.Text, REG_SZ
    
    If chkPersistantLinksFolder.value = 0 Then
        SaveRegSetting HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Toolbar", "LinksFolderName", "", REG_SZ
    Else
        SaveRegSetting HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Toolbar", "LinksFolderName", "Links", REG_SZ
    End If
    If chkRatings.value = 1 Then
        DeleteValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\policies\Ratings", "Key"
    End If
    
    SaveRegSetting HKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\Main", "Search Page", txtSearchPage.Text, REG_SZ
    SaveRegSetting HKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\Main", "Window Title", txtWindowTitle.Text, REG_SZ
End Sub

Private Sub Form_Load()
    txtDefaultURL.Text = GetRegSetting(HKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\Main", "Default_Page_URL")
    txtDefaultSearch.Text = GetRegSetting(HKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\Main", "Default_Search_URL")
    
    If GetRegSetting(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Toolbar", "LinksFolderName") <> "" Then
        chkPersistantLinksFolder.value = 1
    End If
    If GetRegSetting(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\policies\Ratings", "Key") <> "" Then
        chkRatings.value = 1
    End If
    
    txtSearchPage.Text = GetRegSetting(HKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\Main", "Search Page")
    txtWindowTitle.Text = GetRegSetting(HKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\Main", "Window Title")
End Sub
