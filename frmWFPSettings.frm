VERSION 5.00
Begin VB.Form frmWFPSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Windows File Protection - Settings"
   ClientHeight    =   2010
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4095
   Icon            =   "frmWFPSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2010
   ScaleWidth      =   4095
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cboDisable 
      Height          =   315
      Left            =   1800
      TabIndex        =   1
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   350
      Left            =   3000
      TabIndex        =   8
      Top             =   1560
      Width           =   975
   End
   Begin VB.ComboBox cboScan 
      Height          =   315
      Left            =   1800
      TabIndex        =   5
      Top             =   840
      Width           =   2175
   End
   Begin VB.TextBox txtQuota 
      Height          =   285
      Left            =   1800
      TabIndex        =   3
      Top             =   480
      Width           =   2175
   End
   Begin VB.CheckBox chkShowProgress 
      Height          =   255
      Left            =   3720
      TabIndex        =   7
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label lblDisable 
      Caption         =   "Disable"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label lblScan 
      Caption         =   "Scan"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label lblQuota 
      Caption         =   "Quota"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label lblShowProgress 
      Caption         =   "Show Progress"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   1335
   End
End
Attribute VB_Name = "frmWFPSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdApply_Click()
    SaveRegSetting HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows NT\CurrentVersion\Winlogon", "SFCDisable", cboDisable.ListIndex, REG_DWORD
    SaveRegSetting HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows NT\CurrentVersion\Winlogon", "SFCQuota", txtQuota.Text, REG_DWORD
    SaveRegSetting HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows NT\CurrentVersion\Winlogon", "SFCScan", cboScan.ListIndex, REG_DWORD
    SaveRegSetting HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows NT\CurrentVersion\Winlogon", "SFCShowProgress", chkShowProgress.value, REG_DWORD
End Sub

Private Sub Form_Load()
    If WinVersion(-1, 5000000, True) = True Then
        With cboDisable
            .AddItem "Enabled"
            .AddItem "Disabled"
            .AddItem "Disable at Next Boot"
        End With
        With cboScan
            .AddItem "No Scan at Boot"
            .AddItem "Scan at Boot"
            .AddItem "Scan Files Once"
        End With
        
        cboDisable.ListIndex = GetRegSetting(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows NT\CurrentVersion\Winlogon", "SFCDisable")
        txtQuota.Text = GetRegSetting(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows NT\CurrentVersion\Winlogon", "SFCQuota")
        cboScan.ListIndex = GetRegSetting(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows NT\CurrentVersion\Winlogon", "SFCScan")
        chkShowProgress.value = GetRegSetting(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows NT\CurrentVersion\Winlogon", "SFCShowProgress")
        
        If GetRegSetting(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows NT\CurrentVersion\Winlogon", "SFCQuota") = "-1" Then txtQuota.Text = "4294967295"
    Else
        lblDisable.Enabled = False
        cboDisable.Enabled = False
        lblQuota.Enabled = False
        txtQuota.Enabled = False
        lblScan.Enabled = False
        cboScan.Enabled = False
        lblShowProgress.Enabled = False
        chkShowProgress.Enabled = False
        cmdApply.Enabled = False
    End If
End Sub
Private Sub txtQuota_Change()
    txtQuota.Text = CStr(Val(Rem_NonNumeric_Chr(txtQuota.Text)))
    If Val(txtQuota.Text) < 0 Then txtQuota.Text = "0"
    If Val(txtQuota.Text) > 4294967295# Then txtQuota.Text = "4294967295"
End Sub

