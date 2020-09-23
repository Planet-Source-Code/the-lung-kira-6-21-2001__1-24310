VERSION 5.00
Begin VB.Form frmDisplayMonitors 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Display Monitors"
   ClientHeight    =   2655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4815
   Icon            =   "frmDisplayMonitors.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   4815
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtWorkBottom 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   2280
      Width           =   975
   End
   Begin VB.TextBox txtWorkRight 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   2040
      Width           =   975
   End
   Begin VB.TextBox txtWorkTop 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   1800
      Width           =   975
   End
   Begin VB.TextBox txtWorkLeft 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   1560
      Width           =   975
   End
   Begin VB.TextBox txtMonitorBottom 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   2280
      Width           =   975
   End
   Begin VB.TextBox txtMonitorRight 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   2040
      Width           =   975
   End
   Begin VB.TextBox txtMonitorTop 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   1800
      Width           =   975
   End
   Begin VB.TextBox txtMonitorLeft 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   1560
      Width           =   975
   End
   Begin VB.CheckBox chkPrimaryDisplay 
      Enabled         =   0   'False
      Height          =   255
      Left            =   4440
      TabIndex        =   5
      Top             =   720
      Width           =   255
   End
   Begin VB.TextBox txtDeviceName 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   480
      Width           =   2655
   End
   Begin VB.ComboBox cboDisplayMonitors 
      Height          =   315
      Left            =   2040
      TabIndex        =   1
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label lblWork 
      Caption         =   "Work"
      Height          =   255
      Left            =   2520
      TabIndex        =   15
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label lblWorkBottom 
      Caption         =   "Bottom"
      Height          =   255
      Left            =   2520
      TabIndex        =   22
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label lblWorkRight 
      Caption         =   "Right"
      Height          =   255
      Left            =   2520
      TabIndex        =   20
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label lblWorkTop 
      Caption         =   "Top"
      Height          =   255
      Left            =   2520
      TabIndex        =   18
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label lblWorkLeft 
      Caption         =   "Left"
      Height          =   255
      Left            =   2520
      TabIndex        =   16
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label lblMonitor 
      Caption         =   "Monitor"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label lblMonitorBottom 
      Caption         =   "Bottom"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label lblMonitorRight 
      Caption         =   "Right"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label lblMonitorTop 
      Caption         =   "Top"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label lblMonitorLeft 
      Caption         =   "Left"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label lblPrimaryDisplay 
      Caption         =   "Primary Display"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label lblDeviceName 
      Caption         =   "Device Name"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label lblDisplayMonitors 
      Caption         =   "Display Monitors"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmDisplayMonitors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboDisplayMonitors_Click()
    Dim MONITORINFOEX As MONITORINFOEX
    MONITORINFOEX.cbSize = Len(MONITORINFOEX)
    
    If GetMonitorInfo(CLng(cboDisplayMonitors.List(cboDisplayMonitors.ListIndex)), MONITORINFOEX) = False Then Failed "GetMonitorInfo"
    
    With MONITORINFOEX
        txtDeviceName.Text = .szDevice
        If .dwFlags And MONITORINFOF_PRIMARY Then chkPrimaryDisplay.value = 1 Else: chkPrimaryDisplay.value = 0
        
        With .rcMonitor
            txtMonitorLeft.Text = .Left
            txtMonitorTop.Text = .Top
            txtMonitorRight.Text = .Right
            txtMonitorBottom.Text = .Bottom
        End With
        With .rcWork
            txtWorkLeft.Text = .Left
            txtWorkTop.Text = .Top
            txtWorkRight.Text = .Right
            txtWorkBottom.Text = .Bottom
        End With
    End With
End Sub

Private Sub Form_Load()
    If Function_Exist("user32.dll", "EnumDisplayMonitors") = True Then
        If EnumDisplayMonitors(&H0, ByVal &H0, AddressOf MonitorEnumProc, &H0) = False Then Failed "EnumDisplayMonitors"
    Else
        lblDisplayMonitors.Enabled = False
        cboDisplayMonitors.Enabled = False
        lblDeviceName.Enabled = False
        lblPrimaryDisplay.Enabled = False
        lblMonitor.Enabled = False
        lblMonitorLeft.Enabled = False
        lblMonitorTop.Enabled = False
        lblMonitorRight.Enabled = False
        lblMonitorBottom.Enabled = False
        lblWork.Enabled = False
        lblWorkLeft.Enabled = False
        lblWorkTop.Enabled = False
        lblWorkRight.Enabled = False
        lblWorkBottom.Enabled = False
    End If
End Sub
