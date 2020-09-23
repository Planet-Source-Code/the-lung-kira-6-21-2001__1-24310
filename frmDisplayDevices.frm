VERSION 5.00
Begin VB.Form frmDisplayDevices 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Display Devices"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4815
   Icon            =   "frmDisplayDevices.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   4815
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkVGACompatible 
      Enabled         =   0   'False
      Height          =   255
      Left            =   4440
      TabIndex        =   19
      Top             =   2520
      Width           =   255
   End
   Begin VB.CheckBox chkRemovable 
      Enabled         =   0   'False
      Height          =   255
      Left            =   4440
      TabIndex        =   17
      Top             =   2280
      Width           =   255
   End
   Begin VB.CheckBox chkPrimaryDevice 
      Enabled         =   0   'False
      Height          =   255
      Left            =   4440
      TabIndex        =   15
      Top             =   2040
      Width           =   255
   End
   Begin VB.CheckBox chkModesPruned 
      Enabled         =   0   'False
      Height          =   255
      Left            =   4440
      TabIndex        =   13
      Top             =   1800
      Width           =   255
   End
   Begin VB.CheckBox chkMirroringDriver 
      Enabled         =   0   'False
      Height          =   255
      Left            =   4440
      TabIndex        =   11
      Top             =   1560
      Width           =   255
   End
   Begin VB.CheckBox chkAttachedToDesktop 
      Enabled         =   0   'False
      Height          =   255
      Left            =   4440
      TabIndex        =   9
      Top             =   1320
      Width           =   255
   End
   Begin VB.TextBox txtDeviceName 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   960
      Width           =   2655
   End
   Begin VB.TextBox txtDeviceKey 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   720
      Width           =   2655
   End
   Begin VB.TextBox txtDeviceID 
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
   Begin VB.ComboBox cboDisplayDevices 
      Height          =   315
      Left            =   2040
      TabIndex        =   1
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label lblVGACompatible 
      Caption         =   "VGA Compatible"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label lblRemovable 
      Caption         =   "Removable"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label lblPrimaryDevice 
      Caption         =   "Primary Device"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label lblModesPruned 
      Caption         =   "Modes Pruned"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label lblMirroringDriver 
      Caption         =   "Mirroring Driver"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label lblAttachedToDesktop 
      Caption         =   "Attached To Desktop"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label lblDeviceKey 
      Caption         =   "Device Key"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label lblDeviceID 
      Caption         =   "Device ID"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label lblDeviceName 
      Caption         =   "Device Name"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label lblDisplayDevices 
      Caption         =   "Display Devices"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmDisplayDevices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboDisplayDevices_Click()
    Dim DISPLAY_DEVICE As DISPLAY_DEVICE
    DISPLAY_DEVICE.cb = Len(DISPLAY_DEVICE)
    
    If EnumDisplayDevices(0, cboDisplayDevices.ListIndex, DISPLAY_DEVICE, 0) = False Then Failed "EnumDisplayDevices"
    
    With DISPLAY_DEVICE
        txtDeviceID.Text = .DeviceID
        txtDeviceKey.Text = .DeviceKey
        txtDeviceName.Text = .DeviceName
        
        If .StateFlags And DISPLAY_DEVICE_ATTACHED_TO_DESKTOP Then chkAttachedToDesktop.value = 1 Else: chkAttachedToDesktop.value = 0
        If .StateFlags And DISPLAY_DEVICE_MIRRORING_DRIVER Then chkMirroringDriver.value = 1 Else: chkMirroringDriver.value = 0
        If .StateFlags And DISPLAY_DEVICE_MODESPRUNED Then chkModesPruned.value = 1 Else: chkModesPruned.value = 0
        If .StateFlags And DISPLAY_DEVICE_PRIMARY_DEVICE Then chkPrimaryDevice.value = 1 Else: chkPrimaryDevice.value = 0
        If .StateFlags And DISPLAY_DEVICE_REMOVABLE Then chkRemovable.value = 1 Else: chkRemovable.value = 0
        If .StateFlags And DISPLAY_DEVICE_VGA_COMPATIBLE Then chkVGACompatible.value = 1 Else: chkVGACompatible.value = 0
    End With
End Sub

Private Sub Form_Load()
    If Function_Exist("user32.dll", "EnumDisplayDevices") = True Then
        Dim DISPLAY_DEVICE As DISPLAY_DEVICE
        Dim lngIncrement As Long
        DISPLAY_DEVICE.cb = Len(DISPLAY_DEVICE)
        
        Do
            If EnumDisplayDevices(0, lngIncrement, DISPLAY_DEVICE, 0) = False Then
                Failed "EnumDisplayDevices"
                Exit Do
            End If
            
            cboDisplayDevices.AddItem DISPLAY_DEVICE.DeviceString
            lngIncrement = lngIncrement + 1
        Loop
    Else
        lblDisplayDevices.Enabled = False
        cboDisplayDevices.Enabled = False
        lblDeviceID.Enabled = False
        lblDeviceKey.Enabled = False
        lblDeviceName.Enabled = False
        lblAttachedToDesktop.Enabled = False
        lblMirroringDriver.Enabled = False
        lblModesPruned.Enabled = False
        lblPrimaryDevice.Enabled = False
        lblRemovable.Enabled = False
        lblVGACompatible.Enabled = False
    End If
End Sub
