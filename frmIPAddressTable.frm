VERSION 5.00
Begin VB.Form frmIPAddressTable 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "IP Address Table"
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4215
   Icon            =   "frmIPAddressTable.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   4215
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtMaxReasmSize 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   2160
      Width           =   1455
   End
   Begin VB.TextBox txtBroadcastAddress 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   1920
      Width           =   1455
   End
   Begin VB.TextBox txtSubnetMask 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1680
      Width           =   1455
   End
   Begin VB.TextBox txtInterfaceIndex 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   1440
      Width           =   1455
   End
   Begin VB.TextBox txtIPAddress 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1200
      Width           =   1455
   End
   Begin VB.ListBox lstIPAddress_Table 
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   2640
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton cmdRefresh 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Refresh"
      Height          =   350
      Left            =   3120
      TabIndex        =   14
      Top             =   2880
      Width           =   975
   End
   Begin VB.CheckBox chkSorted 
      Height          =   255
      Left            =   3840
      TabIndex        =   13
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label lblSorted 
      Caption         =   "Sorted"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   2520
      Width           =   2295
   End
   Begin VB.Label lblEntry 
      Caption         =   "Entry"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label lblMaxReasmSize 
      Caption         =   "Max Datagram Reassembly Size"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2160
      Width           =   2295
   End
   Begin VB.Label lblBroadcastAddress 
      Caption         =   "Broadcast Address"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1920
      Width           =   2295
   End
   Begin VB.Label lblSubnetMask 
      Caption         =   "Subnet Mask"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   2295
   End
   Begin VB.Label lblInterfaceIndex 
      Caption         =   "Interface Index"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   2295
   End
   Begin VB.Label lblIPAddress 
      Caption         =   "IP Address"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   2295
   End
End
Attribute VB_Name = "frmIPAddressTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim MIB_IPADDRTABLE As MIB_IPADDRTABLE

Private Sub cmdRefresh_Click()
    Dim lngSize As Long
    lngSize = Len(MIB_IPADDRTABLE)
    
    If GetIpAddrTable(MIB_IPADDRTABLE, lngSize, chkSorted.value) <> NO_ERROR Then Failed "GetIpAddrTable"
    
    
    With lstIPAddress_Table
        .Clear
        
        Dim lngIncrement As Long
        For lngIncrement = 0 To MIB_IPADDRTABLE.dwNumEntries - 1
            .AddItem CStr(lngIncrement + 1)
        Next lngIncrement
    End With
    
    txtIPAddress.Text = ""
    txtInterfaceIndex.Text = ""
    txtSubnetMask.Text = ""
    txtBroadcastAddress.Text = ""
    txtMaxReasmSize.Text = ""
End Sub

Private Sub Form_Load()
    If Function_Exist("iphlpapi.dll", "GetIpAddrTable") = True Then
        cmdRefresh_Click
    Else
        lblEntry.Enabled = False
        lblIPAddress.Enabled = False
        lblInterfaceIndex.Enabled = False
        lblSubnetMask.Enabled = False
        lblBroadcastAddress.Enabled = False
        lblMaxReasmSize.Enabled = False
        lblSorted.Enabled = False
        chkSorted.Enabled = False
        cmdRefresh.Enabled = False
    End If

    chkSorted.value = GetRegSetting(HKEY_CURRENT_USER, "Software\Kira\IPAddressTable", "Sorted")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\IPAddressTable", "Sorted", chkSorted.value, REG_DWORD
End Sub

Private Sub lstIPAddress_Table_Click()
    With MIB_IPADDRTABLE.table(lstIPAddress_Table.ListIndex)
        txtIPAddress.Text = CStr(.dwAddr(0)) & "." & CStr(.dwAddr(1)) & "." & CStr(.dwAddr(2)) & "." & CStr(.dwAddr(3))
        txtInterfaceIndex.Text = CStr(.dwIndex)
        txtSubnetMask.Text = CStr(.dwMask(0)) & "." & CStr(.dwMask(1)) & "." & CStr(.dwMask(2)) & "." & CStr(.dwMask(3))
        txtBroadcastAddress.Text = CStr(.dwBCastAddr(0)) & "." & CStr(.dwBCastAddr(1)) & "." & CStr(.dwBCastAddr(2)) & "." & CStr(.dwBCastAddr(3))
        txtMaxReasmSize.Text = CStr(.dwReasmSize)
    End With
End Sub
