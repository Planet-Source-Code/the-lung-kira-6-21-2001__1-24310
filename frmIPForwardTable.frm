VERSION 5.00
Begin VB.Form frmIPForwardTable 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "IP Forward Table"
   ClientHeight    =   4275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3375
   Icon            =   "frmIPForwardTable.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   3375
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkSorted 
      Height          =   255
      Left            =   3000
      TabIndex        =   21
      Top             =   3480
      Width           =   255
   End
   Begin VB.TextBox txtASNextHop 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   3120
      Width           =   1455
   End
   Begin VB.TextBox txtAge 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   2880
      Width           =   1455
   End
   Begin VB.TextBox txtProtocol 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   2640
      Width           =   1455
   End
   Begin VB.TextBox txtRouteType 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   2400
      Width           =   1455
   End
   Begin VB.TextBox txtInterfaceIndex 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   2160
      Width           =   1455
   End
   Begin VB.TextBox txtNextHop 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   1920
      Width           =   1455
   End
   Begin VB.TextBox txtPolicy 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1680
      Width           =   1455
   End
   Begin VB.TextBox txtMask 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   1440
      Width           =   1455
   End
   Begin VB.TextBox txtDestination 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1200
      Width           =   1455
   End
   Begin VB.ListBox lstIPForward_Table 
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
      Left            =   1800
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton cmdRefresh 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Refresh"
      Height          =   350
      Left            =   2280
      TabIndex        =   22
      Top             =   3840
      Width           =   975
   End
   Begin VB.Label lblSorted 
      Caption         =   "Sorted"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label lblASNextHop 
      Caption         =   "AS Next Hop"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label lblAge 
      Caption         =   "Age"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label lblProtocol 
      Caption         =   "Protocol"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label lblRouteType 
      Caption         =   "Route Type"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label lblInterfaceIndex 
      Caption         =   "Interface Index"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label lblNextHop 
      Caption         =   "Next Hop"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label lblPolicy 
      Caption         =   "Policy"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label lblMask 
      Caption         =   "Mask"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label lblDestination 
      Caption         =   "Destination"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label lblEntry 
      Caption         =   "Entry"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmIPForwardTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim MIB_IPFORWARDTABLE As MIB_IPFORWARDTABLE

Private Sub cmdRefresh_Click()
    Dim lngSize As Long
    lngSize = Len(MIB_IPFORWARDTABLE)
    
    If GetIpForwardTable(MIB_IPFORWARDTABLE, lngSize, chkSorted.value) <> NO_ERROR Then Failed "GetIpForwardTable"
    
    
    With lstIPForward_Table
        .Clear
        
        Dim lngIncrement As Long
        For lngIncrement = 0 To MIB_IPFORWARDTABLE.dwNumEntries - 1
            .AddItem CStr(lngIncrement + 1)
        Next lngIncrement
    End With
    
    txtDestination.Text = ""
    txtMask.Text = ""
    txtPolicy.Text = ""
    txtNextHop.Text = ""
    txtInterfaceIndex.Text = ""
    txtRouteType.Text = ""
    txtProtocol.Text = ""
    txtAge.Text = ""
    txtASNextHop.Text = ""
End Sub

Private Sub Form_Load()
    If Function_Exist("iphlpapi.dll", "GetIpForwardTable") = True Then
        cmdRefresh_Click
    Else
        lblEntry.Enabled = False
        lstIPForward_Table.Enabled = False
        lblDestination.Enabled = False
        lblMask.Enabled = False
        lblPolicy.Enabled = False
        lblNextHop.Enabled = False
        lblInterfaceIndex.Enabled = False
        lblRouteType.Enabled = False
        lblProtocol.Enabled = False
        lblAge.Enabled = False
        lblASNextHop.Enabled = False
        lblSorted.Enabled = False
        chkSorted.Enabled = False
        cmdRefresh.Enabled = False
    End If

    chkSorted.value = GetRegSetting(HKEY_CURRENT_USER, "Software\Kira\IPForwardTable", "Sorted")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\IPForwardTable", "Sorted", chkSorted.value, REG_DWORD
End Sub

Private Sub lstIPForward_Table_Click()
    With MIB_IPFORWARDTABLE.table(lstIPForward_Table.ListIndex)
        txtDestination.Text = CStr(.dwForwardDest(0)) & "." & CStr(.dwForwardDest(1)) & "." & CStr(.dwForwardDest(2)) & "." & CStr(.dwForwardDest(3))
        txtMask.Text = CStr(.dwForwardMask(0)) & "." & CStr(.dwForwardMask(1)) & "." & CStr(.dwForwardMask(2)) & "." & CStr(.dwForwardMask(3))
        txtPolicy.Text = CStr(.dwForwardPolicy)
        txtNextHop.Text = CStr(.dwForwardNextHop(0)) & "." & CStr(.dwForwardNextHop(1)) & "." & CStr(.dwForwardNextHop(2)) & "." & CStr(.dwForwardNextHop(3))
        txtInterfaceIndex.Text = CStr(.dwForwardIfIndex)
        
        Select Case .dwForwardType
            Case 4: txtRouteType.Text = "Not Final Destination"
            Case 3: txtRouteType.Text = "Final Destination"
            Case 2: txtRouteType.Text = "Invalid"
            Case 1: txtRouteType.Text = "Other"
            Case Else: txtRouteType.Text = "Unknown"
        End Select
        
        Select Case .dwForwardProto
            Case PROTO_IP_OTHER: txtProtocol.Text = "Other"
            Case PROTO_IP_LOCAL: txtProtocol.Text = "Local"
            Case PROTO_IP_NETMGMT: txtProtocol.Text = "NetMgmt"
            Case PROTO_IP_ICMP: txtProtocol.Text = "ICMP"
            Case PROTO_IP_EGP: txtProtocol.Text = "EGP"
            Case PROTO_IP_GGP: txtProtocol.Text = "GGP"
            Case PROTO_IP_HELLO: txtProtocol.Text = "HELLO"
            Case PROTO_IP_RIP: txtProtocol.Text = "RIP"
            Case PROTO_IP_IS_IS: txtProtocol.Text = "IS IS"
            Case PROTO_IP_ES_IS: txtProtocol.Text = "ES IS"
            Case PROTO_IP_CISCO: txtProtocol.Text = "Cisco"
            Case PROTO_IP_BBN: txtProtocol.Text = "BBN"
            Case PROTO_IP_OSPF: txtProtocol.Text = "OSPF"
            Case PROTO_IP_BGP: txtProtocol.Text = "BGP"
            Case PROTO_IP_BOOTP: txtProtocol.Text = "BootP"
            Case PROTO_IP_NT_AUTOSTATIC: txtProtocol.Text = "NT AutoStatic"
            Case PROTO_IP_NT_STATIC: txtProtocol.Text = "NT Static"
            Case PROTO_IP_NT_STATIC_NON_DOD: txtProtocol.Text = "NT Static Non DOD"
            Case Else: txtProtocol.Text = "Unknown"
        End Select

        txtAge.Text = CStr(.dwForwardAge)
        txtASNextHop.Text = CStr(.dwForwardNextHopAS)
    End With
End Sub

