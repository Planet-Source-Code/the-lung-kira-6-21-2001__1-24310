VERSION 5.00
Begin VB.Form frmAdaptersInfo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Adapters Info"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8175
   Icon            =   "frmAdaptersInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   8175
   Begin VB.CheckBox chkDHCPEnabled 
      Enabled         =   0   'False
      Height          =   255
      Left            =   3720
      TabIndex        =   7
      Top             =   960
      Width           =   255
   End
   Begin VB.TextBox txtAddress 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   720
      Width           =   1815
   End
   Begin VB.CheckBox chkWINS 
      Enabled         =   0   'False
      Height          =   255
      Left            =   3720
      TabIndex        =   17
      Top             =   2160
      Width           =   255
   End
   Begin VB.TextBox txtType 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   1920
      Width           =   1815
   End
   Begin VB.TextBox txtLeaseObtained 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   1680
      Width           =   1815
   End
   Begin VB.TextBox txtLeaseExpires 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   1440
      Width           =   1815
   End
   Begin VB.TextBox txtIndex 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   1200
      Width           =   1815
   End
   Begin VB.TextBox txtSecondaryWINSServerIPMask 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   42
      Top             =   3120
      Width           =   1815
   End
   Begin VB.TextBox txtPrimaryWINSServerIPMask 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   37
      Top             =   2280
      Width           =   1815
   End
   Begin VB.TextBox txtGatewayListIPMask 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   32
      Top             =   1440
      Width           =   1815
   End
   Begin VB.TextBox txtDHCPServerIPMask 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   27
      Top             =   600
      Width           =   1815
   End
   Begin VB.TextBox txtAdapterName 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   480
      Width           =   1815
   End
   Begin VB.TextBox txtSecondaryWINSServerIPAddress 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   40
      Top             =   2880
      Width           =   1815
   End
   Begin VB.TextBox txtPrimaryWINSServerIPAddress 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   35
      Top             =   2040
      Width           =   1815
   End
   Begin VB.TextBox txtDHCPServerIPAddress 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   25
      Top             =   360
      Width           =   1815
   End
   Begin VB.TextBox txtGatewayListIPAddress 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   30
      Top             =   1200
      Width           =   1815
   End
   Begin VB.ComboBox cboIPAddressListIPMask 
      Height          =   315
      Left            =   2160
      TabIndex        =   22
      Top             =   3240
      Width           =   1815
   End
   Begin VB.ComboBox cboIPAddressListIPAddress 
      Height          =   315
      Left            =   2160
      TabIndex        =   20
      Top             =   2880
      Width           =   1815
   End
   Begin VB.ComboBox cboAdapters 
      Height          =   315
      Left            =   2160
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label lblAddress 
      Caption         =   "Address"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label lblLeaseExpires 
      Caption         =   "Lease Expires"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label lblLeaseObtained 
      Caption         =   "Lease Obtained"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label lblSecondaryWINSServer 
      Caption         =   "Secondary WINS Server"
      Height          =   255
      Left            =   4200
      TabIndex        =   38
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label lblSecondaryWINSServerIPAddress 
      Caption         =   "IP Address"
      Height          =   255
      Left            =   4200
      TabIndex        =   39
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Label lblSecondaryWINSServerIPMask 
      Caption         =   "IP Mask"
      Height          =   255
      Left            =   4200
      TabIndex        =   41
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Label lblPrimaryWINSServer 
      Caption         =   "Primary WINS Server"
      Height          =   255
      Left            =   4200
      TabIndex        =   33
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label lblPrimaryWINSServerIPAddress 
      Caption         =   "IP Address"
      Height          =   255
      Left            =   4200
      TabIndex        =   34
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label lblPrimaryWINSServerIPMask 
      Caption         =   "IP Mask"
      Height          =   255
      Left            =   4200
      TabIndex        =   36
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label lblWINS 
      Caption         =   "WINS"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label lblDHCPServer 
      Caption         =   "DHCP Server"
      Height          =   255
      Left            =   4200
      TabIndex        =   23
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label lblDHCPServerIPAddress 
      Caption         =   "IP Address"
      Height          =   255
      Left            =   4200
      TabIndex        =   24
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label lblDHCPServerIPMask 
      Caption         =   "IP Mask"
      Height          =   255
      Left            =   4200
      TabIndex        =   26
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label lblGatewayListIPMask 
      Caption         =   "IP Mask"
      Height          =   255
      Left            =   4200
      TabIndex        =   31
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label lblGatewayListIPAddress 
      Caption         =   "IP Address"
      Height          =   255
      Left            =   4200
      TabIndex        =   29
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label lblGatewayList 
      Caption         =   "Gateway List"
      Height          =   255
      Left            =   4200
      TabIndex        =   28
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label lblIPAddressListIPMask 
      Caption         =   "IP Mask"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Label lblIPAddressList 
      Caption         =   "IP Address List"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Label lblIPAddressListIPAddress 
      Caption         =   "IP Address"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Label lblDHCPEnabled 
      Caption         =   "DHCP Enabled"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label lblType 
      Caption         =   "Type"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label lblIndex 
      Caption         =   "Index"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label lblAdapterName 
      Caption         =   "Adapter Name"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label lblAdapters 
      Caption         =   "Adapters"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "frmAdaptersInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim IP_ADAPTER_INFO() As IP_ADAPTER_INFO
Dim lngCount As Long

Private Sub cboAdapters_Click()
    With IP_ADAPTER_INFO(cboAdapters.ListIndex)
        txtAdapterName.Text = .AdapterName
        
        If Len(.AdapterName) >= .AddressLength Then
            Dim strAddress As String
            Dim lngIncrement As Long
            
            strAddress = Left$(.Address, .AddressLength)
            txtAddress.Text = ""
            
            For lngIncrement = 1 To .AddressLength
                txtAddress.Text = txtAddress.Text & ltoa_(Asc(Mid$(.Address, lngIncrement, 1)), 16)
            Next lngIncrement
        End If
        
        txtIndex.Text = CStr(.Index)
        
        Select Case .Type
            Case MIB_IF_TYPE_OTHER: txtType.Text = "Other"
            Case MIB_IF_TYPE_ETHERNET: txtType.Text = "Ethernet"
            Case MIB_IF_TYPE_TOKENRING: txtType.Text = "Tokenring"
            Case MIB_IF_TYPE_FDDI: txtType.Text = "FDDI"
            Case MIB_IF_TYPE_PPP: txtType.Text = "PPP"
            Case MIB_IF_TYPE_LOOPBACK: txtType.Text = "Loopback"
            Case MIB_IF_TYPE_SLIP: txtType.Text = "Slip"
            Case Else: txtType.Text = "Unknown"
        End Select
        
        chkDHCPEnabled.value = .DhcpEnabled
        
        cboIPAddressListIPAddress.Clear
        cboIPAddressListIPAddress.AddItem .IpAddressList.IpAddress.String
        cboIPAddressListIPAddress.ListIndex = 0
        cboIPAddressListIPMask.Clear
        cboIPAddressListIPMask.AddItem .IpAddressList.IpMask.String
        cboIPAddressListIPMask.ListIndex = 0
        
        txtGatewayListIPAddress.Text = .GatewayList.IpAddress.String
        txtGatewayListIPMask.Text = .GatewayList.IpMask.String
        txtDHCPServerIPAddress.Text = .DhcpServer.IpAddress.String
        txtDHCPServerIPMask.Text = .DhcpServer.IpMask.String
        chkWINS.value = .HaveWins
        txtPrimaryWINSServerIPAddress.Text = .PrimaryWinsServer.IpAddress.String
        txtPrimaryWINSServerIPMask.Text = .PrimaryWinsServer.IpMask.String
        txtSecondaryWINSServerIPAddress.Text = .SecondaryWinsServer.IpAddress.String
        txtSecondaryWINSServerIPMask.Text = .SecondaryWinsServer.IpMask.String
        txtLeaseObtained.Text = .LeaseObtained
        txtLeaseExpires.Text = .LeaseExpires
    End With
End Sub

Private Sub Form_Load()
    If Function_Exist("iphlpapi.dll", "GetAdaptersInfo") = True Then
        Dim lngBuffer As Long
        ReDim IP_ADAPTER_INFO(0)
        
        lngBuffer = Len(IP_ADAPTER_INFO(0))
        If GetAdaptersInfo(IP_ADAPTER_INFO(0), lngBuffer) <> ERROR_SUCCESS Then
            Failed "GetAdaptersInfo"
            
            If Len(IP_ADAPTER_INFO(0)) <> lngBuffer Then
                lngCount = lngBuffer / (Len(IP_ADAPTER_INFO(0)) + 3) - 1
                ReDim IP_ADAPTER_INFO(lngCount)
                
                If GetAdaptersInfo(IP_ADAPTER_INFO(0), lngBuffer) <> ERROR_SUCCESS Then Failed "GetAdaptersInfo"
            Else
                Exit Sub
            End If
        End If
        
        Dim lngIncrement As Long
        For lngIncrement = 0 To lngCount
            If IP_ADAPTER_INFO(lngIncrement).Next <> 0 Then
                CopyMemory IP_ADAPTER_INFO(lngIncrement + 1), ByVal IP_ADAPTER_INFO(lngIncrement).Next, Len(IP_ADAPTER_INFO(lngIncrement))
            End If
            
            cboAdapters.AddItem Trim$(Fix_NullTermStr(IP_ADAPTER_INFO(lngIncrement).Description))
        Next lngIncrement
    Else
        lblAdapters.Enabled = False
        cboAdapters.Enabled = False
        lblAdapterName.Enabled = False
        lblAddress.Enabled = False
        lblDHCPEnabled.Enabled = False
        lblIndex.Enabled = False
        lblLeaseExpires.Enabled = False
        lblLeaseObtained.Enabled = False
        lblType.Enabled = False
        lblWINS.Enabled = False
        lblIPAddressList.Enabled = False
        lblIPAddressListIPAddress.Enabled = False
        cboIPAddressListIPAddress.Enabled = False
        lblIPAddressListIPMask.Enabled = False
        cboIPAddressListIPMask.Enabled = False
        lblDHCPServer.Enabled = False
        lblDHCPServerIPAddress.Enabled = False
        lblDHCPServerIPMask.Enabled = False
        lblGatewayList.Enabled = False
        lblGatewayListIPAddress.Enabled = False
        lblGatewayListIPMask.Enabled = False
        lblPrimaryWINSServer.Enabled = False
        lblPrimaryWINSServerIPAddress.Enabled = False
        lblPrimaryWINSServerIPMask.Enabled = False
        lblSecondaryWINSServer.Enabled = False
        lblSecondaryWINSServerIPAddress.Enabled = False
        lblSecondaryWINSServerIPMask.Enabled = False
    End If
End Sub
