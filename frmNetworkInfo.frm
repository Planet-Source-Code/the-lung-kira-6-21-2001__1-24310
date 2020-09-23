VERSION 5.00
Begin VB.Form frmNetworkInfo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Network Info"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4215
   Icon            =   "frmNetworkInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   4215
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtLocalHostName 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   840
      Width           =   1935
   End
   Begin VB.ComboBox cboLocalIP 
      Height          =   315
      Left            =   2160
      TabIndex        =   3
      Top             =   480
      Width           =   1935
   End
   Begin VB.CheckBox chkNetworkPresent 
      Enabled         =   0   'False
      Height          =   255
      Left            =   3840
      TabIndex        =   1
      Top             =   120
      Width           =   255
   End
   Begin VB.CheckBox chkInetIsOffline 
      Enabled         =   0   'False
      Height          =   255
      Left            =   3840
      TabIndex        =   27
      Top             =   4200
      Width           =   255
   End
   Begin VB.TextBox txtNumberOfInterfaces 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   25
      Top             =   3960
      Width           =   1935
   End
   Begin VB.CheckBox chkEnableDns 
      Enabled         =   0   'False
      Height          =   255
      Left            =   3840
      TabIndex        =   21
      Top             =   3360
      Width           =   255
   End
   Begin VB.CheckBox chkEnableRouting 
      Enabled         =   0   'False
      Height          =   255
      Left            =   3840
      TabIndex        =   23
      Top             =   3600
      Width           =   255
   End
   Begin VB.CheckBox chkEnableProxy 
      Enabled         =   0   'False
      Height          =   255
      Left            =   3840
      TabIndex        =   19
      Top             =   3120
      Width           =   255
   End
   Begin VB.TextBox txtScopeId 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   2760
      Width           =   1935
   End
   Begin VB.CheckBox chkNodeType 
      Enabled         =   0   'False
      Height          =   255
      Left            =   3840
      TabIndex        =   15
      Top             =   2520
      Width           =   255
   End
   Begin VB.ComboBox cboDNSServerList 
      Height          =   315
      Left            =   2160
      TabIndex        =   13
      Top             =   2160
      Width           =   1935
   End
   Begin VB.TextBox txtCurrentDNSServer 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   1800
      Width           =   1935
   End
   Begin VB.TextBox txtDomainName 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1200
      Width           =   1935
   End
   Begin VB.TextBox txtHostName 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label lblLocalHostName 
      Caption         =   "Local Host Name"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label lblLocalIP 
      Caption         =   "Local IP"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label lblNetworkPresent 
      Caption         =   "Network Present"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label lblInetIsOffline 
      Caption         =   "Inet Is Offline"
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   4200
      Width           =   1815
   End
   Begin VB.Label lblNumberOfInterfaces 
      Caption         =   "Number Of Interfaces"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   3960
      Width           =   1815
   End
   Begin VB.Label lblEnableDns 
      Caption         =   "DNS Enabled"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   3360
      Width           =   1815
   End
   Begin VB.Label lblEnableRouting 
      Caption         =   "Routing Enabled"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Label lblEnableProxy 
      Caption         =   "ARP Proxy"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Label lblScopeId 
      Caption         =   "DHCP Scope Name"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Label lblNodeType 
      Caption         =   "Use DHCP"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Label lblDNSServerList 
      Caption         =   "DNS Server List"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label lblCurrentDNSServer 
      Caption         =   "Current DNS Server"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label lblDomainName 
      Caption         =   "Domain Name"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label lblHostName 
      Caption         =   "Host Name"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1440
      Width           =   1815
   End
End
Attribute VB_Name = "frmNetworkInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    chkNetworkPresent.value = CInt(Right$(Right$("00000000" & ltoa_(Asc(GetSystemMetrics(SM_NETWORK)), 2), 8), 1))
    
    
    Dim aryIP() As String
    Dim lngIP As Long
    lngIP = GetIPByHost(Get_ComputerName, aryIP())
    If lngIP > 0 Then txtLocalHostName.Text = GetHostByIP(aryIP(lngIP))
    
    With cboLocalIP
        Dim lngIncrement As Long
        For lngIncrement = 1 To lngIP
            .AddItem aryIP(lngIncrement)
        Next lngIncrement
        
        If .ListCount > 0 Then
            .ListIndex = 0
        End If
    End With
    
    
    If Function_Exist("iphlpapi.dll", "GetNetworkParams") = True Then
        Dim FIXED_INFO As FIXED_INFO
        apiError = GetNetworkParams(FIXED_INFO, Len(FIXED_INFO)): If apiError <> ERROR_SUCCESS Then Errors apiError, "GetNetworkParams"
        
        With FIXED_INFO
            txtCurrentDNSServer.Text = .CurrentDnsServer.IpAddress.String
            cboDNSServerList.AddItem .DnsServerList.IpAddress.String
            cboDNSServerList.ListIndex = 0
            txtDomainName.Text = .DomainName
            If .EnableDns > 0 Then chkEnableDns.value = 1
            If .EnableProxy > 0 Then chkEnableProxy.value = 1
            If .EnableRouting > 0 Then chkEnableRouting.value = 1
            txtHostName.Text = .HostName
            If .NodeType > 0 Then chkNodeType.value = 1
            txtScopeId.Text = .ScopeId
        End With
    Else
        lblCurrentDNSServer.Enabled = False
        lblDNSServerList.Enabled = False
        cboDNSServerList.Enabled = False
        lblDomainName.Enabled = False
        lblEnableDns.Enabled = False
        lblEnableProxy.Enabled = False
        lblEnableRouting.Enabled = False
        lblHostName.Enabled = False
        lblNodeType.Enabled = False
        lblScopeId.Enabled = False
    End If
    
    If Function_Exist("iphlpapi.dll", "GetNumberOfInterfaces") = True Then
        Dim lngInterfaces As Long
        
        apiError = GetNumberOfInterfaces(lngInterfaces): If apiError <> 0 Then Errors apiError, "GetNumberOfInterfaces"
        txtNumberOfInterfaces.Text = CStr(lngInterfaces)
    Else
        lblNumberOfInterfaces.Enabled = False
    End If
    If Function_Exist("url.dll", "InetIsOffline") = True Then
        chkInetIsOffline.value = CInt(InetIsOffline(0))
    Else
        lblInetIsOffline.Enabled = False
    End If
End Sub
