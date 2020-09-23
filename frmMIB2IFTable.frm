VERSION 5.00
Begin VB.Form frmMIB2IFTable 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MIB-II Interface Table"
   ClientHeight    =   4275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7215
   Icon            =   "frmMIB2IFTable.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   7215
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtSpeed 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   3600
      Width           =   1455
   End
   Begin VB.TextBox txtPhysicalAddress 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   3360
      Width           =   1455
   End
   Begin VB.TextBox txtOutputQueueLength 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   3120
      Width           =   1455
   End
   Begin VB.TextBox txtOperationalStatus 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   2880
      Width           =   1455
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   2640
      Width           =   1455
   End
   Begin VB.TextBox txtMTU 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   2400
      Width           =   1455
   End
   Begin VB.TextBox txtLastStatusChange 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   2160
      Width           =   1455
   End
   Begin VB.TextBox txtInterfaceType 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   1920
      Width           =   1455
   End
   Begin VB.TextBox txtInterfaceIndex 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1680
      Width           =   1455
   End
   Begin VB.TextBox txtDescription 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   1440
      Width           =   1455
   End
   Begin VB.TextBox txtErroneousOut 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5640
      Locked          =   -1  'True
      TabIndex        =   47
      Top             =   3120
      Width           =   1455
   End
   Begin VB.TextBox txtDiscardedOut 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5640
      Locked          =   -1  'True
      TabIndex        =   45
      Top             =   2880
      Width           =   1455
   End
   Begin VB.TextBox txtNonUnicastOut 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5640
      Locked          =   -1  'True
      TabIndex        =   43
      Top             =   2640
      Width           =   1455
   End
   Begin VB.TextBox txtUnicastOut 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5640
      Locked          =   -1  'True
      TabIndex        =   41
      Top             =   2400
      Width           =   1455
   End
   Begin VB.TextBox txtOctetsOut 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5640
      Locked          =   -1  'True
      TabIndex        =   39
      Top             =   2160
      Width           =   1455
   End
   Begin VB.TextBox txtUnknownProtocolIn 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5640
      Locked          =   -1  'True
      TabIndex        =   36
      Top             =   1560
      Width           =   1455
   End
   Begin VB.TextBox txtErroneousIn 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5640
      Locked          =   -1  'True
      TabIndex        =   34
      Top             =   1320
      Width           =   1455
   End
   Begin VB.TextBox txtDiscardedIn 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5640
      Locked          =   -1  'True
      TabIndex        =   32
      Top             =   1080
      Width           =   1455
   End
   Begin VB.TextBox txtNonUnicastIn 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5640
      Locked          =   -1  'True
      TabIndex        =   30
      Top             =   840
      Width           =   1455
   End
   Begin VB.TextBox txtUnicastIn 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5640
      Locked          =   -1  'True
      TabIndex        =   28
      Top             =   600
      Width           =   1455
   End
   Begin VB.TextBox txtOctetsIn 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5640
      Locked          =   -1  'True
      TabIndex        =   26
      Top             =   360
      Width           =   1455
   End
   Begin VB.TextBox txtAdminStatus 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CheckBox chkSorted 
      Height          =   255
      Left            =   6840
      TabIndex        =   49
      Top             =   3480
      Width           =   255
   End
   Begin VB.CommandButton cmdRefresh 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Refresh"
      Height          =   350
      Left            =   6120
      TabIndex        =   50
      Top             =   3840
      Width           =   975
   End
   Begin VB.ListBox lstMIB2IF_Table 
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
      Left            =   2040
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label lblDescription 
      Caption         =   "Description"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label lblOutputQueueLength 
      Caption         =   "Output Queue Length"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   3120
      Width           =   1695
   End
   Begin VB.Label lblPacketsOut 
      Caption         =   "Packets Out"
      Height          =   255
      Left            =   3720
      TabIndex        =   37
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label lblErroneousOut 
      Caption         =   "Erroneous"
      Height          =   255
      Left            =   3720
      TabIndex        =   46
      Top             =   3120
      Width           =   1695
   End
   Begin VB.Label lblDiscardedOut 
      Caption         =   "Discarded"
      Height          =   255
      Left            =   3720
      TabIndex        =   44
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Label lblNonUnicastOut 
      Caption         =   "Non Unicast"
      Height          =   255
      Left            =   3720
      TabIndex        =   42
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label lblUnicastOut 
      Caption         =   "Unicast"
      Height          =   255
      Left            =   3720
      TabIndex        =   40
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Label lblOctetsOut 
      Caption         =   "Octets"
      Height          =   255
      Left            =   3720
      TabIndex        =   38
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label lblPacketsIn 
      Caption         =   "Packets In"
      Height          =   255
      Left            =   3720
      TabIndex        =   24
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label lblUnknownProtocolIn 
      Caption         =   "Unknown Protocol"
      Height          =   255
      Left            =   3720
      TabIndex        =   35
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label lblErroneousIn 
      Caption         =   "Erroneous"
      Height          =   255
      Left            =   3720
      TabIndex        =   33
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label lblDiscardedIn 
      Caption         =   "Discarded"
      Height          =   255
      Left            =   3720
      TabIndex        =   31
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label lblNonUnicastIn 
      Caption         =   "Non Unicast"
      Height          =   255
      Left            =   3720
      TabIndex        =   29
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label lblUnicastIn 
      Caption         =   "Unicast"
      Height          =   255
      Left            =   3720
      TabIndex        =   27
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label lblOctetsIn 
      Caption         =   "Octets"
      Height          =   255
      Left            =   3720
      TabIndex        =   25
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label lblLastStatusChange 
      Caption         =   "Last Status Change"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label lblOperationalStatus 
      Caption         =   "Operational Status"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Label lblPhysicalAddress 
      Caption         =   "Physical Address"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Label lblAdminStatus 
      Caption         =   "Admin Status"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label lblSpeed 
      Caption         =   "Speed"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   3600
      Width           =   1695
   End
   Begin VB.Label lblSorted 
      Caption         =   "Sorted"
      Height          =   255
      Left            =   3720
      TabIndex        =   48
      Top             =   3480
      Width           =   1695
   End
   Begin VB.Label lblMTU 
      Caption         =   "MTU"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Label lblInterfaceType 
      Caption         =   "Interface Type"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label lblInterfaceIndex 
      Caption         =   "Interface Index"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label lblName 
      Caption         =   "Name"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label lblEntry 
      Caption         =   "Entry"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmMIB2IFTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim MIB_IFTABLE As MIB_IFTABLE

Private Sub cmdRefresh_Click()
    Dim lngSize As Long
    
    lngSize = Len(MIB_IFTABLE)
    If GetIfTable(MIB_IFTABLE, lngSize, chkSorted.value) <> NO_ERROR Then Failed "GetIfTable"
    
    
    With lstMIB2IF_Table
        .Clear
        
        Dim lngIncrement As Long
        For lngIncrement = 0 To MIB_IFTABLE.dwNumEntries - 1
            .AddItem CStr(lngIncrement + 1)
        Next lngIncrement
    End With
    
    txtAdminStatus.Text = ""
    txtDescription.Text = ""
    txtDiscardedIn.Text = ""
    txtDiscardedOut.Text = ""
    txtErroneousIn.Text = ""
    txtErroneousOut.Text = ""
    txtName.Text = ""
    txtInterfaceIndex.Text = ""
    txtInterfaceType.Text = ""
    txtLastStatusChange.Text = ""
    txtMTU.Text = ""
    txtNonUnicastIn.Text = ""
    txtNonUnicastOut.Text = ""
    txtOctetsIn.Text = ""
    txtOctetsOut.Text = ""
    txtOutputQueueLength.Text = ""
    txtOperationalStatus.Text = ""
    txtPhysicalAddress.Text = ""
    txtSpeed.Text = ""
    txtUnicastIn.Text = ""
    txtUnicastOut.Text = ""
    txtUnknownProtocolIn.Text = ""
End Sub

Private Sub Form_Load()
    If Function_Exist("iphlpapi.dll", "GetIfTable") = True Then
        cmdRefresh_Click
    Else
        lblEntry.Enabled = False
        lstMIB2IF_Table.Enabled = False
        lblAdminStatus.Enabled = False
        lblDescription.Enabled = False
        lblInterfaceIndex.Enabled = False
        lblInterfaceType.Enabled = False
        lblLastStatusChange.Enabled = False
        lblMTU.Enabled = False
        lblName.Enabled = False
        lblOperationalStatus.Enabled = False
        lblOutputQueueLength.Enabled = False
        lblPhysicalAddress.Enabled = False
        lblSpeed.Enabled = False
        lblPacketsIn.Enabled = False
        lblOctetsIn.Enabled = False
        lblUnicastIn.Enabled = False
        lblNonUnicastIn.Enabled = False
        lblDiscardedIn.Enabled = False
        lblErroneousIn.Enabled = False
        lblUnknownProtocolIn.Enabled = False
        lblPacketsOut.Enabled = False
        lblOctetsOut.Enabled = False
        lblUnicastOut.Enabled = False
        lblNonUnicastOut.Enabled = False
        lblDiscardedOut.Enabled = False
        lblErroneousOut.Enabled = False
        lblSorted.Enabled = False
        cmdRefresh.Enabled = False
    End If
    
    chkSorted.value = GetRegSetting(HKEY_CURRENT_USER, "Software\Kira\MIB2IFTable", "Sorted")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\MIB2IFTable", "Sorted", chkSorted.value, REG_DWORD
End Sub

Private Sub lstMIB2IF_Table_Click()
    With MIB_IFTABLE.table(lstMIB2IF_Table.ListIndex)
        txtName.Text = UnicodeToAscii(.wszName, &H0)
        txtInterfaceIndex.Text = CStr(.dwIndex)
        
        Select Case .dwType
            Case MIB_IF_TYPE_OTHER: txtInterfaceType.Text = "Other"
            Case MIB_IF_TYPE_ETHERNET: txtInterfaceType.Text = "Ethernet"
            Case MIB_IF_TYPE_TOKENRING: txtInterfaceType.Text = "Tokenring"
            Case MIB_IF_TYPE_FDDI: txtInterfaceType.Text = "FDDI"
            Case MIB_IF_TYPE_PPP: txtInterfaceType.Text = "PPP"
            Case MIB_IF_TYPE_LOOPBACK: txtInterfaceType.Text = "Loopback"
            Case MIB_IF_TYPE_SLIP: txtInterfaceType.Text = "Slip"
            Case Else: txtInterfaceType.Text = "Unknown"
        End Select
        
        txtMTU.Text = CStr(.dwMtu)
        txtSpeed.Text = CStr(.dwSpeed)
        
        If Len(.bPhysAddr) >= .dwPhysAddrLen Then
            Dim strAddress As String
            Dim lngIncrement As Long
            
            strAddress = Left$(.bPhysAddr, .dwPhysAddrLen)
            txtPhysicalAddress.Text = ""
            
            For lngIncrement = 1 To .dwPhysAddrLen
                txtPhysicalAddress.Text = txtPhysicalAddress.Text & ltoa_(Asc(Mid$(.bPhysAddr, lngIncrement, 1)), 16)
            Next lngIncrement
        End If
        
        Select Case .dwAdminStatus
            Case MIB_IF_ADMIN_STATUS_UP: txtAdminStatus.Text = "Up"
            Case MIB_IF_ADMIN_STATUS_DOWN: txtAdminStatus.Text = "Down"
            Case MIB_IF_ADMIN_STATUS_TESTING: txtAdminStatus.Text = "Testing"
            Case Else: txtAdminStatus.Text = "Unknown"
        End Select
        
        Select Case .dwOperStatus
            Case MIB_IF_OPER_STATUS_NON_OPERATIONAL: txtOperationalStatus.Text = "Non Operational"
            Case MIB_IF_OPER_STATUS_UNREACHABLE: txtOperationalStatus.Text = "Unreachable"
            Case MIB_IF_OPER_STATUS_DISCONNECTED: txtOperationalStatus.Text = "Disconnected"
            Case MIB_IF_OPER_STATUS_CONNECTING: txtOperationalStatus.Text = "Connecting"
            Case MIB_IF_OPER_STATUS_CONNECTED: txtOperationalStatus.Text = "Connected"
            Case MIB_IF_OPER_STATUS_OPERATIONAL: txtOperationalStatus.Text = "Operational"
            Case Else: txtOperationalStatus.Text = "Unknown"
        End Select
        
        txtLastStatusChange.Text = CStr(.dwLastChange)
        txtOctetsIn.Text = CStr(.dwInOctets)
        txtUnicastIn.Text = CStr(.dwInUcastPkts)
        txtNonUnicastIn.Text = CStr(.dwInNUcastPkts)
        txtDiscardedIn.Text = CStr(.dwInDiscards)
        txtErroneousIn.Text = CStr(.dwInErrors)
        txtUnknownProtocolIn.Text = CStr(.dwInUnknownProtos)
        txtOctetsOut.Text = CStr(.dwOutOctets)
        txtUnicastOut.Text = CStr(.dwOutUcastPkts)
        txtNonUnicastOut.Text = CStr(.dwOutNUcastPkts)
        txtDiscardedOut.Text = CStr(.dwOutDiscards)
        txtErroneousOut.Text = CStr(.dwOutErrors)
        txtOutputQueueLength.Text = CStr(.dwOutQLen)
        
        If Len(.bDescr) >= .dwDescrLen Then
            txtDescription.Text = Left$(.bDescr, .dwDescrLen)
        End If
    End With
End Sub

