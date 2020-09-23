VERSION 5.00
Begin VB.Form frmIPNetTable 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "IP Net Table"
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3375
   Icon            =   "frmIPNetTable.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   3375
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkSorted 
      Height          =   255
      Left            =   3000
      TabIndex        =   11
      Top             =   2280
      Width           =   255
   End
   Begin VB.TextBox txtARP 
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
   Begin VB.TextBox txtIPAddress 
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
   Begin VB.TextBox txtPhysicalAddress 
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
   Begin VB.TextBox txtAdapterIndex 
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
   Begin VB.ListBox lstIPNet_Table 
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
      TabIndex        =   12
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label lblSorted 
      Caption         =   "Sorted"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label lblARP 
      Caption         =   "ARP"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label lblIPAddress 
      Caption         =   "IP Address"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label lblPhysicalAddress 
      Caption         =   "Physical Address"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label lblAdapterIndex 
      Caption         =   "Adapter Index"
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
Attribute VB_Name = "frmIPNetTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim MIB_IPNETTABLE As MIB_IPNETTABLE

Private Sub cmdRefresh_Click()
    Dim lngSize As Long
    lngSize = Len(MIB_IPNETTABLE)
    
    If GetIpNetTable(MIB_IPNETTABLE, lngSize, chkSorted.value) <> NO_ERROR Then Failed "GetIpNetTable"
    
    
    With lstIPNet_Table
        .Clear
        
        Dim lngIncrement As Long
        For lngIncrement = 0 To MIB_IPNETTABLE.dwNumEntries - 1
            .AddItem CStr(lngIncrement + 1)
        Next lngIncrement
    End With
    
    txtAdapterIndex.Text = ""
    txtPhysicalAddress.Text = ""
    txtIPAddress.Text = ""
    txtARP.Text = ""
End Sub

Private Sub Form_Load()
    If Function_Exist("iphlpapi.dll", "GetIpNetTable") = True Then
        cmdRefresh_Click
    Else
        lblEntry.Enabled = False
        lstIPNet_Table.Enabled = False
        lblAdapterIndex.Enabled = False
        lblPhysicalAddress.Enabled = False
        lblIPAddress.Enabled = False
        lblARP.Enabled = False
        lblSorted.Enabled = False
        chkSorted.Enabled = False
        cmdRefresh.Enabled = False
    End If

    chkSorted.value = GetRegSetting(HKEY_CURRENT_USER, "Software\Kira\IPNetTable", "Sorted")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\IPNetTable", "Sorted", chkSorted.value, REG_DWORD
End Sub

Private Sub lstIPNet_Table_Click()
    With MIB_IPNETTABLE.table(lstIPNet_Table.ListIndex)
        txtAdapterIndex.Text = CStr(.dwIndex)
        
        Dim lngIncrement As Long
        For lngIncrement = 0 To .dwPhysAddrLen
            txtPhysicalAddress.Text = txtPhysicalAddress.Text & Chr$(.bPhysAddr(lngIncrement))
        Next lngIncrement
        
        txtIPAddress.Text = CStr(.dwAddr(0)) & "." & CStr(.dwAddr(1)) & "." & CStr(.dwAddr(2)) & "." & CStr(.dwAddr(3))
        
        Select Case .dwType
            Case 4: txtARP.Text = "Static"
            Case 3: txtARP.Text = "Dynamic"
            Case 2: txtARP.Text = "Invalid"
            Case 1: txtARP.Text = "Other"
            Case Else: txtARP.Text = "Unknown"
        End Select
    End With
End Sub
