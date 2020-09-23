VERSION 5.00
Begin VB.Form frmUDPTable 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "UDP Table"
   ClientHeight    =   2610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3015
   Icon            =   "frmUDPTable.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2610
   ScaleWidth      =   3015
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkSorted 
      Height          =   255
      Left            =   2640
      TabIndex        =   7
      Top             =   1800
      Width           =   255
   End
   Begin VB.TextBox txtPort 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1440
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
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton cmdRefresh 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Refresh"
      Height          =   350
      Left            =   1920
      TabIndex        =   8
      Top             =   2160
      Width           =   975
   End
   Begin VB.ListBox lstUDP_Table 
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
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label lblSorted 
      Caption         =   "Sorted"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label lblPort 
      Caption         =   "Port"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label lblIPAddress 
      Caption         =   "IP Address"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label lblEntry 
      Caption         =   "Entry"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmUDPTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim MIB_UDPTABLE As MIB_UDPTABLE

Private Sub cmdRefresh_Click()
    Dim lngSize As Long
    lngSize = Len(MIB_UDPTABLE)
    
    If GetUdpTable(MIB_UDPTABLE, lngSize, chkSorted.value) <> NO_ERROR Then Failed "GetUdpTable"
    
    With lstUDP_Table
        .Clear
        
        Dim lngIncrement As Long
        For lngIncrement = 0 To MIB_UDPTABLE.dwNumEntries - 1
            .AddItem CStr(lngIncrement + 1)
        Next lngIncrement
    End With
    
    txtIPAddress.Text = ""
    txtPort.Text = ""
End Sub

Private Sub Form_Load()
    If Function_Exist("iphlpapi.dll", "GetUdpTable") = True Then
        cmdRefresh_Click
    Else
        lblEntry.Enabled = False
        lblIPAddress.Enabled = False
        lblPort.Enabled = False
        lblSorted.Enabled = False
        chkSorted.Enabled = False
        cmdRefresh.Enabled = False
    End If
    
    chkSorted.value = GetRegSetting(HKEY_CURRENT_USER, "Software\Kira\UDPTable", "Sorted")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\UDPTable", "Sorted", chkSorted.value, REG_DWORD
End Sub

Private Sub lstUDP_Table_Click()
    With MIB_UDPTABLE.table(lstUDP_Table.ListIndex)
        txtIPAddress.Text = CStr(.dwLocalAddr(0)) & "." & CStr(.dwLocalAddr(1)) & "." & CStr(.dwLocalAddr(2)) & "." & CStr(.dwLocalAddr(3))
        txtPort.Text = strtoul_(Right$("00" & ltoa_(Asc(Mid$(.dwLocalPort, 1, 1)), 16), 2) & _
                                Right$("00" & ltoa_(Asc(Mid$(.dwLocalPort, 2, 1)), 16), 2), 16)
    End With
End Sub
