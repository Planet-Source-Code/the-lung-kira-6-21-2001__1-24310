VERSION 5.00
Begin VB.Form frmTCPTable 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TCP Table"
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3495
   Icon            =   "frmTCPTable.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   3495
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cboConnectionState 
      Height          =   315
      Left            =   1920
      TabIndex        =   3
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CheckBox chkSorted 
      Height          =   255
      Left            =   3120
      TabIndex        =   13
      Top             =   2760
      Width           =   255
   End
   Begin VB.TextBox txtRemotePort 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   2400
      Width           =   1455
   End
   Begin VB.TextBox txtLocalPort 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1800
      Width           =   1455
   End
   Begin VB.TextBox txtRemoteAddress 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   2160
      Width           =   1455
   End
   Begin VB.TextBox txtLocalAddress 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CommandButton cmdRefresh 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Refresh"
      Height          =   350
      Left            =   2400
      TabIndex        =   15
      Top             =   3120
      Width           =   975
   End
   Begin VB.ListBox lstTCP_Table 
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
      Left            =   1920
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton cmdApply 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Apply"
      Height          =   350
      Left            =   1440
      TabIndex        =   14
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label lblSorted 
      Caption         =   "Sorted"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Label lblRemotePort 
      Caption         =   "Remote Port"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label lblRemoteAddress 
      Caption         =   "Remote Address"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label lblLocalPort 
      Caption         =   "Local Port"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label lblLocalAddress 
      Caption         =   "Local Address"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label lblConnectionState 
      Caption         =   "Connection State"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label lblEntry 
      Caption         =   "Entry"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmTCPTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim MIB_TCPTABLE As MIB_TCPTABLE

Private Sub cmdApply_Click()
    Select Case cboConnectionState.ListIndex
        Case 0: MIB_TCPTABLE.table(lstTCP_Table.ListIndex).dwState = MIB_TCP_STATE_CLOSED
        Case 1: MIB_TCPTABLE.table(lstTCP_Table.ListIndex).dwState = MIB_TCP_STATE_LISTEN
        Case 2: MIB_TCPTABLE.table(lstTCP_Table.ListIndex).dwState = MIB_TCP_STATE_SYN_SENT
        Case 3: MIB_TCPTABLE.table(lstTCP_Table.ListIndex).dwState = MIB_TCP_STATE_SYN_RCVD
        Case 4: MIB_TCPTABLE.table(lstTCP_Table.ListIndex).dwState = MIB_TCP_STATE_ESTAB
        Case 5: MIB_TCPTABLE.table(lstTCP_Table.ListIndex).dwState = MIB_TCP_STATE_FIN_WAIT1
        Case 6: MIB_TCPTABLE.table(lstTCP_Table.ListIndex).dwState = MIB_TCP_STATE_FIN_WAIT2
        Case 7: MIB_TCPTABLE.table(lstTCP_Table.ListIndex).dwState = MIB_TCP_STATE_CLOSE_WAIT
        Case 8: MIB_TCPTABLE.table(lstTCP_Table.ListIndex).dwState = MIB_TCP_STATE_CLOSING
        Case 9: MIB_TCPTABLE.table(lstTCP_Table.ListIndex).dwState = MIB_TCP_STATE_LAST_ACK
        Case 10: MIB_TCPTABLE.table(lstTCP_Table.ListIndex).dwState = MIB_TCP_STATE_TIME_WAIT
        Case 11: MIB_TCPTABLE.table(lstTCP_Table.ListIndex).dwState = MIB_TCP_STATE_DELETE_TCB
    End Select
    
    If SetTcpEntry(MIB_TCPTABLE.table(lstTCP_Table.ListIndex)) <> NO_ERROR Then Failed "SetTcpEntry"
End Sub

Private Sub cmdRefresh_Click()
    Dim lngSize As Long
    lngSize = Len(MIB_TCPTABLE)
    
    If GetTcpTable(MIB_TCPTABLE, lngSize, chkSorted.value) <> NO_ERROR Then Failed "GetTcpTable"
    
    
    With lstTCP_Table
        .Clear
        
        Dim lngIncrement As Long
        For lngIncrement = 0 To MIB_TCPTABLE.dwNumEntries - 1
            .AddItem CStr(lngIncrement + 1)
        Next lngIncrement
    End With
    
    cboConnectionState.ListIndex = -1
    txtLocalAddress.Text = ""
    txtLocalPort.Text = ""
    txtRemoteAddress.Text = ""
    txtRemotePort.Text = ""
End Sub

Private Sub Form_Load()
    If Function_Exist("iphlpapi.dll", "GetTcpTable") = True Then
        cmdRefresh_Click
    Else
        lblEntry.Enabled = False
        lstTCP_Table.Enabled = False
        lblConnectionState.Enabled = False
        cboConnectionState.Enabled = False
        lblLocalAddress.Enabled = False
        lblLocalPort.Enabled = False
        lblRemoteAddress.Enabled = False
        lblRemotePort.Enabled = False
        lblSorted.Enabled = False
        chkSorted.Enabled = False
        cmdApply.Enabled = False
        cmdRefresh.Enabled = False
    End If
    
    With cboConnectionState
        .AddItem "Closed"
        .AddItem "Listen"
        .AddItem "Syn Sent"
        .AddItem "Syn Received"
        .AddItem "Established"
        .AddItem "Finished Wait1"
        .AddItem "Finished Wait2"
        .AddItem "Close Wait"
        .AddItem "Closing"
        .AddItem "Last Acknowledge"
        .AddItem "Time Wait"
        .AddItem "Delete TCB"
        .AddItem "Unknown"
    End With
    
    chkSorted.value = GetRegSetting(HKEY_CURRENT_USER, "Software\Kira\TCPTable", "Sorted")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\TCPTable", "Sorted", chkSorted.value, REG_DWORD
End Sub

Private Sub lstTCP_Table_Click()
    With MIB_TCPTABLE.table(lstTCP_Table.ListIndex)
        Select Case .dwState
            Case MIB_TCP_STATE_CLOSED: cboConnectionState.ListIndex = 0
            Case MIB_TCP_STATE_LISTEN: cboConnectionState.ListIndex = 1
            Case MIB_TCP_STATE_SYN_SENT: cboConnectionState.ListIndex = 2
            Case MIB_TCP_STATE_SYN_RCVD: cboConnectionState.ListIndex = 3
            Case MIB_TCP_STATE_ESTAB: cboConnectionState.ListIndex = 4
            Case MIB_TCP_STATE_FIN_WAIT1: cboConnectionState.ListIndex = 5
            Case MIB_TCP_STATE_FIN_WAIT2: cboConnectionState.ListIndex = 6
            Case MIB_TCP_STATE_CLOSE_WAIT: cboConnectionState.ListIndex = 7
            Case MIB_TCP_STATE_CLOSING: cboConnectionState.ListIndex = 8
            Case MIB_TCP_STATE_LAST_ACK: cboConnectionState.ListIndex = 9
            Case MIB_TCP_STATE_TIME_WAIT: cboConnectionState.ListIndex = 10
            Case MIB_TCP_STATE_DELETE_TCB: cboConnectionState.ListIndex = 11
            Case Else: cboConnectionState.ListIndex = 12
        End Select
        
        txtLocalAddress.Text = CStr(.dwLocalAddr(0)) & "." & CStr(.dwLocalAddr(1)) & "." & CStr(.dwLocalAddr(2)) & "." & CStr(.dwLocalAddr(3))
        txtLocalPort.Text = strtoul_(Right$("00" & ltoa_(Asc(Mid$(.dwLocalPort, 1, 1)), 16), 2) & _
                                     Right$("00" & ltoa_(Asc(Mid$(.dwLocalPort, 2, 1)), 16), 2), 16)
        txtRemoteAddress.Text = CStr(.dwRemoteAddr(0)) & "." & CStr(.dwRemoteAddr(1)) & "." & CStr(.dwRemoteAddr(2)) & "." & CStr(.dwRemoteAddr(3))
        txtRemotePort.Text = strtoul_(Right$("00" & ltoa_(Asc(Mid$(.dwRemotePort, 1, 1)), 16), 2) & _
                                      Right$("00" & ltoa_(Asc(Mid$(.dwRemotePort, 2, 1)), 16), 2), 16)
    End With
End Sub
