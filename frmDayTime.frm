VERSION 5.00
Begin VB.Form frmDayTime 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DayTime"
   ClientHeight    =   2850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3975
   Icon            =   "frmDayTime.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   3975
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtPort 
      Height          =   285
      Left            =   1560
      TabIndex        =   4
      Top             =   480
      Width           =   2295
   End
   Begin VB.TextBox txtDayTime 
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   2520
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtRoundTripTime 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   2040
      Width           =   2295
   End
   Begin VB.TextBox txtHostIP 
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Top             =   120
      Width           =   2295
   End
   Begin VB.TextBox txtReturned 
      Height          =   525
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   1440
      Width           =   3735
   End
   Begin VB.ComboBox cboMethod 
      Height          =   315
      Left            =   1560
      TabIndex        =   5
      Top             =   840
      Width           =   2295
   End
   Begin VB.CommandButton cmdGetData 
      Caption         =   "Get Data"
      Height          =   350
      Left            =   2880
      TabIndex        =   12
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Enabled         =   0   'False
      Height          =   350
      Left            =   1920
      TabIndex        =   11
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label lblPort 
      Caption         =   "Port"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label lblRoundTripTime 
      Caption         =   "Round Trip Time"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label lblHostIP 
      Caption         =   "Host / IP"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblReturned 
      Caption         =   "Returned"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label lblMethod 
      Caption         =   "Method"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
End
Attribute VB_Name = "frmDayTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdGetData_Click()
    cmdGetData.Enabled = False
    cmdStop.Enabled = True
    txtReturned.Text = ""
    txtRoundTripTime.Text = ""
    
    
    Close_Socket wsDayTime_Socket
    wsDayTime_RTT = 0
    
    
    Select Case cboMethod.ListIndex
        Case 0 'UDP
            wsDayTime_Socket = socket(AF_INET, SOCK_DGRAM, IPPROTO_UDP): If wsDayTime_Socket = INVALID_SOCKET Then WinsockError "socket"
            If WSAAsyncSelect(wsDayTime_Socket, txtDayTime.hwnd, WM_PROJECT_WS, FD_CLOSE Or FD_READ) = SOCKET_ERROR Then WinsockError "WSAAsyncSelect"
            
            With wsDayTime_sockaddr
                .sin_addr = HostIPToInAddr(txtHostIP.Text & Chr$(0))
                .sin_family = AF_INET
                .sin_port = htons(strtoul_(txtPort.Text, 10))
                .sin_zero = String$(8, 0)
            End With
            
            If sendto(wsDayTime_Socket, ByVal "", 0, 0, wsDayTime_sockaddr, Len(wsDayTime_sockaddr)) = SOCKET_ERROR Then WinsockError "sendto"
            wsDayTime_RTT = PerformanceCounter
        Case 1 'TCP
            wsDayTime_Socket = socket(AF_INET, SOCK_STREAM, IPPROTO_TCP): If wsDayTime_Socket = INVALID_SOCKET Then WinsockError "socket"
            
            With wsDayTime_sockaddr
                .sin_addr = HostIPToInAddr(txtHostIP.Text & Chr$(0))
                .sin_family = AF_INET
                .sin_port = htons(strtoul_(txtPort.Text, 10))
                .sin_zero = String$(8, 0)
            End With
            
            If connect(wsDayTime_Socket, wsDayTime_sockaddr, Len(wsDayTime_sockaddr)) = SOCKET_ERROR Then WinsockError "connect"
            wsDayTime_RTT = PerformanceCounter
            If WSAAsyncSelect(wsDayTime_Socket, txtDayTime.hwnd, WM_PROJECT_WS, FD_CLOSE Or FD_READ) = SOCKET_ERROR Then WinsockError "WSAAsyncSelect"
    End Select
End Sub

Private Sub cmdStop_Click()
    If shutdown(wsDayTime_Socket, SD_BOTH) = SOCKET_ERROR Then WinsockError "shutdown"
    
    cmdStop.Enabled = False
    cmdGetData.Enabled = True
End Sub

Private Sub Form_Load()
    With cboMethod
        .AddItem "UDP"
        .AddItem "TCP"
    End With
    
    txtHostIP.Text = GetRegSetting(HKEY_CURRENT_USER, "Software\Kira\DayTime", "HostIP")
    cboMethod.ListIndex = GetRegSetting(HKEY_CURRENT_USER, "Software\Kira\DayTime", "Method")
    txtPort.Text = GetRegSetting(HKEY_CURRENT_USER, "Software\Kira\DayTime", "Port")
    
    wsDayTime_OldProc = SetWindowLong(txtDayTime.hwnd, GWL_WNDPROC, AddressOf wsDayTime_Proc): If wsDayTime_OldProc = 0 Then Failed "SetWindowLong"
    
    
    If WS2 = False Then
        lblHostIP.Enabled = False
        txtHostIP.Enabled = False
        lblPort.Enabled = False
        txtPort.Enabled = False
        lblMethod.Enabled = False
        cboMethod.Enabled = False
        lblReturned.Enabled = False
        txtReturned.Enabled = False
        lblRoundTripTime.Enabled = False
        cmdStop.Enabled = False
        cmdGetData.Enabled = False
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\DayTime", "HostIP", txtHostIP.Text, REG_SZ
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\DayTime", "Method", cboMethod.ListIndex, REG_DWORD
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\DayTime", "Port", txtPort.Text, REG_DWORD
    
    If wsDayTime_Socket <> 0 Then
        If shutdown(wsDayTime_Socket, SD_BOTH) = SOCKET_ERROR Then WinsockError "shutdown"
        Close_Socket wsDayTime_Socket
        
        wsDayTime_RTT = 0
        
        Dim sockaddr As sockaddr
        wsDayTime_sockaddr = sockaddr
    End If
    
    If SetWindowLong(txtDayTime.hwnd, GWL_WNDPROC, wsDayTime_OldProc) = 0 Then Failed "SetWindowLong"
End Sub

Private Sub txtPort_Change()
    txtPort.Text = CStr(Val(Rem_NonNumeric_Chr(txtPort.Text)))
    If Val(txtPort.Text) < 0 Then txtPort.Text = "0"
    If Val(txtPort.Text) > 65535 Then txtPort.Text = "65535"
End Sub
