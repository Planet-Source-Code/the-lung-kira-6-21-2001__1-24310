VERSION 5.00
Begin VB.Form frmEcho 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Echo"
   ClientHeight    =   2730
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3975
   Icon            =   "frmEcho.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   3975
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtPort 
      Height          =   285
      Left            =   1560
      TabIndex        =   3
      Top             =   480
      Width           =   2295
   End
   Begin VB.CheckBox chkReturnOK 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   255
      Left            =   3600
      TabIndex        =   9
      Top             =   1560
      Width           =   255
   End
   Begin VB.TextBox txtDataSize 
      Height          =   285
      Left            =   1560
      TabIndex        =   7
      Top             =   1200
      Width           =   2295
   End
   Begin VB.TextBox txtRoundTripTime 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   1920
      Width           =   2295
   End
   Begin VB.CommandButton cmdSendData 
      Caption         =   "Send Data"
      Height          =   350
      Left            =   2880
      TabIndex        =   14
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Enabled         =   0   'False
      Height          =   350
      Left            =   1920
      TabIndex        =   13
      Top             =   2280
      Width           =   975
   End
   Begin VB.TextBox txtHostIP 
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Top             =   120
      Width           =   2295
   End
   Begin VB.ComboBox cboMethod 
      Height          =   315
      Left            =   1560
      TabIndex        =   5
      Top             =   840
      Width           =   2295
   End
   Begin VB.TextBox txtEcho 
      Height          =   195
      Left            =   120
      TabIndex        =   12
      Top             =   2280
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Label lblPort 
      Caption         =   "Port"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label lblReturnOK 
      Caption         =   "Return OK"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label lblDataSize 
      Caption         =   "Data Size"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label lblRoundTripTime 
      Caption         =   "Round Trip Time"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1920
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
   Begin VB.Label lblMethod 
      Caption         =   "Method"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1215
   End
End
Attribute VB_Name = "frmEcho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdSendData_Click()
    cmdSendData.Enabled = False
    cmdStop.Enabled = True
    chkReturnOK.value = 0
    txtRoundTripTime.Text = ""
    
    
    Close_Socket wsEcho_Socket
    wsEcho_Data = ""
    wsEcho_RTT = 0
    
    
    wsEcho_Data = String$(Val(txtDataSize.Text), Chr$(0))
    Select Case cboMethod.ListIndex
        Case 0 'UDP
            wsEcho_Socket = socket(AF_INET, SOCK_DGRAM, IPPROTO_UDP): If wsEcho_Socket = INVALID_SOCKET Then WinsockError "socket"
            If WSAAsyncSelect(wsEcho_Socket, txtEcho.hwnd, WM_PROJECT_WS, FD_CLOSE Or FD_READ) = SOCKET_ERROR Then WinsockError "WSAAsyncSelect"
            
            With wsEcho_sockaddr
                .sin_addr = HostIPToInAddr(txtHostIP.Text & Chr$(0))
                .sin_family = AF_INET
                .sin_port = htons(strtoul_(txtPort.Text, 10))
                .sin_zero = String$(8, 0)
            End With
            
            If sendto(wsEcho_Socket, ByVal wsEcho_Data, Len(wsEcho_Data), 0, wsEcho_sockaddr, Len(wsEcho_sockaddr)) = SOCKET_ERROR Then WinsockError "sendto"
            wsEcho_RTT = PerformanceCounter
        Case 1 'TCP
            wsEcho_Socket = socket(AF_INET, SOCK_STREAM, IPPROTO_TCP): If wsEcho_Socket = INVALID_SOCKET Then WinsockError "socket"
            
            With wsEcho_sockaddr
                .sin_addr = HostIPToInAddr(txtHostIP.Text & Chr$(0))
                .sin_family = AF_INET
                .sin_port = htons(strtoul_(txtPort.Text, 10))
                .sin_zero = String$(8, 0)
            End With
            
            If connect(wsEcho_Socket, wsEcho_sockaddr, Len(wsEcho_sockaddr)) = SOCKET_ERROR Then WinsockError "connect"
            If WSAAsyncSelect(wsEcho_Socket, txtEcho.hwnd, WM_PROJECT_WS, FD_CLOSE Or FD_READ) = SOCKET_ERROR Then WinsockError "WSAAsyncSelect"
            
            If send(wsEcho_Socket, ByVal wsEcho_Data, Len(wsEcho_Data), 0) = SOCKET_ERROR Then WinsockError "send"
            wsEcho_RTT = PerformanceCounter
    End Select
End Sub

Private Sub cmdStop_Click()
    If shutdown(wsEcho_Socket, SD_BOTH) = SOCKET_ERROR Then WinsockError "shutdown"
    
    chkReturnOK.value = 0
    cmdStop.Enabled = False
    cmdSendData.Enabled = True
End Sub

Private Sub Form_Load()
    With cboMethod
        .AddItem "UDP"
        .AddItem "TCP"
    End With
    
    
    txtDataSize.Text = GetRegSetting(HKEY_CURRENT_USER, "Software\Kira\Echo", "DataSize")
    txtHostIP.Text = GetRegSetting(HKEY_CURRENT_USER, "Software\Kira\Echo", "HostIP")
    cboMethod.ListIndex = GetRegSetting(HKEY_CURRENT_USER, "Software\Kira\Echo", "Method")
    txtPort.Text = GetRegSetting(HKEY_CURRENT_USER, "Software\Kira\Echo", "Port")
    
    wsEcho_OldProc = SetWindowLong(txtEcho.hwnd, GWL_WNDPROC, AddressOf wsEcho_Proc): If wsEcho_OldProc = 0 Then Failed "SetWindowLong"
    
    
    If WS2 = False Then
        lblHostIP.Enabled = False
        txtHostIP.Enabled = False
        lblPort.Enabled = False
        txtPort.Enabled = False
        lblMethod.Enabled = False
        cboMethod.Enabled = False
        lblDataSize.Enabled = False
        txtDataSize.Enabled = False
        lblReturnOK.Enabled = False
        lblRoundTripTime.Enabled = False
        cmdStop.Enabled = False
        cmdSendData.Enabled = False
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\Echo", "DataSize", txtDataSize.Text, REG_DWORD
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\Echo", "HostIP", txtHostIP.Text, REG_SZ
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\Echo", "Method", cboMethod.ListIndex, REG_DWORD
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\Echo", "Port", txtPort.Text, REG_DWORD
    
    If wsEcho_Socket <> 0 Then
        If shutdown(wsEcho_Socket, SD_BOTH) = SOCKET_ERROR Then WinsockError "shutdown"
        Close_Socket wsEcho_Socket
        
        wsEcho_RTT = 0
        wsEcho_Data = ""
        
        Dim sockaddr As sockaddr
        wsEcho_sockaddr = sockaddr
    End If
    
    If SetWindowLong(txtEcho.hwnd, GWL_WNDPROC, wsEcho_OldProc) = 0 Then Failed "SetWindowLong"
End Sub

Private Sub txtDataSize_Change()
    txtDataSize.Text = CStr(Val(Rem_NonNumeric_Chr(txtDataSize.Text)))
    If Val(txtDataSize.Text) < 0 Then txtDataSize.Text = "0"
    If Val(txtDataSize.Text) > 65535 Then txtDataSize.Text = "65535"
End Sub

Private Sub txtPort_Change()
    txtPort.Text = CStr(Val(Rem_NonNumeric_Chr(txtPort.Text)))
    If Val(txtPort.Text) < 0 Then txtPort.Text = "0"
    If Val(txtPort.Text) > 65535 Then txtPort.Text = "65535"
End Sub
