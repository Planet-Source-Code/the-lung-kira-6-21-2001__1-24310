VERSION 5.00
Begin VB.Form frmTime 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Time"
   ClientHeight    =   3450
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3975
   Icon            =   "frmTime.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   3975
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtPort 
      Height          =   285
      Left            =   1560
      TabIndex        =   3
      Top             =   480
      Width           =   2295
   End
   Begin VB.CheckBox chkDaylightSavings 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   3600
      TabIndex        =   13
      Top             =   2280
      Width           =   255
   End
   Begin VB.CommandButton cmdSetTime 
      Caption         =   "Set Time"
      Height          =   350
      Left            =   2880
      TabIndex        =   19
      Top             =   3000
      Width           =   975
   End
   Begin VB.TextBox txtRoundTripTime 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   2640
      Width           =   2295
   End
   Begin VB.TextBox txtReturnedLocal 
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   1560
      Width           =   2295
   End
   Begin VB.CommandButton cmdGetData 
      Caption         =   "Get Data"
      Height          =   350
      Left            =   1800
      TabIndex        =   18
      Top             =   3000
      Width           =   975
   End
   Begin VB.ComboBox cboMethod 
      Height          =   315
      Left            =   1560
      TabIndex        =   5
      Top             =   840
      Width           =   2295
   End
   Begin VB.TextBox txtReturnedGMT 
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1200
      Width           =   2295
   End
   Begin VB.TextBox txtHostIP 
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Top             =   120
      Width           =   2295
   End
   Begin VB.TextBox txtUnFormatted 
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   1920
      Width           =   2295
   End
   Begin VB.TextBox txtTime 
      Height          =   195
      Left            =   120
      TabIndex        =   16
      Top             =   3000
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Enabled         =   0   'False
      Height          =   350
      Left            =   840
      TabIndex        =   17
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label lblPort 
      Caption         =   "Port"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label lblDaylightSavings 
      Caption         =   "Daylight Savings"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label lblRoundTripTime 
      Caption         =   "Round Trip Time"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label lblReturnedLocal 
      Caption         =   "Returned Local"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1560
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
   Begin VB.Label lblReturnedGMT 
      Caption         =   "Returned GMT"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1200
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
   Begin VB.Label lblUnFormatted 
      Caption         =   "UnFormatted"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1920
      Width           =   1215
   End
End
Attribute VB_Name = "frmTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdGetData_Click()
    cmdSetTime.Enabled = False
    cmdGetData.Enabled = False
    cmdStop.Enabled = True
    txtReturnedGMT.Text = ""
    txtReturnedLocal.Text = ""
    txtUnFormatted.Text = ""
    txtRoundTripTime.Text = ""
    
    
    Close_Socket wsTime_Socket
    wsTime_RTT = 0
    
    
    Select Case cboMethod.ListIndex
        Case 0 'UDP
            wsTime_Socket = socket(AF_INET, SOCK_DGRAM, IPPROTO_UDP): If wsTime_Socket = INVALID_SOCKET Then WinsockError "socket"
            
            If WSAAsyncSelect(wsTime_Socket, txtTime.hwnd, WM_PROJECT_WS, FD_CLOSE Or FD_READ) = SOCKET_ERROR Then WinsockError "WSAAsyncSelect"
            
            With wsTime_sockaddr
                .sin_addr = HostIPToInAddr(txtHostIP.Text & Chr$(0))
                .sin_family = AF_INET
                .sin_port = htons(strtoul_(txtPort.Text, 10))
                .sin_zero = String$(8, 0)
            End With
            
            If sendto(wsTime_Socket, ByVal "", 0, 0, wsTime_sockaddr, Len(wsTime_sockaddr)) = SOCKET_ERROR Then WinsockError "sendto"
            wsTime_RTT = PerformanceCounter
        Case 1 'TCP
            wsTime_Socket = socket(AF_INET, SOCK_STREAM, IPPROTO_TCP): If wsTime_Socket = INVALID_SOCKET Then WinsockError "socket"
            
            With wsTime_sockaddr
                .sin_addr = HostIPToInAddr(txtHostIP.Text & Chr$(0))
                .sin_family = AF_INET
                .sin_port = htons(strtoul_(txtPort.Text, 10))
                .sin_zero = String$(8, 0)
            End With
            
            If connect(wsTime_Socket, wsTime_sockaddr, Len(wsTime_sockaddr)) = SOCKET_ERROR Then WinsockError "connect"
            wsTime_RTT = PerformanceCounter
            If WSAAsyncSelect(wsTime_Socket, txtTime.hwnd, WM_PROJECT_WS, FD_CLOSE Or FD_READ) = SOCKET_ERROR Then WinsockError "WSAAsyncSelect"
    End Select
End Sub

Private Sub cmdSetTime_Click()
    wsTime_SetTime = True
    cmdGetData_Click
End Sub

Private Sub cmdStop_Click()
    If shutdown(wsTime_Socket, SD_BOTH) = SOCKET_ERROR Then WinsockError "shutdown"
    
    wsTime_SetTime = False
    
    cmdStop.Enabled = False
    cmdGetData.Enabled = True
    cmdSetTime.Enabled = True
End Sub

Private Sub Form_Load()
    With cboMethod
        .AddItem "UDP"
        .AddItem "TCP"
    End With
    
    chkDaylightSavings.value = GetRegSetting(HKEY_CURRENT_USER, "Software\Kira\Time", "DaylightSavings")
    txtHostIP.Text = GetRegSetting(HKEY_CURRENT_USER, "Software\Kira\Time", "HostIP")
    cboMethod.ListIndex = GetRegSetting(HKEY_CURRENT_USER, "Software\Kira\Time", "Method")
    txtPort.Text = GetRegSetting(HKEY_CURRENT_USER, "Software\Kira\Time", "Port")
    
    
    wsTime_OldProc = SetWindowLong(txtTime.hwnd, GWL_WNDPROC, AddressOf wsTime_Proc): If wsTime_OldProc = 0 Then Failed "SetWindowLong"
    
    
    If WS2 = False Then
        lblHostIP.Enabled = False
        txtHostIP.Enabled = False
        lblPort.Enabled = False
        txtPort.Enabled = False
        lblMethod.Enabled = False
        cboMethod.Enabled = False
        lblReturnedGMT.Enabled = False
        txtReturnedGMT.Enabled = False
        lblReturnedLocal.Enabled = False
        txtReturnedLocal.Enabled = False
        lblUnFormatted.Enabled = False
        txtUnFormatted.Enabled = False
        lblDaylightSavings.Enabled = False
        chkDaylightSavings.Enabled = False
        lblRoundTripTime.Enabled = False
        cmdStop.Enabled = False
        cmdGetData.Enabled = False
        cmdSetTime.Enabled = False
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\Time", "DaylightSavings", chkDaylightSavings.value, REG_DWORD
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\Time", "HostIP", txtHostIP.Text, REG_SZ
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\Time", "Method", cboMethod.ListIndex, REG_DWORD
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\Time", "Port", txtPort.Text, REG_DWORD
    
    If wsTime_Socket <> 0 Then
        If shutdown(wsTime_Socket, SD_BOTH) = SOCKET_ERROR Then WinsockError "shutdown"
        Close_Socket wsTime_Socket
        
        wsTime_RTT = 0
        wsTime_SetTime = False
        
        Dim sockaddr As sockaddr
        wsTime_sockaddr = sockaddr
    End If
    
    If SetWindowLong(txtTime.hwnd, GWL_WNDPROC, wsTime_OldProc) = 0 Then Failed "SetWindowLong"
End Sub

Private Sub txtPort_Change()
    txtPort.Text = CStr(Val(Rem_NonNumeric_Chr(txtPort.Text)))
    If Val(txtPort.Text) < 0 Then txtPort.Text = "0"
    If Val(txtPort.Text) > 65535 Then txtPort.Text = "65535"
End Sub
