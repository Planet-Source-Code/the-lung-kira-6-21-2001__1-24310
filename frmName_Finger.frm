VERSION 5.00
Begin VB.Form frmName_Finger 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Name/Finger"
   ClientHeight    =   5235
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5775
   Icon            =   "frmName_Finger.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   5775
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtReturned 
      Height          =   3285
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   1440
      Width           =   5535
   End
   Begin VB.TextBox txtPort 
      Height          =   285
      Left            =   3360
      TabIndex        =   3
      Top             =   480
      Width           =   2295
   End
   Begin VB.TextBox txtHostIP 
      Height          =   285
      Left            =   3360
      TabIndex        =   1
      Top             =   120
      Width           =   2295
   End
   Begin VB.TextBox txtSend 
      Height          =   285
      Left            =   3360
      TabIndex        =   5
      Top             =   840
      Width           =   2295
   End
   Begin VB.CommandButton cmdSendData 
      Caption         =   "Send Data"
      Height          =   350
      Left            =   4680
      TabIndex        =   10
      Top             =   4800
      Width           =   975
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Enabled         =   0   'False
      Height          =   350
      Left            =   3720
      TabIndex        =   9
      Top             =   4800
      Width           =   975
   End
   Begin VB.TextBox txtName_Finger 
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   4800
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Label lblReturned 
      Caption         =   "Returned"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label lblPort 
      Caption         =   "Port"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
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
   Begin VB.Label lblSend 
      Caption         =   "Send"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1215
   End
End
Attribute VB_Name = "frmName_Finger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdSendData_Click()
    cmdSendData.Enabled = False
    cmdStop.Enabled = True
    txtReturned.Text = ""
    
    
    Close_Socket wsName_Finger_Socket
    
    wsName_Finger_Socket = socket(AF_INET, SOCK_STREAM, IPPROTO_TCP): If wsName_Finger_Socket = INVALID_SOCKET Then WinsockError "socket"
    
    With wsName_Finger_sockaddr
        .sin_addr = HostIPToInAddr(txtHostIP.Text & Chr$(0))
        .sin_family = AF_INET
        .sin_port = htons(strtoul_(txtPort.Text, 10))
        .sin_zero = String$(8, 0)
    End With
    
    If connect(wsName_Finger_Socket, wsName_Finger_sockaddr, Len(wsName_Finger_sockaddr)) = SOCKET_ERROR Then WinsockError "connect"
    If WSAAsyncSelect(wsName_Finger_Socket, txtName_Finger.hwnd, WM_PROJECT_WS, FD_CLOSE Or FD_READ) = SOCKET_ERROR Then WinsockError "WSAAsyncSelect"
    If send(wsName_Finger_Socket, ByVal txtSend.Text & Chr$(13) & Chr$(10), Len(txtSend.Text & Chr$(13) & Chr$(10)), 0) = SOCKET_ERROR Then WinsockError "send"
End Sub

Private Sub cmdStop_Click()
    If shutdown(wsName_Finger_Socket, SD_BOTH) = SOCKET_ERROR Then WinsockError "shutdown"
    
    cmdStop.Enabled = False
    cmdSendData.Enabled = True
End Sub

Private Sub Form_Load()
    txtHostIP.Text = GetRegSetting(HKEY_CURRENT_USER, "Software\Kira\Name_Finger", "HostIP")
    txtPort.Text = GetRegSetting(HKEY_CURRENT_USER, "Software\Kira\Name_Finger", "Port")
    txtSend.Text = GetRegSetting(HKEY_CURRENT_USER, "Software\Kira\Name_Finger", "Send")
    
    wsName_Finger_OldProc = SetWindowLong(txtName_Finger.hwnd, GWL_WNDPROC, AddressOf wsName_Finger_Proc): If wsName_Finger_OldProc = 0 Then Failed "SetWindowLong"
    
    
    If WS2 = False Then
        lblHostIP.Enabled = False
        txtHostIP.Enabled = False
        lblPort.Enabled = False
        txtPort.Enabled = False
        lblSend.Enabled = False
        txtSend.Enabled = False
        lblReturned.Enabled = False
        txtReturned.Enabled = False
        cmdStop.Enabled = False
        cmdSendData.Enabled = False
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\Name_Finger", "HostIP", txtHostIP.Text, REG_SZ
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\Name_Finger", "Port", txtPort.Text, REG_DWORD
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\Name_Finger", "Send", txtSend.Text, REG_SZ
    
    If wsName_Finger_Socket <> 0 Then
        If shutdown(wsName_Finger_Socket, SD_BOTH) = SOCKET_ERROR Then WinsockError "shutdown"
        Close_Socket wsName_Finger_Socket
        
        Dim sockaddr As sockaddr
        wsName_Finger_sockaddr = sockaddr
    End If
    
    If SetWindowLong(txtName_Finger.hwnd, GWL_WNDPROC, wsName_Finger_OldProc) = 0 Then Failed "SetWindowLong"
End Sub

Private Sub txtPort_Change()
    txtPort.Text = CStr(Val(Rem_NonNumeric_Chr(txtPort.Text)))
    If Val(txtPort.Text) < 0 Then txtPort.Text = "0"
    If Val(txtPort.Text) > 65535 Then txtPort.Text = "65535"
End Sub

