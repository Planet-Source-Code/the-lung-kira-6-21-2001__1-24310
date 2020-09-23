VERSION 5.00
Begin VB.Form frmICMP_Echo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ICMP Echo"
   ClientHeight    =   3315
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3975
   Icon            =   "frmICMP_Echo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   3975
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer timerTimeout 
      Enabled         =   0   'False
      Left            =   120
      Top             =   1560
   End
   Begin VB.TextBox txtFailed 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   2520
      Width           =   2295
   End
   Begin VB.TextBox txtTimeout 
      Height          =   285
      Left            =   1560
      TabIndex        =   5
      Top             =   840
      Width           =   2295
   End
   Begin VB.TextBox txtNumber 
      Height          =   285
      Left            =   1560
      TabIndex        =   3
      Top             =   480
      Width           =   2295
   End
   Begin VB.TextBox txtAvg 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   2280
      Width           =   2295
   End
   Begin VB.ListBox lstRoundTripTime 
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
      Left            =   1560
      TabIndex        =   7
      Top             =   1200
      Width           =   2295
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Height          =   350
      Left            =   2880
      TabIndex        =   14
      Top             =   2880
      Width           =   975
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Enabled         =   0   'False
      Height          =   350
      Left            =   1920
      TabIndex        =   13
      Top             =   2880
      Width           =   975
   End
   Begin VB.TextBox txtHostIP 
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Top             =   120
      Width           =   2295
   End
   Begin VB.TextBox txtICMP_Echo 
      Height          =   195
      Left            =   120
      TabIndex        =   12
      Top             =   2880
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Label lblTimeout 
      Caption         =   "Timeout"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label lblRoundTripTime 
      Caption         =   "Round Trip Time"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label lblNumber 
      Caption         =   "Number"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label lblFailed 
      Caption         =   "Failed"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label lblAvg 
      Caption         =   "Avg"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2280
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
End
Attribute VB_Name = "frmICMP_Echo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdSend_Click()
    cmdSend.Enabled = False
    cmdStop.Enabled = True
    
    
    lstRoundTripTime.Clear
    txtAvg.Text = ""
    txtFailed.Text = ""
    wsICMP_Echo_Ret = False
    Erase wsICMP_Echo_RTTs()
    wsICMP_Echo_RTT_Num = 0
    
    
    With wsICMP_Echo_sockaddr
        .sin_addr = HostIPToInAddr(txtHostIP.Text & Chr$(0))
        .sin_family = AF_INET
        .sin_port = 0
        .sin_zero = String$(8, 0)
    End With
    Dim ICMPHDR As ICMPHDR
    With ICMPHDR
        .Type = ICMP_ECHO
        .ID = GetCurrentProcessId
        .Seq = 0
    End With
    
    
    Dim lngIncrement As Long
    For lngIncrement = 1 To Val(txtNumber.Text)
        Close_Socket wsICMP_Echo_Socket
        wsICMP_Echo_RTT = 0
        
        wsICMP_Echo_Socket = socket(AF_INET, SOCK_RAW, IPPROTO_ICMP): If wsICMP_Echo_Socket = INVALID_SOCKET Then WinsockError "socket"
        If WSAAsyncSelect(wsICMP_Echo_Socket, txtICMP_Echo.hwnd, WM_PROJECT_WS, FD_CLOSE Or FD_READ) = SOCKET_ERROR Then WinsockError "WSAAsyncSelect"
        
        With ICMPHDR
            .Seq = .Seq + 1
            .checksum = 0
            .checksum = checksum(ICMPHDR, Len(ICMPHDR))
        End With
        
        wsICMP_Echo_RTT = PerformanceCounter
        timerTimeout.Enabled = True
        If sendto(wsICMP_Echo_Socket, ICMPHDR, Len(ICMPHDR), 0, wsICMP_Echo_sockaddr, Len(wsICMP_Echo_sockaddr)) = SOCKET_ERROR Then WinsockError "sendto"
        
        Do
            DoEvents
            If wsICMP_Echo_Ret = True Then Exit Do
        Loop
        
        timerTimeout.Enabled = False
        wsICMP_Echo_Ret = False
    Next lngIncrement
    
    Dim dblAvg As Double
    Dim numAvg As Long
    Dim lngFailed As Long
    For lngIncrement = 1 To wsICMP_Echo_RTT_Num
        If wsICMP_Echo_RTTs(lngIncrement) > -1 Then
            dblAvg = dblAvg + wsICMP_Echo_RTTs(lngIncrement)
            numAvg = numAvg + 1
        Else
            lngFailed = lngFailed + 1
        End If
    Next lngIncrement
    
    If dblAvg > 0 Then dblAvg = dblAvg / numAvg
    txtAvg.Text = Round(dblAvg, 0)
    txtFailed.Text = CStr(Percentage(lngFailed, Val(txtNumber.Text), 0)) & "%"
    
    
    cmdSend.Enabled = True
    cmdStop.Enabled = False
End Sub

Private Sub cmdStop_Click()
    If shutdown(wsICMP_Echo_Socket, SD_BOTH) = SOCKET_ERROR Then WinsockError "shutdown"
    timerTimeout.Enabled = False
    wsICMP_Echo_Ret = False
    
    cmdStop.Enabled = False
    cmdSend.Enabled = True
End Sub

Private Sub Form_Load()
    txtHostIP.Text = GetRegSetting(HKEY_CURRENT_USER, "Software\Kira\ICMP_Echo", "HostIP")
    txtNumber.Text = GetRegSetting(HKEY_CURRENT_USER, "Software\Kira\ICMP_Echo", "Number")
    txtTimeout.Text = GetRegSetting(HKEY_CURRENT_USER, "Software\Kira\ICMP_Echo", "Timeout")
    
    timerTimeout.Interval = Val(txtTimeout.Text)
    wsICMP_Echo_OldProc = SetWindowLong(txtICMP_Echo.hwnd, GWL_WNDPROC, AddressOf wsICMP_Echo_Proc): If wsICMP_Echo_OldProc = 0 Then Failed "SetWindowLong"
    
    
    If WS2 = False Then
        lblHostIP.Enabled = False
        txtHostIP.Enabled = False
        lblNumber.Enabled = False
        txtNumber.Enabled = False
        lblTimeout.Enabled = False
        txtTimeout.Enabled = False
        lblRoundTripTime.Enabled = False
        lstRoundTripTime.Enabled = False
        lblAvg.Enabled = False
        lblFailed.Enabled = False
        cmdStop.Enabled = False
        cmdSend.Enabled = False
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\ICMP_Echo", "HostIP", txtHostIP.Text, REG_SZ
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\ICMP_Echo", "Number", txtNumber.Text, REG_DWORD
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\ICMP_Echo", "Timeout", txtTimeout.Text, REG_DWORD
    
    timerTimeout.Enabled = False
    
    If wsICMP_Echo_Socket <> 0 Then
        If shutdown(wsICMP_Echo_Socket, SD_BOTH) = SOCKET_ERROR Then WinsockError "shutdown"
        Close_Socket wsICMP_Echo_Socket
        
        wsICMP_Echo_Ret = False
        wsICMP_Echo_RTT = 0
        Erase wsICMP_Echo_RTTs()
        wsICMP_Echo_RTT_Num = 0
        
        Dim sockaddr As sockaddr
        wsICMP_Echo_sockaddr = sockaddr
    End If
    
    If SetWindowLong(txtICMP_Echo.hwnd, GWL_WNDPROC, wsICMP_Echo_OldProc) = 0 Then Failed "SetWindowLong"
End Sub

Private Sub timerTimeout_Timer()
    wsICMP_Echo_RTT_Num = wsICMP_Echo_RTT_Num + 1
    ReDim Preserve wsICMP_Echo_RTTs(wsICMP_Echo_RTT_Num)
    wsICMP_Echo_RTTs(wsICMP_Echo_RTT_Num) = -1
    
    lstRoundTripTime.AddItem Left$((lstRoundTripTime.ListCount + 1) & Space$(7), 7) & "Timeout"
    
    wsICMP_Echo_Ret = True
End Sub

Private Sub txtNumber_Change()
    txtNumber.Text = CStr(Val(Rem_NonNumeric_Chr(txtNumber.Text)))
    If Val(txtNumber.Text) < 0 Then txtNumber.Text = "0"
    If Val(txtNumber.Text) > 32767 Then txtNumber.Text = "32767"
End Sub

Private Sub txtTimeOut_Change()
    txtTimeout.Text = CStr(Val(Rem_NonNumeric_Chr(txtTimeout.Text)))
    If Val(txtTimeout.Text) < 0 Then txtTimeout.Text = "0"
    If Val(txtTimeout.Text) > 30000 Then txtTimeout.Text = "30000"
End Sub
