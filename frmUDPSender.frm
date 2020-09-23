VERSION 5.00
Begin VB.Form frmUDPSender 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "UDP Sender"
   ClientHeight    =   2355
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3975
   Icon            =   "frmUDPSender.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2355
   ScaleWidth      =   3975
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtPort 
      Height          =   285
      Left            =   1560
      TabIndex        =   3
      Top             =   480
      Width           =   2295
   End
   Begin VB.TextBox txtSent 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   1560
      Width           =   2295
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   350
      Left            =   2880
      TabIndex        =   11
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Enabled         =   0   'False
      Height          =   350
      Left            =   1920
      TabIndex        =   10
      Top             =   1920
      Width           =   975
   End
   Begin VB.TextBox txtNumber 
      Height          =   285
      Left            =   1560
      TabIndex        =   7
      Top             =   1200
      Width           =   2295
   End
   Begin VB.TextBox txtDataSize 
      Height          =   285
      Left            =   1560
      TabIndex        =   5
      Top             =   840
      Width           =   2295
   End
   Begin VB.TextBox txtHostIP 
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label lblPort 
      Caption         =   "Port"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label lblSent 
      Caption         =   "Sent"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label lblNumber 
      Caption         =   "Number"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label lblDataSize 
      Caption         =   "Data Size"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
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
Attribute VB_Name = "frmUDPSender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lngSocket As Long
Dim sockaddr As sockaddr

Private Sub cmdStart_Click()
    cmdStop.Enabled = True
    cmdStart.Enabled = False
    
    
    Dim strData As String
    Dim lngNumber As Long
    Dim lngIncrement As Long
    
    Close_Socket lngSocket
    
    lngSocket = socket(AF_INET, SOCK_DGRAM, IPPROTO_UDP): If lngSocket = INVALID_SOCKET Then WinsockError "socket"
    
    With sockaddr
        .sin_addr = HostIPToInAddr(txtHostIP.Text & Chr$(0))
        .sin_family = AF_INET
        .sin_port = htons(Val(txtPort.Text))
        .sin_zero = String$(8, 0)
    End With
    
    strData = String$(Val(txtDataSize.Text), Chr$(0))
    lngNumber = Val(txtNumber.Text)
    
    For lngIncrement = 1 To lngNumber
        If sendto(lngSocket, ByVal strData, Len(strData), 0, sockaddr, Len(sockaddr)) = SOCKET_ERROR Then
            WinsockError "sendto"
            Exit For
        End If
        
        txtSent.Text = lngIncrement
        DoEvents
    Next lngIncrement
    
    If shutdown(lngSocket, SD_BOTH) = SOCKET_ERROR Then WinsockError "shutdown"
    cmdStop.Enabled = False
    cmdStart.Enabled = True
End Sub

Private Sub cmdStop_Click()
    If shutdown(lngSocket, SD_BOTH) = SOCKET_ERROR Then WinsockError "shutdown"
    
    cmdStop.Enabled = False
    cmdStart.Enabled = True
End Sub

Private Sub Form_Load()
    txtDataSize.Text = GetRegSetting(HKEY_CURRENT_USER, "Software\Kira\UDPSender", "DataSize")
    txtHostIP.Text = GetRegSetting(HKEY_CURRENT_USER, "Software\Kira\UDPSender", "HostIP")
    txtNumber.Text = GetRegSetting(HKEY_CURRENT_USER, "Software\Kira\UDPSender", "Number")
    txtPort.Text = GetRegSetting(HKEY_CURRENT_USER, "Software\Kira\UDPSender", "Port")
    
    If WS2 = False Then
        lblHostIP.Enabled = False
        txtHostIP.Enabled = False
        lblPort.Enabled = False
        txtPort.Enabled = False
        lblDataSize.Enabled = False
        txtDataSize.Enabled = False
        lblNumber.Enabled = False
        txtNumber.Enabled = False
        lblSent.Enabled = False
        cmdStop.Enabled = False
        cmdStart.Enabled = False
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\UDPSender", "DataSize", txtDataSize.Text, REG_DWORD
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\UDPSender", "HostIP", txtHostIP.Text, REG_SZ
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\UDPSender", "Number", txtNumber.Text, REG_DWORD
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\UDPSender", "Port", txtPort.Text, REG_DWORD
    
    If lngSocket <> 0 Then
        If shutdown(lngSocket, SD_BOTH) = SOCKET_ERROR Then WinsockError "shutdown"
        Close_Socket lngSocket
    End If
End Sub

Private Sub txtDataSize_Change()
    txtDataSize.Text = CStr(Val(Rem_NonNumeric_Chr(txtDataSize.Text)))
    If Val(txtDataSize.Text) < 0 Then txtDataSize.Text = "0"
    If Val(txtDataSize.Text) > 65535 Then txtDataSize.Text = "65535"
End Sub

Private Sub txtNumber_Change()
    txtNumber.Text = CStr(Val(Rem_NonNumeric_Chr(txtNumber.Text)))
    If Val(txtNumber.Text) < 1 Then txtNumber.Text = "1"
    If Val(txtNumber.Text) > 2147483647 Then txtNumber.Text = "2147483647"
End Sub

Private Sub txtPort_Change()
    txtPort.Text = CStr(Val(Rem_NonNumeric_Chr(txtPort.Text)))
    If Val(txtPort.Text) < 0 Then txtPort.Text = "0"
    If Val(txtPort.Text) > 65535 Then txtPort.Text = "65535"
End Sub
