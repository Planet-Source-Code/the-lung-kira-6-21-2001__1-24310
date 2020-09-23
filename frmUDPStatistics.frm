VERSION 5.00
Begin VB.Form frmUDPStatistics 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "UDP Statistics"
   ClientHeight    =   1455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3375
   Icon            =   "frmUDPStatistics.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1455
   ScaleWidth      =   3375
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer timerUDPStatistics 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1560
      Top             =   600
   End
   Begin VB.TextBox txtListenerTableEntries 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox txtSentDatagrams 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox txtErrorsReceived 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox txtNoPort 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox txtReceivedDatagrams 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblReceivedDatagrams 
      Caption         =   "Received Datagrams"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label lblNoPort 
      Caption         =   "No Port"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label lblErrorsReceived 
      Caption         =   "Errors Received"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label lblSentDatagrams 
      Caption         =   "Sent Datagrams"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label lblListenerTableEntries 
      Caption         =   "Listener Table Entries"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1080
      Width           =   1695
   End
End
Attribute VB_Name = "frmUDPStatistics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    If Function_Exist("iphlpapi.dll", "GetUdpStatistics") = True Then
        timerUDPStatistics_Timer
        timerUDPStatistics.Enabled = True
    Else
        lblReceivedDatagrams.Enabled = False
        lblNoPort.Enabled = False
        lblErrorsReceived.Enabled = False
        lblSentDatagrams.Enabled = False
        lblListenerTableEntries.Enabled = False
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    timerUDPStatistics.Enabled = False
End Sub

Private Sub timerUDPStatistics_Timer()
    Dim MIB_UDPSTATS As MIB_UDPSTATS
    
    If GetUdpStatistics(MIB_UDPSTATS) <> NO_ERROR Then Failed "GetUdpStatistics"
    
    With MIB_UDPSTATS
        txtReceivedDatagrams.Text = CStr(.dwInDatagrams)
        txtNoPort.Text = CStr(.dwNoPorts)
        txtErrorsReceived.Text = CStr(.dwInErrors)
        txtSentDatagrams.Text = CStr(.dwOutDatagrams)
        txtListenerTableEntries.Text = CStr(.dwNumAddrs)
    End With
End Sub

