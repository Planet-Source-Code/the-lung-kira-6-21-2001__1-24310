VERSION 5.00
Begin VB.Form frmTCPStatistics 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TCP Statistics"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3975
   Icon            =   "frmTCPStatistics.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   3975
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtCumulativeConnections 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   27
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox txtOutgoingResets 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   25
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox txtIncomingErrors 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox txtSegmentsRetransmitted 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox txtSegmentsReceived 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox txtEstablishedConnections 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Timer timerTCPStatistics 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2280
      Top             =   120
   End
   Begin VB.TextBox txtEstablishedConnectionsReset 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox txtFailedAttempts 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox txtPassiveOpens 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox txtActiveOpens 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox txtMaximumConnections 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox txtMaximumTimeOut 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox txtMinimumTimeOut 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox txtTimeOutAlgorithm 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblCumulativeConnections 
      Caption         =   "Cumulative Connections"
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   3240
      Width           =   2295
   End
   Begin VB.Label lblOutgoingResets 
      Caption         =   "Outgoing Resets"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   3000
      Width           =   2295
   End
   Begin VB.Label lblIncomingErrors 
      Caption         =   "Incoming Errors"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   2760
      Width           =   2295
   End
   Begin VB.Label lblSegmentsRetransmitted 
      Caption         =   "Segments Retransmitted"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   2520
      Width           =   2295
   End
   Begin VB.Label lblSegmentsReceived 
      Caption         =   "Segments Received"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   2280
      Width           =   2295
   End
   Begin VB.Label lblEstablishedConnections 
      Caption         =   "Established Connections"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Label lblEstablishedConnectionsReset 
      Caption         =   "Established Connections Reset"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Label lblFailedAttempts 
      Caption         =   "Failed Attempts"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1560
      Width           =   2295
   End
   Begin VB.Label lblPassiveOpens 
      Caption         =   "Passive Opens"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1320
      Width           =   2295
   End
   Begin VB.Label lblActiveOpens 
      Caption         =   "Active Opens"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1080
      Width           =   2295
   End
   Begin VB.Label lblMaximumConnections 
      Caption         =   "Maximum Connections"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label lblMaximumTimeOut 
      Caption         =   "Maximum Time Out"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label lblMinimumTimeOut 
      Caption         =   "Minimum Time Out"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   2295
   End
   Begin VB.Label lblTimeOutAlgorithm 
      Caption         =   "Time Out Algorithm"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "frmTCPStatistics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    If Function_Exist("iphlpapi.dll", "GetTcpStatistics") = True Then
        timerTCPStatistics_Timer
        timerTCPStatistics.Enabled = True
    Else
        lblTimeOutAlgorithm.Enabled = False
        lblMinimumTimeOut.Enabled = False
        lblMaximumTimeOut.Enabled = False
        lblMaximumConnections.Enabled = False
        lblActiveOpens.Enabled = False
        lblPassiveOpens.Enabled = False
        lblFailedAttempts.Enabled = False
        lblEstablishedConnectionsReset.Enabled = False
        lblEstablishedConnections.Enabled = False
        lblSegmentsReceived.Enabled = False
        lblSegmentsRetransmitted.Enabled = False
        lblIncomingErrors.Enabled = False
        lblOutgoingResets.Enabled = False
        lblCumulativeConnections.Enabled = False
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    timerTCPStatistics.Enabled = False
End Sub

Private Sub timerTCPStatistics_Timer()
    Dim MIB_TCPSTATS As MIB_TCPSTATS
    
    If GetTcpStatistics(MIB_TCPSTATS) <> NO_ERROR Then Failed "GetTcpStatistics"
    
    With MIB_TCPSTATS
        Select Case .dwRtoAlgorithm
            Case MIB_TCP_RTO_OTHER: txtTimeOutAlgorithm.Text = "Other"
            Case MIB_TCP_RTO_CONSTANT: txtTimeOutAlgorithm.Text = "Constant Time-out"
            Case MIB_TCP_RTO_RSRE: txtTimeOutAlgorithm.Text = "MIL-STD-1778 Appendix B"
            Case MIB_TCP_RTO_VANJ: txtTimeOutAlgorithm.Text = "Van Jacobson's Algorithm"
        End Select
        
        txtMinimumTimeOut.Text = CStr(.dwRtoMin)
        txtMaximumTimeOut.Text = CStr(.dwRtoMax)
        txtMaximumConnections.Text = CStr(.dwMaxConn)
        txtActiveOpens.Text = CStr(.dwActiveOpens)
        txtPassiveOpens.Text = CStr(.dwPassiveOpens)
        txtFailedAttempts.Text = CStr(.dwAttemptFails)
        txtEstablishedConnectionsReset.Text = CStr(.dwEstabResets)
        txtEstablishedConnections.Text = CStr(.dwCurrEstab)
        txtSegmentsReceived.Text = CStr(.dwInSegs)
        txtSegmentsRetransmitted.Text = CStr(.dwRetransSegs)
        txtIncomingErrors.Text = CStr(.dwInErrs)
        txtOutgoingResets.Text = CStr(.dwOutRsts)
        txtCumulativeConnections.Text = CStr(.dwNumConns)
    End With
End Sub
