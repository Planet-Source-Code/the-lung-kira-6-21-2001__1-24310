VERSION 5.00
Begin VB.Form frmICMPStatistics 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ICMP Statistics"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6975
   Icon            =   "frmICMPStatistics.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   6975
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtOutAddressMaskReplies 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   53
      Top             =   3360
      Width           =   1095
   End
   Begin VB.TextBox txtOutAddressMaskRequests 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   51
      Top             =   3120
      Width           =   1095
   End
   Begin VB.TextBox txtInAddressMaskReplies 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   26
      Top             =   3360
      Width           =   1095
   End
   Begin VB.TextBox txtInAddressMaskRequests 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   24
      Top             =   3120
      Width           =   1095
   End
   Begin VB.TextBox txtInTimeStampReplies 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   2880
      Width           =   1095
   End
   Begin VB.TextBox txtOutTimeStampReplies 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   49
      Top             =   2880
      Width           =   1095
   End
   Begin VB.TextBox txtOutTimeStampRequests 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   47
      Top             =   2640
      Width           =   1095
   End
   Begin VB.TextBox txtOutEchoReplies 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   45
      Top             =   2400
      Width           =   1095
   End
   Begin VB.TextBox txtInTimeStampRequests 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   2640
      Width           =   1095
   End
   Begin VB.TextBox txtInEchoReplies 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   2400
      Width           =   1095
   End
   Begin VB.TextBox txtOutEchoRequests 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   43
      Top             =   2160
      Width           =   1095
   End
   Begin VB.TextBox txtOutRedirection 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   41
      Top             =   1920
      Width           =   1095
   End
   Begin VB.TextBox txtOutSourceQuench 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   39
      Top             =   1680
      Width           =   1095
   End
   Begin VB.TextBox txtInEchoRequests 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   2160
      Width           =   1095
   End
   Begin VB.TextBox txtInRedirection 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   1920
      Width           =   1095
   End
   Begin VB.TextBox txtInSourceQuench 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   1680
      Width           =   1095
   End
   Begin VB.TextBox txtOutParameterProblem 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   37
      Top             =   1440
      Width           =   1095
   End
   Begin VB.TextBox txtInParameterProblem 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   1440
      Width           =   1095
   End
   Begin VB.TextBox txtInTTLExceeded 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   1200
      Width           =   1095
   End
   Begin VB.TextBox txtInDestinationUnreachable 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox txtOutTTLExceeded 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   35
      Top             =   1200
      Width           =   1095
   End
   Begin VB.TextBox txtOutDestinationUnreachable 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   33
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox txtInErrors 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   720
      Width           =   1095
   End
   Begin VB.TextBox txtOutErrors 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   31
      Top             =   720
      Width           =   1095
   End
   Begin VB.Timer timerICMPStatistics 
      Enabled         =   0   'False
      Interval        =   945
      Left            =   3000
      Top             =   0
   End
   Begin VB.TextBox txtOutMessages 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   29
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox txtInMessages 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label lblOutTimeStampRequests 
      Caption         =   "Time-Stamp Requests"
      Height          =   255
      Left            =   3600
      TabIndex        =   46
      Top             =   2640
      Width           =   1935
   End
   Begin VB.Label lblOutTimeStampReplies 
      Caption         =   "Time-Stamp Replies"
      Height          =   255
      Left            =   3600
      TabIndex        =   48
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Label lblOutTTLExceeded 
      Caption         =   "TTL Exceeded"
      Height          =   255
      Left            =   3600
      TabIndex        =   34
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label lblOutSourceQuench 
      Caption         =   "Source Quench"
      Height          =   255
      Left            =   3600
      TabIndex        =   38
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Label lblOutRedirection 
      Caption         =   "Redirection"
      Height          =   255
      Left            =   3600
      TabIndex        =   40
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Label lblOutParameterProblem 
      Caption         =   "Parameter Problem"
      Height          =   255
      Left            =   3600
      TabIndex        =   36
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label lblOutMessages 
      Caption         =   "Messages"
      Height          =   255
      Left            =   3600
      TabIndex        =   28
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label lblOutErrors 
      Caption         =   "Errors"
      Height          =   255
      Left            =   3600
      TabIndex        =   30
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label lblOutEchoRequests 
      Caption         =   "Echo Requests"
      Height          =   255
      Left            =   3600
      TabIndex        =   42
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Label lblOutEchoReplies 
      Caption         =   "Echo Replies"
      Height          =   255
      Left            =   3600
      TabIndex        =   44
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Label lblOutDestinationUnreachable 
      Caption         =   "Destination Unreachable"
      Height          =   255
      Left            =   3600
      TabIndex        =   32
      Top             =   960
      Width           =   1935
   End
   Begin VB.Label lblOutAddressMaskRequests 
      Caption         =   "Address Mask Requests"
      Height          =   255
      Left            =   3600
      TabIndex        =   50
      Top             =   3120
      Width           =   1935
   End
   Begin VB.Label lblOutAddressMaskReplies 
      Caption         =   "Address Mask Replies"
      Height          =   255
      Left            =   3600
      TabIndex        =   52
      Top             =   3360
      Width           =   1935
   End
   Begin VB.Label lblInTimeStampRequests 
      Caption         =   "Time-Stamp Requests"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   2640
      Width           =   1935
   End
   Begin VB.Label lblInTimeStampReplies 
      Caption         =   "Time-Stamp Replies"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Label lblInTTLExceeded 
      Caption         =   "TTL Exceeded"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label lblInSourceQuench 
      Caption         =   "Source Quench"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Label lblInRedirection 
      Caption         =   "Redirection"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Label lblInParameterProblem 
      Caption         =   "Parameter Problem"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label lblInMessages 
      Caption         =   "Messages"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label lblInErrors 
      Caption         =   "Errors"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label lblInEchoRequests 
      Caption         =   "Echo Requests"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Label lblInEchoReplies 
      Caption         =   "Echo Replies"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Label lblInDestinationUnreachable 
      Caption         =   "Destination Unreachable"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   1935
   End
   Begin VB.Label lblInAddressMaskRequests 
      Caption         =   "Address Mask Requests"
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   3120
      Width           =   1935
   End
   Begin VB.Label lblInAddressMaskReplies 
      Caption         =   "Address Mask Replies"
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   3360
      Width           =   1935
   End
   Begin VB.Label lblOut 
      Caption         =   "Out"
      Height          =   255
      Left            =   3600
      TabIndex        =   27
      Top             =   120
      Width           =   375
   End
   Begin VB.Label lblIn 
      Caption         =   "In"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   375
   End
End
Attribute VB_Name = "frmICMPStatistics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    If Function_Exist("iphlpapi.dll", "GetUdpStatistics") = True Then
        timerICMPStatistics_Timer
        timerICMPStatistics.Enabled = True
    Else
        lblInMessages.Enabled = False
        lblInErrors.Enabled = False
        lblInDestinationUnreachable.Enabled = False
        lblInTTLExceeded.Enabled = False
        lblInParameterProblem.Enabled = False
        lblInSourceQuench.Enabled = False
        lblInRedirection.Enabled = False
        lblInEchoRequests.Enabled = False
        lblInEchoReplies.Enabled = False
        lblInTimeStampRequests.Enabled = False
        lblInTimeStampReplies.Enabled = False
        lblInAddressMaskRequests.Enabled = False
        lblInAddressMaskReplies.Enabled = False
        lblOutMessages.Enabled = False
        lblOutErrors.Enabled = False
        lblOutDestinationUnreachable.Enabled = False
        lblOutTTLExceeded.Enabled = False
        lblOutParameterProblem.Enabled = False
        lblOutSourceQuench.Enabled = False
        lblOutRedirection.Enabled = False
        lblOutEchoRequests.Enabled = False
        lblOutEchoReplies.Enabled = False
        lblOutTimeStampRequests.Enabled = False
        lblOutTimeStampReplies.Enabled = False
        lblOutAddressMaskRequests.Enabled = False
        lblOutAddressMaskReplies.Enabled = False
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    timerICMPStatistics.Enabled = False
End Sub

Private Sub timerICMPStatistics_Timer()
    Dim MIB_ICMP As MIB_ICMP
    
    If GetIcmpStatistics(MIB_ICMP) <> NO_ERROR Then Failed "GetIcmpStatistics"
    
    
    With MIB_ICMP.stats.icmpInStats
        txtInMessages.Text = CStr(.dwMsgs)
        txtInErrors.Text = CStr(.dwErrors)
        txtInDestinationUnreachable.Text = CStr(.dwDestUnreachs)
        txtInTTLExceeded.Text = CStr(.dwTimeExcds)
        txtInParameterProblem.Text = CStr(.dwParmProbs)
        txtInSourceQuench.Text = CStr(.dwSrcQuenchs)
        txtInRedirection.Text = CStr(.dwRedirects)
        txtInEchoRequests.Text = CStr(.dwEchos)
        txtInEchoReplies.Text = CStr(.dwEchoReps)
        txtInTimeStampRequests.Text = CStr(.dwTimestamps)
        txtInTimeStampReplies.Text = CStr(.dwTimestampReps)
        txtInAddressMaskRequests.Text = CStr(.dwAddrMasks)
        txtInAddressMaskReplies.Text = CStr(.dwAddrMaskReps)
    End With
    
    With MIB_ICMP.stats.icmpOutStats
        txtOutMessages.Text = CStr(.dwMsgs)
        txtOutErrors.Text = CStr(.dwErrors)
        txtOutDestinationUnreachable.Text = CStr(.dwDestUnreachs)
        txtOutTTLExceeded.Text = CStr(.dwTimeExcds)
        txtOutParameterProblem.Text = CStr(.dwParmProbs)
        txtOutSourceQuench.Text = CStr(.dwSrcQuenchs)
        txtOutRedirection.Text = CStr(.dwRedirects)
        txtOutEchoRequests.Text = CStr(.dwEchos)
        txtOutEchoReplies.Text = CStr(.dwEchoReps)
        txtOutTimeStampRequests.Text = CStr(.dwTimestamps)
        txtOutTimeStampReplies.Text = CStr(.dwTimestampReps)
        txtOutAddressMaskRequests.Text = CStr(.dwAddrMasks)
        txtOutAddressMaskReplies.Text = CStr(.dwAddrMaskReps)
    End With
End Sub
