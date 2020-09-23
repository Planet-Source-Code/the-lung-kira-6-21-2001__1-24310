VERSION 5.00
Begin VB.Form frmAccessTimeout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Access Timeout"
   ClientHeight    =   1530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3135
   Icon            =   "frmAccessTimeout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1530
   ScaleWidth      =   3135
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   350
      Left            =   2040
      TabIndex        =   6
      Top             =   1080
      Width           =   975
   End
   Begin VB.CheckBox chkTimeOutOn 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   2760
      TabIndex        =   3
      Top             =   360
      Width           =   255
   End
   Begin VB.CheckBox chkOnOffFeedback 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   2760
      TabIndex        =   1
      Top             =   120
      Width           =   255
   End
   Begin VB.TextBox txtTimeOut 
      Height          =   285
      Left            =   1680
      TabIndex        =   5
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label lblTimeOutOn 
      Caption         =   "Time Out On"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label lblOnOffFeedback 
      Caption         =   "On Off Feedback"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label lblTimeOut 
      Caption         =   "Time Out"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   1335
   End
End
Attribute VB_Name = "frmAccessTimeout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdApply_Click()
    Dim ACCESSTIMEOUT As ACCESSTIMEOUT
    ACCESSTIMEOUT.cbSize = Len(ACCESSTIMEOUT)
        
    Dim ONOFFFEEDBACK As Long
    Dim TIMEOUTON As Long
    
    If chkOnOffFeedback.value = 1 Then ONOFFFEEDBACK = ATF_ONOFFFEEDBACK
    If chkTimeOutOn.value = 1 Then TIMEOUTON = ATF_TIMEOUTON
    
    ACCESSTIMEOUT.dwFlags = ONOFFFEEDBACK Or TIMEOUTON
    ACCESSTIMEOUT.iTimeOutMSec = Val(txtTimeOut.Text)
    
    If SystemParametersInfo(SPI_SETACCESSTIMEOUT, ACCESSTIMEOUT.cbSize, ACCESSTIMEOUT, SPIF_UPDATEINIFILE) = False Then Failed "SystemParametersInfo"
End Sub

Private Sub Form_Load()
    Dim ACCESSTIMEOUT As ACCESSTIMEOUT
    ACCESSTIMEOUT.cbSize = Len(ACCESSTIMEOUT)
    
    If SystemParametersInfo(SPI_GETACCESSTIMEOUT, ACCESSTIMEOUT.cbSize, ACCESSTIMEOUT, 0) = False Then Failed "SystemParametersInfo"
    
    With ACCESSTIMEOUT
        If .dwFlags And ATF_ONOFFFEEDBACK Then chkOnOffFeedback.value = 1
        If .dwFlags And ATF_TIMEOUTON Then chkTimeOutOn.value = 1
        
        txtTimeOut.Text = CStr(.iTimeOutMSec)
    End With
End Sub

Private Sub txtTimeOut_Change()
    txtTimeOut.Text = CStr(Val(Rem_NonNumeric_Chr(txtTimeOut.Text)))
    If Val(txtTimeOut.Text) < 0 Then txtTimeOut.Text = "0"
    If Val(txtTimeOut.Text) > 2147483647 Then txtTimeOut.Text = "2147483647"
End Sub
