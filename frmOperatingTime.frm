VERSION 5.00
Begin VB.Form frmOperatingTime 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Operating Time"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2655
   Icon            =   "frmOperatingTime.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   2655
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer timerUpTime 
      Interval        =   945
      Left            =   1080
      Top             =   120
   End
   Begin VB.TextBox txtDays 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox txtHours 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox txtMinutes 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox txtSeconds 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1200
      Width           =   1095
   End
   Begin VB.TextBox txtMilliseconds 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox txtUnFormatted 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label lblDays 
      Caption         =   "Days"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label lblHours 
      Caption         =   "Hours"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label lblMinutes 
      Caption         =   "Minutes"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label lblSeconds 
      Caption         =   "Seconds"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label lblMilliseconds 
      Caption         =   "Milliseconds"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label lblUnFormatted 
      Caption         =   "UnFormatted"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2040
      Width           =   1095
   End
End
Attribute VB_Name = "frmOperatingTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    timerUpTime_Timer
End Sub

Private Sub timerUpTime_Timer()
    Dim UpTime As Long
    UpTime = GetTickCount
    
    txtUnFormatted.Text = CStr(UpTime)
    txtMilliseconds.Text = CStr(UpTime - ((UpTime \ 1000) * 1000))
    UpTime = UpTime \ 1000
    txtSeconds.Text = CStr(UpTime - ((UpTime \ 60) * 60))
    UpTime = UpTime \ 60
    txtMinutes.Text = CStr(UpTime - ((UpTime \ 60) * 60))
    UpTime = UpTime \ 60
    txtHours.Text = CStr(UpTime - ((UpTime \ 24) * 24))
    UpTime = UpTime \ 24
    txtDays.Text = CStr(UpTime)
End Sub
