VERSION 5.00
Begin VB.Form frmMouseMonitor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mouse Monitor"
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5295
   Icon            =   "frmMouseMonitor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   5295
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtX2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   1440
      Width           =   1455
   End
   Begin VB.TextBox txtX1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   1200
      Width           =   1455
   End
   Begin VB.TextBox txtRight 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   960
      Width           =   1455
   End
   Begin VB.TextBox txtMiddle 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   720
      Width           =   1455
   End
   Begin VB.TextBox txtLeft 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   480
      Width           =   1455
   End
   Begin VB.TextBox txtWheel 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   1560
      Width           =   1455
   End
   Begin VB.TextBox txtY 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   720
      Width           =   1455
   End
   Begin VB.TextBox txtX 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   480
      Width           =   1455
   End
   Begin VB.TextBox txtTotalClicks 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   1800
      Width           =   1455
   End
   Begin VB.TextBox txtTotalMovement 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label lblWheel 
      Caption         =   "Wheel"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label lblX1 
      Caption         =   "X1"
      Height          =   255
      Left            =   2760
      TabIndex        =   16
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label lblX2 
      Caption         =   "X2"
      Height          =   255
      Left            =   2760
      TabIndex        =   18
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label lblTotalClicks 
      Caption         =   "Total"
      Height          =   255
      Left            =   2760
      TabIndex        =   20
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label lblRight 
      Caption         =   "Right"
      Height          =   255
      Left            =   2760
      TabIndex        =   14
      Top             =   960
      Width           =   735
   End
   Begin VB.Label lblLeft 
      Caption         =   "Left"
      Height          =   255
      Left            =   2760
      TabIndex        =   10
      Top             =   480
      Width           =   735
   End
   Begin VB.Label lblMiddle 
      Caption         =   "Middle"
      Height          =   255
      Left            =   2760
      TabIndex        =   12
      Top             =   720
      Width           =   735
   End
   Begin VB.Label lblClicks 
      Caption         =   "Clicks"
      Height          =   255
      Left            =   2760
      TabIndex        =   9
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label lblMovement 
      Caption         =   "Movement"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label lbY 
      Caption         =   "Y"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   735
   End
   Begin VB.Label lblX 
      Caption         =   "X"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   735
   End
   Begin VB.Label lblTotalMovement 
      Caption         =   "Total"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   735
   End
End
Attribute VB_Name = "frmMouseMonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    With MouseMonitor
        txtX.Text = CStr(.TotalXMovement)
        txtY.Text = CStr(.TotalYMovement)
        txtWheel.Text = CStr(.TotalWheelMovement)
        txtTotalMovement.Text = CStr(.TotalXMovement + .TotalYMovement)
        
        txtLeft.Text = CStr(.TotalLClicks)
        txtMiddle.Text = CStr(.TotalMClicks)
        txtRight.Text = CStr(.TotalRClicks)
        txtX1.Text = CStr(.TotalX1Clicks)
        txtX2.Text = CStr(.TotalX2Clicks)
        txtTotalClicks.Text = CStr(.TotalLClicks + .TotalMClicks + .TotalRClicks + .TotalX1Clicks + .TotalX2Clicks)
    End With
    
    'InchesX = (MouseMovX * (Screen.TwipsPerPixelX / 20)) / 72
End Sub
