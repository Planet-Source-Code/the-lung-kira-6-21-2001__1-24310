VERSION 5.00
Begin VB.Form frmMouseInfo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mouse Info"
   ClientHeight    =   2655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2775
   Icon            =   "frmMouseInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   2775
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkMousePresent 
      Enabled         =   0   'False
      Height          =   255
      Left            =   2400
      TabIndex        =   13
      Top             =   2040
      Width           =   255
   End
   Begin VB.CheckBox chkMouseWheel 
      Enabled         =   0   'False
      Height          =   255
      Left            =   2400
      TabIndex        =   15
      Top             =   2280
      Width           =   255
   End
   Begin VB.TextBox txtDragDropWidth 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   1320
      Width           =   735
   End
   Begin VB.TextBox txtDragDropHeight 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1080
      Width           =   735
   End
   Begin VB.TextBox txtButtons 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox txtSwapButton 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   1680
      Width           =   735
   End
   Begin VB.TextBox txtCursorWidth 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   720
      Width           =   735
   End
   Begin VB.TextBox txtCursorHeight 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   480
      Width           =   735
   End
   Begin VB.Label lblMousePresent 
      Caption         =   "Mouse Present"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label lblDragDropWidth 
      Caption         =   "Drag Drop Width"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label lblDragDropHeight 
      Caption         =   "Drag Drop Height"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label lblButtons 
      Caption         =   "Buttons"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label lblMouseWheel 
      Caption         =   "Mouse Wheel"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label lblSwapButton 
      Caption         =   "Main Button"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label lblCursorWidth 
      Caption         =   "Cursor Width"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label lblCursorHeight 
      Caption         =   "Cursor Height"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1575
   End
End
Attribute VB_Name = "frmMouseInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    txtButtons.Text = CStr(GetSystemMetrics(SM_CMOUSEBUTTONS))
    txtCursorHeight.Text = CStr(GetSystemMetrics(SM_CYCURSOR))
    txtCursorWidth.Text = CStr(GetSystemMetrics(SM_CXCURSOR))
    txtDragDropHeight.Text = CStr(GetSystemMetrics(SM_CXDRAG))
    txtDragDropWidth.Text = CStr(GetSystemMetrics(SM_CYDRAG))
    chkMousePresent.value = GetSystemMetrics(SM_MOUSEPRESENT)
    
    If WinVersion(4010000, 0, True) = True Then
        chkMouseWheel.value = CStr(GetSystemMetrics(SM_MOUSEWHEELPRESENT))
    End If
    
    If GetSystemMetrics(SM_SWAPBUTTON) = True Then
        txtSwapButton.Text = "Right"
    Else
        txtSwapButton.Text = "Left"
    End If
End Sub
