VERSION 5.00
Begin VB.Form frmIconMetrics 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Icon Metrics"
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2775
   Icon            =   "frmIconMetrics.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   2775
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtSpacingHeight 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   1560
      Width           =   855
   End
   Begin VB.TextBox txtSmallHeight 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   960
      Width           =   855
   End
   Begin VB.TextBox txtSmallWidth 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   720
      Width           =   855
   End
   Begin VB.TextBox txtSpacingWidth 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   1320
      Width           =   855
   End
   Begin VB.TextBox txtDHeight 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   360
      Width           =   855
   End
   Begin VB.TextBox txtDWidth 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.Label lblSmallHeight 
      Caption         =   "Small Icon Height"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label lblSmallWidth 
      Caption         =   "Small Icon Width"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label lblSpacingWidth 
      Caption         =   "Spacing Width"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label lblSpacingHeight 
      Caption         =   "Spacing Height"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label lblDHeight 
      Caption         =   "Default Height"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label lblDWidth 
      Caption         =   "Default Width"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmIconMetrics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    txtDWidth.Text = CStr(GetSystemMetrics(SM_CXICON))
    txtDHeight.Text = CStr(GetSystemMetrics(SM_CYICON))
    txtSmallWidth.Text = CStr(GetSystemMetrics(SM_CXSMICON))
    txtSmallHeight.Text = CStr(GetSystemMetrics(SM_CYSMICON))
    txtSpacingWidth.Text = CStr(GetSystemMetrics(SM_CXICONSPACING))
    txtSpacingHeight.Text = CStr(GetSystemMetrics(SM_CYICONSPACING))
End Sub
