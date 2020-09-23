VERSION 5.00
Begin VB.Form frmExtra 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5295
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   2475
   ControlBox      =   0   'False
   Icon            =   "frmExtra.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmExtra.frx":000C
   ScaleHeight     =   5295
   ScaleWidth      =   2475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblKira 
      BackStyle       =   0  'Transparent
      Caption         =   "For my love, Kira."
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   4920
      Width           =   2175
   End
End
Attribute VB_Name = "frmExtra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Click()
    Unload frmExtra
End Sub

Private Sub lblKira_Click()
    Unload frmExtra
End Sub
