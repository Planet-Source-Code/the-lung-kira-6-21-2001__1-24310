VERSION 5.00
Begin VB.Form frmCachedPasswords 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cached Passwords"
   ClientHeight    =   2610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3975
   Icon            =   "frmCachedPasswords.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2610
   ScaleWidth      =   3975
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstCachedPasswords 
      Height          =   2010
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   3975
   End
   Begin VB.CommandButton cmdGetData 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Get Data"
      Enabled         =   0   'False
      Height          =   350
      Left            =   2880
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   2160
      Width           =   975
   End
End
Attribute VB_Name = "frmCachedPasswords"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdGetData_Click()
    lstCachedPasswords.Clear
    If WNetEnumCachedPasswords("", 0, &HFF, AddressOf EnumCachedPasswordsProc, 0) <> 0 Then Failed "WNetEnumCachedPasswords"
End Sub

Private Sub Form_Load()
    If Function_Exist("mpr.dll", "WNetEnumCachedPasswords") = True Then
        cmdGetData_Click
    Else
        lstCachedPasswords.Enabled = False
        cmdGetData.Enabled = False
    End If
End Sub
