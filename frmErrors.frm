VERSION 5.00
Begin VB.Form frmErrors 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Errors"
   ClientHeight    =   2835
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4215
   Icon            =   "frmErrors.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2835
   ScaleWidth      =   4215
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtErrorNumber 
      Height          =   285
      Left            =   1800
      TabIndex        =   3
      Top             =   480
      Width           =   2295
   End
   Begin VB.CommandButton cmdGetInfo 
      Caption         =   "Get Info"
      Height          =   350
      Left            =   3120
      TabIndex        =   6
      Top             =   2400
      Width           =   975
   End
   Begin VB.TextBox txtDescription 
      Height          =   1095
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   1200
      Width           =   3975
   End
   Begin VB.ComboBox cboErrorType 
      Height          =   315
      Left            =   1800
      TabIndex        =   1
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label lblErrorNumber 
      Caption         =   "Error Number"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label lblDescription 
      Caption         =   "Description"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label lblErrorType 
      Caption         =   "Error Type"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmErrors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdGetInfo_Click()
    If txtErrorNumber.Text = "" Then
        txtErrorNumber.Text = "0"
    End If
    
    
    Dim errDescription As String
    apiError = CLng(txtErrorNumber.Text)
    
    Select Case cboErrorType.ListIndex
        Case 0: Errors apiError, "", errDescription, True
        Case 1: CommDlgError apiError, "", errDescription, True
        Case 2: mciError apiError, "", errDescription, True
        Case 3: PdhError apiError, "", errDescription, True
    End Select
    
    txtDescription.Text = errDescription
End Sub

Private Sub Form_Load()
    With cboErrorType
        .AddItem "Win32"
        .AddItem "Common Dialog"
        .AddItem "MCI"
        .AddItem "PDH"
    End With
    
    txtErrorNumber.Text = GetRegSetting(HKEY_CURRENT_USER, "Software\Kira\Errors", "Number")
    cboErrorType.ListIndex = GetRegSetting(HKEY_CURRENT_USER, "Software\Kira\Errors", "Type")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\Errors", "Number", Val(txtErrorNumber.Text), REG_DWORD
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\Errors", "Type", cboErrorType.ListIndex, REG_DWORD
End Sub

Private Sub txtErrorNumber_Change()
    txtErrorNumber.Text = CStr(Val(Rem_NonNumeric_Chr(txtErrorNumber.Text)))
    If Val(txtErrorNumber.Text) < -2147483648# Then txtErrorNumber.Text = "-2147483648"
    If Val(txtErrorNumber.Text) > 2147483647 Then txtErrorNumber.Text = "2147483647"
End Sub
