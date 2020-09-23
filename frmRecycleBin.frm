VERSION 5.00
Begin VB.Form frmRecycleBin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recycle Bin"
   ClientHeight    =   4305
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4335
   Icon            =   "frmRecycleBin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   4335
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkRename 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   3960
      TabIndex        =   16
      Top             =   3000
      Width           =   255
   End
   Begin VB.CheckBox chkDelete 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   3960
      TabIndex        =   18
      Top             =   3240
      Width           =   255
   End
   Begin VB.CheckBox chkProperties 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   3960
      TabIndex        =   20
      Top             =   3480
      Width           =   255
   End
   Begin VB.TextBox txtDisplayName 
      Height          =   285
      Left            =   1560
      TabIndex        =   13
      Top             =   2280
      Width           =   2655
   End
   Begin VB.TextBox txtInfoTip 
      Height          =   285
      Left            =   1560
      TabIndex        =   11
      Top             =   1920
      Width           =   2655
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   350
      Left            =   3240
      TabIndex        =   21
      Top             =   3840
      Width           =   975
   End
   Begin VB.CommandButton cmdEmpty 
      Caption         =   "Empty"
      Height          =   350
      Left            =   3240
      TabIndex        =   8
      Top             =   960
      Width           =   975
   End
   Begin VB.ComboBox cboDrive 
      Height          =   315
      Left            =   1560
      TabIndex        =   7
      Top             =   960
      Width           =   1575
   End
   Begin VB.CheckBox chkSound 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   3960
      TabIndex        =   5
      Top             =   600
      Width           =   255
   End
   Begin VB.CheckBox chkProgressUI 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   3960
      TabIndex        =   3
      Top             =   360
      Width           =   255
   End
   Begin VB.CheckBox chkConfirmation 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   3960
      TabIndex        =   1
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton cmdEmptyAll 
      Caption         =   "Empty All"
      Height          =   350
      Left            =   3240
      TabIndex        =   9
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label lblDesktopIcon 
      Caption         =   "Desktop Icon"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label lblRename 
      Caption         =   "Rename"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label lblDelete 
      Caption         =   "Delete"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label lblProperties 
      Caption         =   "Properties"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label lblDisplayName 
      Caption         =   "Display Name"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label lblInfoTip 
      Caption         =   "Info Tip"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label lblDrive 
      Caption         =   "Drive"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label lblSound 
      Caption         =   "Sound"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label lblProgressUI 
      Caption         =   "Progress UI"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label lblConfirmation 
      Caption         =   "Confirmation"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmRecycleBin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdApply_Click()
    SaveRegSetting HKEY_CLASSES_ROOT, "CLSID\{645FF040-5081-101B-9F08-00AA002F954E}", "", txtDisplayName.Text, REG_SZ
    SaveRegSetting HKEY_CLASSES_ROOT, "CLSID\{645FF040-5081-101B-9F08-00AA002F954E}", "InfoTip", txtInfoTip.Text, REG_SZ
    
    
    Dim lngValue As Long
    Dim strInput As String
    
    If chkRename.value = 1 Then lngValue = lngValue + 10
    If chkDelete.value = 1 Then lngValue = lngValue + 20
    If chkProperties.value = 1 Then lngValue = lngValue + 40
    
    strInput = GetRegSetting(HKEY_CLASSES_ROOT, "CLSID\{645FF040-5081-101B-9F08-00AA002F954E}\ShellFolder", "Attributes")
    If Len(strInput) >= 1 Then
        strInput = Chr$(strtoul_(lngValue, 16)) & Right$(strInput, Len(strInput) - 1)
        SaveRegSetting HKEY_CLASSES_ROOT, "CLSID\{645FF040-5081-101B-9F08-00AA002F954E}\ShellFolder", "Attributes", strInput, REG_BINARY
    End If
End Sub

Private Sub cmdEmpty_Click()
    If cboDrive.ListIndex = -1 Then Exit Sub
    
    Dim lngFlags As Long
    
    If chkConfirmation.value = 0 Then lngFlags = lngFlags Or SHERB_NOCONFIRMATION
    If chkProgressUI.value = 0 Then lngFlags = lngFlags Or SHERB_NOPROGRESSUI
    If chkSound.value = 0 Then lngFlags = lngFlags Or SHERB_NOSOUND
    
    If SHEmptyRecycleBin(&H0, cboDrive.List(cboDrive.ListIndex), lngFlags) <> S_OK Then Failed "SHEmptyRecycleBin"
End Sub

Private Sub cmdEmptyAll_Click()
    Dim lngFlags As Long
    
    If chkConfirmation.value = 0 Then lngFlags = lngFlags Or SHERB_NOCONFIRMATION
    If chkProgressUI.value = 0 Then lngFlags = lngFlags Or SHERB_NOPROGRESSUI
    If chkSound.value = 0 Then lngFlags = lngFlags Or SHERB_NOSOUND
    
    If SHEmptyRecycleBin(&H0, "", lngFlags) <> S_OK Then Failed "SHEmptyRecycleBin"
End Sub

Private Sub Form_Load()
    Dim strDrives As String
    Dim lngIncrement As Long
    
    strDrives = Left$(StrReverse(ltoa_(GetLogicalDrives, 2)) & String$(32, "0"), 32)
    
    With cboDrive
        For lngIncrement = 1 To Len(strDrives)
            If Mid$(strDrives, lngIncrement, 1) = "1" Then
                .AddItem Chr$(&H40 + lngIncrement) & ":\"
            End If
        Next lngIncrement
    End With
    
    
    chkConfirmation.value = GetRegSetting(HKEY_CURRENT_USER, "Software\Kira\RecycleBin", "Confirmation")
    chkProgressUI.value = GetRegSetting(HKEY_CURRENT_USER, "Software\Kira\RecycleBin", "ProgressUI")
    chkSound.value = GetRegSetting(HKEY_CURRENT_USER, "Software\Kira\RecycleBin", "Sound")
    
    txtDisplayName.Text = GetRegSetting(HKEY_CLASSES_ROOT, "CLSID\{645FF040-5081-101B-9F08-00AA002F954E}", "")
    txtInfoTip.Text = GetRegSetting(HKEY_CLASSES_ROOT, "CLSID\{645FF040-5081-101B-9F08-00AA002F954E}", "InfoTip")
    
    
    Dim strReturn As String
    strReturn = GetRegSetting(HKEY_CLASSES_ROOT, "CLSID\{645FF040-5081-101B-9F08-00AA002F954E}\ShellFolder", "Attributes")
    strReturn = Right$("00" & ltoa_(Asc(Mid$(strReturn, 1, 1)), 16), 2) & _
                Right$("00" & ltoa_(Asc(Mid$(strReturn, 2, 1)), 16), 2) & _
                Right$("00" & ltoa_(Asc(Mid$(strReturn, 3, 1)), 16), 2) & _
                Right$("00" & ltoa_(Asc(Mid$(strReturn, 4, 1)), 16), 2)
    
    Select Case strReturn
        Case "10010020"
            chkRename.value = 1
        Case "20010020"
            chkDelete.value = 1
        Case "30010020"
            chkRename.value = 1
            chkDelete.value = 1
        Case "40010020"
            chkProperties.value = 1
        Case "50010020"
            chkRename.value = 1
            chkProperties.value = 1
        Case "60010020"
            chkDelete.value = 1
            chkProperties.value = 1
        Case "70010020"
            chkRename.value = 1
            chkDelete.value = 1
            chkProperties.value = 1
    End Select
    
    
    If Function_Exist("shell32.dll", "SHEmptyRecycleBinA") = False Then
        cmdEmpty.Enabled = False
        cmdEmptyAll.Enabled = False
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\RecycleBin", "Confirmation", chkConfirmation.value, REG_DWORD
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\RecycleBin", "ProgressUI", chkProgressUI.value, REG_DWORD
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\RecycleBin", "Sound", chkSound.value, REG_DWORD
End Sub
