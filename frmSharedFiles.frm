VERSION 5.00
Begin VB.Form frmSharedFiles 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Shared Files"
   ClientHeight    =   3210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7455
   Icon            =   "frmSharedFiles.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   7455
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   350
      Left            =   5160
      TabIndex        =   5
      Top             =   2760
      Width           =   975
   End
   Begin VB.CheckBox chkExists 
      Height          =   255
      Left            =   2040
      TabIndex        =   4
      Top             =   2760
      Width           =   255
   End
   Begin VB.TextBox txtLocation 
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   2400
      Width           =   7215
   End
   Begin VB.ListBox lstLocation 
      Height          =   2010
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   7215
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Del Entry"
      Height          =   350
      Left            =   6360
      TabIndex        =   6
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label lblSharedFiles 
      Caption         =   "Shared Files"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label lblExists 
      Caption         =   "Exists"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2760
      Width           =   1575
   End
End
Attribute VB_Name = "frmSharedFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDelete_Click()
    If lstLocation.List(lstLocation.ListIndex) <> "" Then
        DeleteValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\SharedDLLs", lstLocation.List(lstLocation.ListIndex)
    End If
End Sub

Private Sub cmdSave_Click()
    Dim strFilename As String
    Dim lngIncrement As Long
    Dim strOutput As String
    
    strFilename = GetSaveName(frmSharedFiles.hwnd, "All Files (*.*)" & Chr$(0) & "*.*" & Chr$(0), 2, "Save", OFN_EXPLORER Or OFN_HIDEREADONLY Or OFN_OVERWRITEPROMPT Or OFN_DONTADDTORECENT)
    If strFilename = "" Then Exit Sub
    
    With lstLocation
        For lngIncrement = 0 To .ListCount - 2
            strOutput = strOutput & .List(lngIncrement) & Chr$(13) & Chr$(10)
        Next lngIncrement
        
        strOutput = strOutput & .List(lngIncrement)
    End With
    
    WriteFile_String strFilename, strOutput, 0, CREATE_ALWAYS
End Sub

Private Sub Form_Load()
    Dim strValueName() As String
    Dim strData() As String
    Dim lngDataType() As Long
    Dim lngCount As Long
    
    EnumValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\SharedDLLs", strValueName(), strData(), lngDataType(), lngCount
    
    Dim lngIncrement As Long
    For lngIncrement = 0 To lngCount - 1
        If Trim$(strValueName(lngIncrement)) <> "" Then
            lstLocation.AddItem strValueName(lngIncrement)
        End If
    Next lngIncrement
End Sub

Private Sub lstLocation_Click()
    txtLocation.Text = lstLocation.List(lstLocation.ListIndex)
    
    If File_Exist(txtLocation.Text) = True Then
        chkExists.value = 1
    Else
        chkExists.value = 0
    End If
End Sub
