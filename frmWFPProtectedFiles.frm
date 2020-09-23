VERSION 5.00
Begin VB.Form frmWFPProtectedFiles 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Windows File Protection - Protected Files"
   ClientHeight    =   3555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9255
   Icon            =   "frmWFPProtectedFiles.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   9255
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   350
      Left            =   6960
      TabIndex        =   3
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   350
      Left            =   8160
      TabIndex        =   4
      Top             =   3120
      Width           =   975
   End
   Begin VB.ListBox lstProtectedFiles 
      Height          =   2400
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   360
      Width           =   9015
   End
   Begin VB.TextBox txtProtectedFiles 
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   2760
      Width           =   9015
   End
   Begin VB.Label lblProtectedFiles 
      Caption         =   "Protected Files"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "frmWFPProtectedFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdRefresh_Click()
    Dim PROTECTED_FILE_DATA As PROTECTED_FILE_DATA
    
    lstProtectedFiles.Clear
    
    PROTECTED_FILE_DATA.FileNumber = 0
    If SfcGetNextProtectedFile(&H0, PROTECTED_FILE_DATA) = 0 Then Failed "SfcGetNextProtectedFile"
    
    lstProtectedFiles.AddItem UnicodeToAscii(PROTECTED_FILE_DATA.FileName, &H0)
    
    
    With PROTECTED_FILE_DATA
        Do
            PROTECTED_FILE_DATA.FileNumber = .FileNumber + 1
            If SfcGetNextProtectedFile(&H0, PROTECTED_FILE_DATA) = 0 Then
                Failed "SfcGetNextProtectedFile"
                Exit Do
            End If
            
            lstProtectedFiles.AddItem UnicodeToAscii(.FileName, &H0)
        Loop
    End With
End Sub

Private Sub cmdSave_Click()
    Dim strFileName As String
    Dim lngIncrement As Long
    Dim strOutput As String
    
    strFileName = GetSaveName(frmWFPProtectedFiles.hwnd, "All Files (*.*)" & Chr$(0) & "*.*" & Chr$(0), 2, "Save", OFN_EXPLORER Or OFN_HIDEREADONLY Or OFN_OVERWRITEPROMPT Or OFN_DONTADDTORECENT)
    If strFileName = "" Then Exit Sub
    
    With lstProtectedFiles
        For lngIncrement = 0 To .ListCount - 2
            strOutput = strOutput & .List(lngIncrement) & Chr$(13) & Chr$(10)
        Next lngIncrement
        
        strOutput = strOutput & .List(lngIncrement)
    End With
    
    WriteFile_String strFileName, strOutput, 0, CREATE_ALWAYS
End Sub

Private Sub Form_Load()
    If Function_Exist("sfc.dll", "SfcGetNextProtectedFile") = True Then
        cmdRefresh_Click
    Else
        lblProtectedFiles.Enabled = False
        lstProtectedFiles.Enabled = False
        txtProtectedFiles.Enabled = False
        cmdRefresh.Enabled = False
        cmdSave.Enabled = False
    End If
End Sub

Private Sub lstProtectedFiles_Click()
    txtProtectedFiles.Text = lstProtectedFiles.List(lstProtectedFiles.ListIndex)
End Sub
