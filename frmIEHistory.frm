VERSION 5.00
Begin VB.Form frmIEHistory 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "IE History"
   ClientHeight    =   1155
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4575
   Icon            =   "frmIEHistory.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1155
   ScaleWidth      =   4575
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   350
      Left            =   2280
      TabIndex        =   2
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   350
      Left            =   3480
      TabIndex        =   3
      Top             =   720
      Width           =   975
   End
   Begin VB.ComboBox cboTypedURLs 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   4335
   End
   Begin VB.Label lblTypedURLs 
      Caption         =   "Type URLs"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmIEHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClear_Click()
    Dim strValueName() As String
    Dim strData() As String
    Dim lngDataType() As Long
    Dim lngCount As Long
    
    EnumValue HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\TypedURLs", strValueName(), strData(), lngDataType(), lngCount
    
    
    Dim lngIncrement As Long
    
    For lngIncrement = 0 To lngCount - 1
        DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\TypedURLs", strValueName(lngIncrement)
    Next lngIncrement
End Sub

Private Sub cmdSave_Click()
    Dim strFileName As String
    Dim lngIncrement As Long
    Dim strOutput As String
    
    strFileName = GetSaveName(frmIEHistory.hwnd, "All Files (*.*)" & Chr$(0) & "*.*" & Chr$(0), 2, "Save", OFN_EXPLORER Or OFN_HIDEREADONLY Or OFN_OVERWRITEPROMPT Or OFN_DONTADDTORECENT)
    If strFileName = "" Then Exit Sub
    
    With cboTypedURLs
        For lngIncrement = 0 To .ListCount - 2
            strOutput = strOutput & .List(lngIncrement) & Chr$(13) & Chr$(10)
        Next lngIncrement
        
        strOutput = strOutput & .List(lngIncrement)
    End With
    
    WriteFile_String strFileName, strOutput, 0, CREATE_ALWAYS
End Sub

Private Sub Form_Load()
    Dim strValueName() As String
    Dim strData() As String
    Dim lngDataType() As Long
    Dim lngCount As Long
    
    EnumValue HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\TypedURLs", strValueName(), strData(), lngDataType(), lngCount
    
    
    Dim lngIncrement As Long
    
    With cboTypedURLs
        For lngIncrement = 0 To lngCount - 1
            If lngDataType(lngIncrement) = REG_SZ Then
                .AddItem strData(lngIncrement)
            End If
        Next lngIncrement
    End With
End Sub
