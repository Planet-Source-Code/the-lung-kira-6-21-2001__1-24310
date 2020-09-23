VERSION 5.00
Begin VB.Form frmFileAttributes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "File Attributes"
   ClientHeight    =   5205
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6135
   Icon            =   "frmFileAttributes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   6135
   StartUpPosition =   3  'Windows Default
   Begin VB.DriveListBox drvFileAttributes 
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   4800
      Width           =   2175
   End
   Begin VB.FileListBox fileFileAttributes 
      Height          =   2040
      Hidden          =   -1  'True
      Left            =   120
      System          =   -1  'True
      TabIndex        =   3
      Top             =   2760
      Width           =   2175
   End
   Begin VB.DirListBox dirFileAttributes 
      Height          =   1890
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   2175
   End
   Begin VB.TextBox txtSelected 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   360
      Width           =   5895
   End
   Begin VB.CheckBox chkTemporary 
      Height          =   255
      Left            =   5760
      TabIndex        =   30
      Top             =   3720
      Width           =   255
   End
   Begin VB.CheckBox chkSystem 
      Height          =   255
      Left            =   5760
      TabIndex        =   28
      Top             =   3480
      Width           =   255
   End
   Begin VB.CheckBox chkSparseFile 
      Enabled         =   0   'False
      Height          =   255
      Left            =   5760
      TabIndex        =   26
      Top             =   3240
      Width           =   255
   End
   Begin VB.CheckBox chkReparsePoint 
      Enabled         =   0   'False
      Height          =   255
      Left            =   5760
      TabIndex        =   24
      Top             =   3000
      Width           =   255
   End
   Begin VB.CheckBox chkReadOnly 
      Height          =   255
      Left            =   5760
      TabIndex        =   22
      Top             =   2760
      Width           =   255
   End
   Begin VB.CheckBox chkOffline 
      Height          =   255
      Left            =   5760
      TabIndex        =   20
      Top             =   2520
      Width           =   255
   End
   Begin VB.CheckBox chkNotContentIndexed 
      Height          =   255
      Left            =   5760
      TabIndex        =   18
      Top             =   2280
      Width           =   255
   End
   Begin VB.CheckBox chkNormal 
      Height          =   255
      Left            =   5760
      TabIndex        =   16
      Top             =   2040
      Width           =   255
   End
   Begin VB.CheckBox chkHidden 
      Height          =   255
      Left            =   5760
      TabIndex        =   14
      Top             =   1800
      Width           =   255
   End
   Begin VB.CheckBox chkEncrypted 
      Enabled         =   0   'False
      Height          =   255
      Left            =   5760
      TabIndex        =   12
      Top             =   1560
      Width           =   255
   End
   Begin VB.CheckBox chkDirectory 
      Enabled         =   0   'False
      Height          =   255
      Left            =   5760
      TabIndex        =   10
      Top             =   1320
      Width           =   255
   End
   Begin VB.CheckBox chkCompressed 
      Enabled         =   0   'False
      Height          =   255
      Left            =   5760
      TabIndex        =   8
      Top             =   1080
      Width           =   255
   End
   Begin VB.CheckBox chkArchive 
      Height          =   255
      Left            =   5760
      TabIndex        =   6
      Top             =   840
      Width           =   255
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Enabled         =   0   'False
      Height          =   350
      Left            =   5040
      TabIndex        =   31
      Top             =   4080
      Width           =   975
   End
   Begin VB.Label lblSelected 
      Caption         =   "Selected"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label lblTemporary 
      Caption         =   "Temporary"
      Height          =   255
      Left            =   2520
      TabIndex        =   29
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Label lblSystem 
      Caption         =   "System"
      Height          =   255
      Left            =   2520
      TabIndex        =   27
      Top             =   3480
      Width           =   1695
   End
   Begin VB.Label lblSparseFile 
      Caption         =   "Sparse File"
      Height          =   255
      Left            =   2520
      TabIndex        =   25
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Label lblReparsePoint 
      Caption         =   "Reparse Point"
      Height          =   255
      Left            =   2520
      TabIndex        =   23
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Label lblReadOnly 
      Caption         =   "Read Only"
      Height          =   255
      Left            =   2520
      TabIndex        =   21
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label lblOffline 
      Caption         =   "Offline"
      Height          =   255
      Left            =   2520
      TabIndex        =   19
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label lblNotContentIndexed 
      Caption         =   "Not Content Indexed"
      Height          =   255
      Left            =   2520
      TabIndex        =   17
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label lblNormal 
      Caption         =   "Normal"
      Height          =   255
      Left            =   2520
      TabIndex        =   15
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label lblHidden 
      Caption         =   "Hidden"
      Height          =   255
      Left            =   2520
      TabIndex        =   13
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label lblEncrypted 
      Caption         =   "Encrypted"
      Height          =   255
      Left            =   2520
      TabIndex        =   11
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label lblDirectory 
      Caption         =   "Directory"
      Height          =   255
      Left            =   2520
      TabIndex        =   9
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label lblCompressed 
      Caption         =   "Compressed"
      Height          =   255
      Left            =   2520
      TabIndex        =   7
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label lblArchive 
      Caption         =   "Archive"
      Height          =   255
      Left            =   2520
      TabIndex        =   5
      Top             =   840
      Width           =   1695
   End
End
Attribute VB_Name = "frmFileAttributes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim FileAttributes As Long

Private Sub chkNormal_Click()
    If chkNormal.value = 1 Then
        chkArchive.value = 0
        chkHidden.value = 0
        chkNotContentIndexed.value = 0
        chkOffline.value = 0
        chkReadOnly.value = 0
        chkSystem.value = 0
        chkTemporary.value = 0
    End If
End Sub

Private Sub cmdApply_Click()
    If txtSelected.Text = "" Then Exit Sub
    
    FileAttributes = &H0
    
    Dim Archive As Long
    Dim Hidden As Long
    Dim NotContentIndexed As Long
    Dim Offline As Long
    Dim ReadOnly As Long
    Dim System As Long
    Dim Temporary As Long
    
    If chkArchive.value = 1 Then Archive = FILE_ATTRIBUTE_ARCHIVE
    If chkHidden.value = 1 Then Hidden = FILE_ATTRIBUTE_HIDDEN
    If chkNotContentIndexed.value = 1 Then NotContentIndexed = FILE_ATTRIBUTE_NOT_CONTENT_INDEXED
    If chkOffline.value = 1 Then Offline = FILE_ATTRIBUTE_OFFLINE
    If chkReadOnly.value = 1 Then ReadOnly = FILE_ATTRIBUTE_READONLY
    If chkSystem.value = 1 Then System = FILE_ATTRIBUTE_SYSTEM
    If chkTemporary.value = 1 Then Temporary = FILE_ATTRIBUTE_TEMPORARY
    
    If chkNormal.value = 1 Then
        FileAttributes = FILE_ATTRIBUTE_NORMAL
    Else
        FileAttributes = Archive Or Hidden Or NotContentIndexed Or Offline Or ReadOnly Or System Or Temporary
    End If

    If SetFileAttributes(txtSelected.Text, FileAttributes) = False Then Failed "SetFileAttributes"
End Sub

Private Sub dirFileAttributes_Change()
    fileFileAttributes.Path = dirFileAttributes.Path
    txtSelected.Text = Fix_Dir(dirFileAttributes.Path)
    Process txtSelected.Text
End Sub

Private Sub dirFileAttributes_Click()
    fileFileAttributes.Path = dirFileAttributes.Path
    txtSelected.Text = Fix_Dir(dirFileAttributes.Path)
    Process txtSelected.Text
End Sub

Private Sub drvFileAttributes_Change()
    On Error Resume Next
    dirFileAttributes.Path = drvFileAttributes.Drive
    On Error GoTo 0
End Sub

Private Sub fileFileAttributes_Click()
    txtSelected.Text = Fix_Dir(dirFileAttributes.Path) & "\" & fileFileAttributes.FileName
    Process txtSelected.Text
End Sub

Private Sub Process(strFileName As String)
    FileAttributes = GetFileAttributes(strFileName)
    If FileAttributes = -1 Then
        Failed "GetFileAttributes"
        cmdApply.Enabled = False
    Else
        cmdApply.Enabled = True
    End If
    
    If FileAttributes And FILE_ATTRIBUTE_ARCHIVE Then chkArchive.value = 1 Else: chkArchive.value = 0
    If FileAttributes And FILE_ATTRIBUTE_COMPRESSED Then chkCompressed.value = 1 Else: chkCompressed.value = 0
    If FileAttributes And FILE_ATTRIBUTE_DIRECTORY Then chkDirectory.value = 1 Else: chkDirectory.value = 0
    If FileAttributes And FILE_ATTRIBUTE_ENCRYPTED Then chkEncrypted.value = 1 Else: chkEncrypted.value = 0
    If FileAttributes And FILE_ATTRIBUTE_HIDDEN Then chkHidden.value = 1 Else: chkHidden.value = 0
    If FileAttributes And FILE_ATTRIBUTE_NORMAL Then chkNormal.value = 1 Else: chkNormal.value = 0
    If FileAttributes And FILE_ATTRIBUTE_NOT_CONTENT_INDEXED Then chkNotContentIndexed.value = 1 Else: chkNotContentIndexed.value = 0
    If FileAttributes And FILE_ATTRIBUTE_OFFLINE Then chkOffline.value = 1 Else: chkOffline.value = 0
    If FileAttributes And FILE_ATTRIBUTE_READONLY Then chkReadOnly.value = 1 Else: chkReadOnly.value = 0
    If FileAttributes And FILE_ATTRIBUTE_REPARSE_POINT Then chkReparsePoint.value = 1 Else: chkReparsePoint.value = 0
    If FileAttributes And FILE_ATTRIBUTE_SPARSE_FILE Then chkSparseFile.value = 1 Else: chkSparseFile.value = 0
    If FileAttributes And FILE_ATTRIBUTE_SYSTEM Then chkSystem.value = 1 Else: chkSystem.value = 0
    If FileAttributes And FILE_ATTRIBUTE_TEMPORARY Then chkTemporary.value = 1 Else: chkTemporary.value = 0
End Sub
