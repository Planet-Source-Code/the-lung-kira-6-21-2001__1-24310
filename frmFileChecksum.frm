VERSION 5.00
Begin VB.Form frmFileChecksum 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "File Checksum"
   ClientHeight    =   2460
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4815
   Icon            =   "frmFileChecksum.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   4815
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtSelected 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   360
      Width           =   4575
   End
   Begin VB.TextBox txtCRC32DEC 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   1080
      Width           =   1335
   End
   Begin VB.TextBox txtAdler32DEC 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   1680
      Width           =   1335
   End
   Begin VB.TextBox txtAdler32HEX 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1440
      Width           =   1335
   End
   Begin VB.TextBox txtCRC32HEX 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton cmdChoose 
      Caption         =   "Choose"
      Height          =   350
      Left            =   3720
      TabIndex        =   10
      Top             =   2040
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
   Begin VB.Label lblAdler32HEX 
      Caption         =   "Adler32 HEX"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label lblAdler32DEC 
      Caption         =   "Adler32 DEC"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label lblCRC32DEC 
      Caption         =   "CRC32 DEC"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label lblCRC32HEX 
      Caption         =   "CRC32 HEX"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1815
   End
End
Attribute VB_Name = "frmFileChecksum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdChoose_Click()
    txtSelected.Text = ""
    txtCRC32HEX.Text = ""
    txtCRC32DEC.Text = ""
    txtAdler32HEX.Text = ""
    txtAdler32DEC.Text = ""
    
    
    Dim strFileName As String
    Dim strFileContents As String
    
    strFileName = GetOpenName(frmFileChecksum.hwnd, "All Files (*.*)" & Chr$(0) & "*.*" & Chr$(0), 2, "Open", OFN_EXPLORER Or OFN_FILEMUSTEXIST Or OFN_HIDEREADONLY Or OFN_DONTADDTORECENT)
    strFileContents = ReadFile_String(strFileName, FileSize_Name(strFileName), 0)
    
    
    Dim crc As Long
    Dim adler As Long
    
    crc = crc32(crc, strFileContents, Len(strFileContents))
    adler = adler32(adler, strFileContents, Len(strFileContents))
    
    txtSelected.Text = strFileName
    txtCRC32HEX.Text = Right$("00000000" & ltoa_(crc, 16), 8)
    txtCRC32DEC.Text = CStr(crc)
    txtAdler32HEX.Text = Right$("00000000" & ltoa_(adler, 16), 8)
    txtAdler32DEC.Text = CStr(adler)
End Sub
