VERSION 5.00
Begin VB.Form frmGIF 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GIF"
   ClientHeight    =   3900
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4815
   Icon            =   "frmGIF.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   4815
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtPixelAspectRatio 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   3120
      Width           =   1455
   End
   Begin VB.TextBox txtBackgroundColorIndex 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   2880
      Width           =   1455
   End
   Begin VB.CheckBox chkGlobalColorTableFlag 
      Enabled         =   0   'False
      Height          =   255
      Left            =   4440
      TabIndex        =   17
      Top             =   2640
      Width           =   255
   End
   Begin VB.TextBox txtColorResolution 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CheckBox chkSortFlag 
      Enabled         =   0   'False
      Height          =   255
      Left            =   4440
      TabIndex        =   13
      Top             =   2160
      Width           =   255
   End
   Begin VB.TextBox txtGlobalColorTableSize 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   1920
      Width           =   1455
   End
   Begin VB.TextBox txtLogicalScreenHeight 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   1680
      Width           =   1455
   End
   Begin VB.TextBox txtLogicalScreenWidth 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1440
      Width           =   1455
   End
   Begin VB.TextBox txtVersion 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton cmdChoose 
      Caption         =   "Choose"
      Height          =   350
      Left            =   3720
      TabIndex        =   22
      Top             =   3480
      Width           =   975
   End
   Begin VB.CheckBox chkSignature 
      Enabled         =   0   'False
      Height          =   255
      Left            =   4440
      TabIndex        =   3
      Top             =   840
      Width           =   255
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
      Width           =   4575
   End
   Begin VB.Label lblBackgroundColorIndex 
      Caption         =   "Background Color Index"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   2880
      Width           =   2415
   End
   Begin VB.Label lblPixelAspectRatio 
      Caption         =   "Pixel Aspect Ratio"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   3120
      Width           =   2415
   End
   Begin VB.Label lblGlobalColorTableFlag 
      Caption         =   "Global Color Table Flag"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   2640
      Width           =   2415
   End
   Begin VB.Label lblColorResolution 
      Caption         =   "Color Resolution"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   2400
      Width           =   2415
   End
   Begin VB.Label lblSortFlag 
      Caption         =   "Sort Flag"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   2160
      Width           =   2415
   End
   Begin VB.Label lblGlobalColorTableSize 
      Caption         =   "Global Color Table Size"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1920
      Width           =   2415
   End
   Begin VB.Label lblLogicalScreenHeight 
      Caption         =   "Logical Screen Height"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1680
      Width           =   2415
   End
   Begin VB.Label lblLogicalScreenWidth 
      Caption         =   "Logical Screen Width"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   2415
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   2415
   End
   Begin VB.Label lblSignature 
      Caption         =   "Signature"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   2415
   End
   Begin VB.Label lblSelected 
      Caption         =   "Selected"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmGIF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdChoose_Click()
    chkSignature.value = 0
    txtVersion.Text = ""
    txtLogicalScreenWidth.Text = ""
    txtLogicalScreenHeight.Text = ""
    chkGlobalColorTableFlag.value = 0
    txtColorResolution.Text = ""
    chkSortFlag.value = 0
    txtGlobalColorTableSize.Text = ""
    txtBackgroundColorIndex.Text = ""
    txtPixelAspectRatio.Text = ""
    
    
    Dim strFileName As String
    strFileName = GetOpenName(frmGIF.hwnd, "All Files (*.*)" & Chr$(0) & "*.*" & Chr$(0), 2, "Open", OFN_EXPLORER Or OFN_FILEMUSTEXIST Or OFN_HIDEREADONLY Or OFN_DONTADDTORECENT)
    txtSelected.Text = strFileName
    If strFileName = "" Then Exit Sub
    
    If FileSize_Name(strFileName) = 0 Then
        If MessageBoxEx(&H0, "File size is 0.", "Error", MB_OK Or MB_ICONWARNING Or MB_SETFOREGROUND, 0) = 0 Then Failed "MessageBoxEx"
        Exit Sub
    End If
    
    
    Dim strFileContents As String
    
    strFileContents = ReadFile_String(strFileName, 13, 0)
    If Len(strFileContents) < 13 Then
        strFileContents = strFileContents & String$(13 - Len(strFileContents), &H0)
    End If
    
    
    If Mid$(strFileContents, 1, 3) <> "GIF" Then Exit Sub
    Dim strBinary As String
    
    chkSignature.value = 1
    txtVersion.Text = Mid$(strFileContents, 4, 3)
    txtLogicalScreenWidth.Text = CStr(strtoul_(Right$("00" & ltoa_(Asc(Mid$(strFileContents, 8, 1)), 16), 2) & _
                                    Right$("00" & ltoa_(Asc(Mid$(strFileContents, 7, 1)), 16), 2), 16))
    txtLogicalScreenHeight.Text = CStr(strtoul_(Right$("00" & ltoa_(Asc(Mid$(strFileContents, 10, 1)), 16), 2) & _
                                    Right$("00" & ltoa_(Asc(Mid$(strFileContents, 9, 1)), 16), 2), 16))
                                    
    strBinary = ltoa_(Asc(Mid$(strFileContents, 11, 1)), 2)
    chkGlobalColorTableFlag.value = CLng(Mid$(strBinary, 8, 1))
    txtColorResolution.Text = CStr(strtoul_(Mid$(strBinary, 5, 3), 2))
    chkSortFlag.value = CLng(Mid$(strBinary, 4, 1))
    txtGlobalColorTableSize.Text = CStr(strtoul_(Mid$(strBinary, 1, 3), 2))
    
    txtBackgroundColorIndex.Text = CStr(Asc(Mid$(strFileContents, 12, 1)))
    txtPixelAspectRatio.Text = CStr(Asc(Mid$(strFileContents, 13, 1)))
End Sub
