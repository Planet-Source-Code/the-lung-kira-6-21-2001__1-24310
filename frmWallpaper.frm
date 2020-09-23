VERSION 5.00
Begin VB.Form frmWallpaper 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Wallpaper"
   ClientHeight    =   3975
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5415
   Icon            =   "frmWallpaper.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   5415
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cboDisplay 
      Height          =   315
      Left            =   3600
      TabIndex        =   6
      Top             =   2160
      Width           =   1695
   End
   Begin VB.ListBox lstNewWallpaper 
      Height          =   840
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   5175
   End
   Begin VB.TextBox txtCurrent 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   5175
   End
   Begin VB.CommandButton cmdChoose 
      Caption         =   "Choose"
      Height          =   350
      Left            =   4320
      TabIndex        =   8
      Top             =   3480
      Width           =   975
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   350
      Left            =   3360
      TabIndex        =   7
      Top             =   3480
      Width           =   975
   End
   Begin VB.Label lblPreview 
      Caption         =   "Preview"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Image imgPreview 
      Height          =   1335
      Left            =   120
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label lblDisplay 
      Caption         =   "Display"
      Height          =   255
      Left            =   2160
      TabIndex        =   5
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label lblNewWallpaper 
      Caption         =   "New Wallpaper"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label lblCurrent 
      Caption         =   "Current"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmWallpaper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdApply_Click()
    Select Case cboDisplay.ListIndex
        Case 0: SaveRegSetting HKEY_CURRENT_USER, "Control Panel\Desktop", "WallpaperStyle", "0", REG_SZ
        Case 1: SaveRegSetting HKEY_CURRENT_USER, "Control Panel\Desktop", "WallpaperStyle", "2", REG_SZ
        Case 2: SaveRegSetting HKEY_CURRENT_USER, "Control Panel\Desktop", "WallpaperStyle", "1", REG_SZ
    End Select
    
    Select Case lstNewWallpaper.List(lstNewWallpaper.ListIndex)
        Case "Default"
            SaveRegSetting HKEY_CURRENT_USER, "Control Panel\Desktop", "Wallpaper", "", REG_SZ
            If SystemParametersInfo(SPI_SETDESKWALLPAPER, 0, &H0, 0) = 0 Then Failed "SystemParametersInfo"
            
            txtCurrent.Text = ""
        Case "None"
            SaveRegSetting HKEY_CURRENT_USER, "Control Panel\Desktop", "Wallpaper", "", REG_SZ
            If SystemParametersInfo(SPI_SETDESKWALLPAPER, 0, "", 0) = 0 Then Failed "SystemParametersInfo"
            
            txtCurrent.Text = ""
        Case Else
            SaveRegSetting HKEY_CURRENT_USER, "Control Panel\Desktop", "Wallpaper", lstNewWallpaper.List(lstNewWallpaper.ListIndex), REG_SZ
            If SystemParametersInfo(SPI_SETDESKWALLPAPER, 0, ByVal lstNewWallpaper.List(lstNewWallpaper.ListIndex), SPIF_UPDATEINIFILE Or SPIF_SENDWININICHANGE) = 0 Then Failed "SystemParametersInfo"
            
            txtCurrent.Text = lstNewWallpaper.List(lstNewWallpaper.ListIndex)
    End Select
End Sub

Private Sub cmdChoose_Click()
    Dim strFileName As String
    strFileName = GetOpenName(frmWallpaper.hwnd, "All Files (*.*)" & Chr$(0) & "*.*" & Chr$(0), 2, "Open", OFN_EXPLORER Or OFN_FILEMUSTEXIST Or OFN_HIDEREADONLY Or OFN_DONTADDTORECENT)
    If strFileName = "" Then Exit Sub
    
    If FileSize_Name(strFileName) = 0 Then
        If MessageBoxEx(&H0, "File size is 0.", "Error", MB_OK Or MB_ICONWARNING Or MB_SETFOREGROUND, 0) = 0 Then Failed "MessageBoxEx"
    Else
        lstNewWallpaper.AddItem strFileName
    End If
End Sub

Private Sub Form_Load()
    With lstNewWallpaper
        .AddItem "Default"
        .AddItem "None"
    End With
    With cboDisplay
        .AddItem "Center"
        .AddItem "Stretch"
        .AddItem "Tiled"
    End With
    
    
    Dim strWallpaper As String
    strWallpaper = GetRegSetting(HKEY_CURRENT_USER, "Control Panel\Desktop", "Wallpaper")
    If strWallpaper <> "" Then
        lstNewWallpaper.AddItem strWallpaper
        txtCurrent.Text = strWallpaper
    End If
    
    Select Case GetRegSetting(HKEY_CURRENT_USER, "Control Panel\Desktop", "WallpaperStyle")
        Case 0: cboDisplay.ListIndex = 0
        Case 1: cboDisplay.ListIndex = 2
        Case 2: cboDisplay.ListIndex = 1
    End Select
End Sub

Private Sub lstNewWallpaper_Click()
    imgPreview.Picture = Nothing
    
    If lstNewWallpaper.ListIndex > 1 Then
        If File_Exist(lstNewWallpaper.List(lstNewWallpaper.ListIndex)) = True Then
            On Error Resume Next
            imgPreview.Picture = LoadPicture(lstNewWallpaper.List(lstNewWallpaper.ListIndex))
            On Error GoTo 0
        End If
    End If
End Sub
