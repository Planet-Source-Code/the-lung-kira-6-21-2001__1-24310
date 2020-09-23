VERSION 5.00
Begin VB.Form frmDirectories 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Directories"
   ClientHeight    =   3900
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9255
   Icon            =   "frmDirectories.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   9255
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   350
      Left            =   8160
      TabIndex        =   3
      Top             =   3480
      Width           =   975
   End
   Begin VB.TextBox txtDirectories 
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   3120
      Width           =   9015
   End
   Begin VB.ListBox lstDirectories 
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2760
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   9015
   End
   Begin VB.Label lblDirectories 
      Caption         =   "Directories"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "frmDirectories"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdSave_Click()
    Dim strFileName As String
    Dim lngIncrement As Long
    Dim strOutput As String
    
    strFileName = GetSaveName(frmDirectories.hwnd, "All Files (*.*)" & Chr$(0) & "*.*" & Chr$(0), 2, "Save", OFN_EXPLORER Or OFN_HIDEREADONLY Or OFN_OVERWRITEPROMPT Or OFN_DONTADDTORECENT)
    If strFileName = "" Then Exit Sub
    
    With lstDirectories
        For lngIncrement = 0 To .ListCount - 2
            strOutput = strOutput & .List(lngIncrement) & Chr$(13) & Chr$(10)
        Next lngIncrement
        
        strOutput = strOutput & .List(lngIncrement)
    End With
    
    WriteFile_String strFileName, strOutput, 0, CREATE_ALWAYS
End Sub

Private Sub Form_Load()
    If Function_Exist("shell32.dll", "SHGetSpecialFolderPathA") = False Then
        lblDirectories.Enabled = False
        lstDirectories.Enabled = False
        txtDirectories.Enabled = False
        cmdSave.Enabled = False
        
        Exit Sub
    End If
    
    With lstDirectories
        .AddItem Left$("Admin Tools" & Space$(22), 22) & Get_FolderPath(frmDirectories.hwnd, CSIDL_ADMINTOOLS)
        .AddItem Left$("Alt Startup" & Space$(22), 22) & Get_FolderPath(frmDirectories.hwnd, CSIDL_ALTSTARTUP)
        .AddItem Left$("App Data" & Space$(22), 22) & Get_FolderPath(frmDirectories.hwnd, CSIDL_APPDATA)
        .AddItem Left$("Common Admin Tools" & Space$(22), 22) & Get_FolderPath(frmDirectories.hwnd, CSIDL_COMMON_ADMINTOOLS)
        .AddItem Left$("Common Alt Startup" & Space$(22), 22) & Get_FolderPath(frmDirectories.hwnd, CSIDL_COMMON_ALTSTARTUP)
        .AddItem Left$("Common App Data" & Space$(22), 22) & Get_FolderPath(frmDirectories.hwnd, CSIDL_COMMON_APPDATA)
        .AddItem Left$("Common Desktop" & Space$(22), 22) & Get_FolderPath(frmDirectories.hwnd, CSIDL_COMMON_DESKTOPDIRECTORY)
        .AddItem Left$("Common Documents" & Space$(22), 22) & Get_FolderPath(frmDirectories.hwnd, CSIDL_COMMON_DOCUMENTS)
        .AddItem Left$("Common Favorites" & Space$(22), 22) & Get_FolderPath(frmDirectories.hwnd, CSIDL_COMMON_FAVORITES)
        .AddItem Left$("Common Program Files" & Space$(22), 22) & Get_FolderPath(frmDirectories.hwnd, CSIDL_PROGRAM_FILES_COMMON)
        .AddItem Left$("Common Programs" & Space$(22), 22) & Get_FolderPath(frmDirectories.hwnd, CSIDL_COMMON_PROGRAMS)
        .AddItem Left$("Common StartMenu" & Space$(22), 22) & Get_FolderPath(frmDirectories.hwnd, CSIDL_COMMON_STARTMENU)
        .AddItem Left$("Common Startup" & Space$(22), 22) & Get_FolderPath(frmDirectories.hwnd, CSIDL_COMMON_STARTUP)
        .AddItem Left$("Common Templates" & Space$(22), 22) & Get_FolderPath(frmDirectories.hwnd, CSIDL_COMMON_TEMPLATES)
        .AddItem Left$("Controls" & Space$(22), 22) & Get_FolderPath(frmDirectories.hwnd, CSIDL_CONTROLS)
        .AddItem Left$("Cookies" & Space$(22), 22) & Get_FolderPath(frmDirectories.hwnd, CSIDL_COOKIES)
        .AddItem Left$("Desktop" & Space$(22), 22) & Get_FolderPath(frmDirectories.hwnd, CSIDL_DESKTOP)
        .AddItem Left$("Desktop Directory" & Space$(22), 22) & Get_FolderPath(frmDirectories.hwnd, CSIDL_DESKTOPDIRECTORY)
        .AddItem Left$("Drives" & Space$(22), 22) & Get_FolderPath(frmDirectories.hwnd, CSIDL_DRIVES)
        .AddItem Left$("Favorites" & Space$(22), 22) & Get_FolderPath(frmDirectories.hwnd, CSIDL_FAVORITES)
        .AddItem Left$("Fonts" & Space$(22), 22) & Get_FolderPath(frmDirectories.hwnd, CSIDL_FONTS)
        .AddItem Left$("History" & Space$(22), 22) & Get_FolderPath(frmDirectories.hwnd, CSIDL_HISTORY)
        .AddItem Left$("Internet" & Space$(22), 22) & Get_FolderPath(frmDirectories.hwnd, CSIDL_INTERNET)
        .AddItem Left$("Internet Cache" & Space$(22), 22) & Get_FolderPath(frmDirectories.hwnd, CSIDL_INTERNET_CACHE)
        .AddItem Left$("Local App Data" & Space$(22), 22) & Get_FolderPath(frmDirectories.hwnd, CSIDL_LOCAL_APPDATA)
        .AddItem Left$("My Documents" & Space$(22), 22) & Get_FolderPath(frmDirectories.hwnd, CSIDL_PERSONAL)
        .AddItem Left$("My Network Places" & Space$(22), 22) & Get_FolderPath(frmDirectories.hwnd, CSIDL_NETHOOD)
        .AddItem Left$("My Pictures" & Space$(22), 22) & Get_FolderPath(frmDirectories.hwnd, CSIDL_MYPICTURES)
        .AddItem Left$("Network Neighborhood" & Space$(22), 22) & Get_FolderPath(frmDirectories.hwnd, CSIDL_NETWORK)
        .AddItem Left$("Printers" & Space$(22), 22) & Get_FolderPath(frmDirectories.hwnd, CSIDL_PRINTERS)
        .AddItem Left$("Print Hood" & Space$(22), 22) & Get_FolderPath(frmDirectories.hwnd, CSIDL_PRINTHOOD)
        .AddItem Left$("Profile" & Space$(22), 22) & Get_FolderPath(frmDirectories.hwnd, CSIDL_PROFILE)
        .AddItem Left$("Program Files" & Space$(22), 22) & Get_FolderPath(frmDirectories.hwnd, CSIDL_PROGRAM_FILES)
        .AddItem Left$("Programs" & Space$(22), 22) & Get_FolderPath(frmDirectories.hwnd, CSIDL_PROGRAMS)
        .AddItem Left$("Recent" & Space$(22), 22) & Get_FolderPath(frmDirectories.hwnd, CSIDL_RECENT)
        .AddItem Left$("RecyleBin" & Space$(22), 22) & Get_FolderPath(frmDirectories.hwnd, CSIDL_BITBUCKET)
        .AddItem Left$("SendTo" & Space$(22), 22) & Get_FolderPath(frmDirectories.hwnd, CSIDL_SENDTO)
        .AddItem Left$("StartMenu" & Space$(22), 22) & Get_FolderPath(frmDirectories.hwnd, CSIDL_STARTMENU)
        .AddItem Left$("Startup" & Space$(22), 22) & Get_FolderPath(frmDirectories.hwnd, CSIDL_STARTUP)
        .AddItem Left$("System" & Space$(22), 22) & Get_FolderPath(frmDirectories.hwnd, CSIDL_SYSTEM)
        .AddItem Left$("Templates" & Space$(22), 22) & Get_FolderPath(frmDirectories.hwnd, CSIDL_TEMPLATES)
        .AddItem Left$("Windows" & Space$(22), 22) & Get_FolderPath(frmDirectories.hwnd, CSIDL_WINDOWS)
    End With
End Sub

Private Sub lstDirectories_Click()
    txtDirectories.Text = Right$(lstDirectories.List(lstDirectories.ListIndex), Len(lstDirectories.List(lstDirectories.ListIndex)) - 22)
End Sub
