Attribute VB_Name = "shlobj"
Option Explicit


Public Declare Function SHGetSpecialFolderPath Lib "shell32" Alias "SHGetSpecialFolderPathA" (ByVal hwnd As Long, ByVal pszPath As String, ByVal csidl As Long, ByVal fCreate As Boolean) As Boolean


Public Const CSIDL_DESKTOP = &H0
Public Const CSIDL_INTERNET = &H1
Public Const CSIDL_PROGRAMS = &H2
Public Const CSIDL_CONTROLS = &H3
Public Const CSIDL_PRINTERS = &H4
Public Const CSIDL_PERSONAL = &H5
Public Const CSIDL_FAVORITES = &H6
Public Const CSIDL_STARTUP = &H7
Public Const CSIDL_RECENT = &H8
Public Const CSIDL_SENDTO = &H9
Public Const CSIDL_BITBUCKET = &HA
Public Const CSIDL_STARTMENU = &HB
Public Const CSIDL_DESKTOPDIRECTORY = &H10
Public Const CSIDL_DRIVES = &H11
Public Const CSIDL_NETWORK = &H12
Public Const CSIDL_NETHOOD = &H13
Public Const CSIDL_FONTS = &H14
Public Const CSIDL_TEMPLATES = &H15
Public Const CSIDL_COMMON_STARTMENU = &H16
Public Const CSIDL_COMMON_PROGRAMS = &H17
Public Const CSIDL_COMMON_STARTUP = &H18
Public Const CSIDL_COMMON_DESKTOPDIRECTORY = &H19
Public Const CSIDL_APPDATA = &H1A
Public Const CSIDL_PRINTHOOD = &H1B
Public Const CSIDL_LOCAL_APPDATA = &H1C
Public Const CSIDL_ALTSTARTUP = &H1D
Public Const CSIDL_COMMON_ALTSTARTUP = &H1E
Public Const CSIDL_COMMON_FAVORITES = &H1F
Public Const CSIDL_INTERNET_CACHE = &H20
Public Const CSIDL_COOKIES = &H21
Public Const CSIDL_HISTORY = &H22
Public Const CSIDL_COMMON_APPDATA = &H23
Public Const CSIDL_WINDOWS = &H24
Public Const CSIDL_SYSTEM = &H25
Public Const CSIDL_PROGRAM_FILES = &H26
Public Const CSIDL_MYPICTURES = &H27
Public Const CSIDL_PROFILE = &H28
Public Const CSIDL_SYSTEMX86 = &H29
Public Const CSIDL_PROGRAM_FILESX86 = &H2A
Public Const CSIDL_PROGRAM_FILES_COMMON = &H2B
Public Const CSIDL_PROGRAM_FILES_COMMONX86 = &H2C
Public Const CSIDL_COMMON_TEMPLATES = &H2D
Public Const CSIDL_COMMON_DOCUMENTS = &H2E
Public Const CSIDL_COMMON_ADMINTOOLS = &H2F
Public Const CSIDL_ADMINTOOLS = &H30
Public Const CSIDL_CONNECTIONS = &H31

Public Const CSIDL_FLAG_CREATE = &H8000
Public Const CSIDL_FLAG_DONT_VERIFY = &H4000
Public Const CSIDL_FLAG_MASK = &HFF00


Public Function Get_FolderPath(ByVal hwnd As Long, ByVal lngFolder As Long) As String
    If Function_Exist("shell32.dll", "SHGetSpecialFolderPathA") = True Then
        Dim strPath As String
        strPath = String$(MAX_PATH, &H0)
        
        If SHGetSpecialFolderPath(hwnd, strPath, lngFolder, False) = False Then Failed "SHGetSpecialFolderPath"
        Get_FolderPath = Fix_Dir(Fix_NullTermStr(strPath))
    End If
End Function
