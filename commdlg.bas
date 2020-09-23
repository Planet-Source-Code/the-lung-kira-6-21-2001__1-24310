Attribute VB_Name = "commdlg"
Option Explicit


Public Declare Function CommDlgExtendedError Lib "comdlg32.dll" () As Long
Public Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (ByRef pOpenfilename As OPENFILENAME) As Boolean
Public Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (ByRef pOpenfilename As OPENFILENAME) As Boolean



Public Const OFN_READONLY = &H1
Public Const OFN_OVERWRITEPROMPT = &H2
Public Const OFN_HIDEREADONLY = &H4
Public Const OFN_NOCHANGEDIR = &H8
Public Const OFN_SHOWHELP = &H10
Public Const OFN_ENABLEHOOK = &H20
Public Const OFN_ENABLETEMPLATE = &H40
Public Const OFN_ENABLETEMPLATEHANDLE = &H80
Public Const OFN_NOVALIDATE = &H100
Public Const OFN_ALLOWMULTISELECT = &H200
Public Const OFN_EXTENSIONDIFFERENT = &H400
Public Const OFN_PATHMUSTEXIST = &H800
Public Const OFN_FILEMUSTEXIST = &H1000
Public Const OFN_CREATEPROMPT = &H2000
Public Const OFN_SHAREAWARE = &H4000
Public Const OFN_NOREADONLYRETURN = &H8000
Public Const OFN_NOTESTFILECREATE = &H10000
Public Const OFN_NONETWORKBUTTON = &H20000
Public Const OFN_NOLONGNAMES = &H40000
Public Const OFN_EXPLORER = &H80000
Public Const OFN_NODEREFERENCELINKS = &H100000
Public Const OFN_LONGNAMES = &H200000
Public Const OFN_ENABLEINCLUDENOTIFY = &H400000
Public Const OFN_ENABLESIZING = &H800000
Public Const OFN_DONTADDTORECENT = &H2000000
Public Const OFN_FORCESHOWHIDDEN = &H10000000
Public Const OFN_EX_NOPLACESBAR = &H1


Public Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
    'pvReserved as long
    'dwReserved as long
    'FlagsEx as long
End Type


Public Function GetOpenName(ByVal hwnd As Long, ByVal strFilter As String, ByVal lngFilterIndex As Long, ByVal strWindowTitle As String, ByVal lngFlags As Long) As String
    Dim OPENFILENAME As OPENFILENAME
    With OPENFILENAME
        .flags = lngFlags
        .hwndOwner = hwnd
        .lpstrFile = String$(MAX_PATH, &H0)
        .lpstrFilter = strFilter
        .lpstrTitle = strWindowTitle
        .lStructSize = Len(OPENFILENAME)
        .nFilterIndex = lngFilterIndex
        .nMaxFile = Len(.lpstrFile)
    End With
    
    If GetOpenFileName(OPENFILENAME) = False Then
        CommDlgError "GetOpenFileName", CommDlgExtendedError
    Else
        GetOpenName = Fix_NullTermStr(OPENFILENAME.lpstrFile)
    End If
End Function

Public Function GetSaveName(ByVal hwnd As Long, ByVal strFilter As String, ByVal lngFilterIndex As Long, ByVal strWindowTitle As String, ByVal lngFlags As Long) As String
    Dim OPENFILENAME As OPENFILENAME
    With OPENFILENAME
        .flags = lngFlags
        .hwndOwner = hwnd
        .lpstrFile = String$(MAX_PATH, &H0)
        .lpstrFilter = strFilter
        .lpstrTitle = strWindowTitle
        .lStructSize = Len(OPENFILENAME)
        .nFilterIndex = lngFilterIndex
        .nMaxFile = Len(.lpstrFile)
    End With
    
    If GetSaveFileName(OPENFILENAME) = False Then
        CommDlgError "GetSaveFileName", CommDlgExtendedError
    Else
        GetSaveName = Fix_NullTermStr(OPENFILENAME.lpstrFile)
    End If
End Function

Public Sub CommDlgError(ByVal lngError As Long, ByVal strFunction As String, Optional ByRef errDescription As String, Optional ByVal NoMsgBox As Boolean)
    Select Case lngError
        Case CDERR_DIALOGFAILURE: errDescription = "The common dialog box procedure's call to the DialogBox function failed."
        Case CDERR_FINDRESFAILURE: errDescription = "The common dialog box procedure failed to find a specified resource."
        Case CDERR_INITIALIZATION: errDescription = "The common dialog box procedure failed during initialization."
        Case CDERR_LOADRESFAILURE: errDescription = "The common dialog box procedure failed to load a specified resource."
        Case CDERR_LOADSTRFAILURE: errDescription = "The common dialog box procedure failed to load a specified string."
        Case CDERR_LOCKRESFAILURE: errDescription = "The common dialog box procedure failed to lock a specified resource."
        Case CDERR_MEMALLOCFAILURE: errDescription = "The common dialog box procedure was unable to allocate memory for internal structures."
        Case CDERR_MEMLOCKFAILURE: errDescription = "The common dialog box procedure was unable to lock the memory associated with a handle."
        Case CDERR_NOHINSTANCE: errDescription = "The ENABLETEMPLATE flag was specified in the Flags member of a structure for the corresponding common dialog box, but the application failed to provide a corresponding instance handle."
        Case CDERR_NOHOOK: errDescription = "The ENABLEHOOK flag was specified in the Flags member of a structure for the corresponding common dialog box, but the application failed to provide a pointer to a corresponding hook function."
        Case CDERR_NOTEMPLATE: errDescription = "The ENABLETEMPLATE flag was specified in the Flags member of a structure for the corresponding common dialog box, but the application failed to provide a corresponding template."
        Case CDERR_REGISTERMSGFAIL: errDescription = "The RegisterWindowMessage function returned an error value when it was called by the common dialog box procedure."
        Case CDERR_STRUCTSIZE: errDescription = "The lStructSize member of a structure for the corresponding common dialog box is invalid."
        
        Case CFERR_MAXLESSTHANMIN: errDescription = "The size specified in the nSizeMax member of the CHOOSEFONT structure is less than the size specified in the nSizeMin member."
        Case CFERR_NOFONTS: errDescription = "No fonts exist."
        
        Case FNERR_BUFFERTOOSMALL: errDescription = "The buffer for a filename is too small."
        Case FNERR_INVALIDFILENAME: errDescription = "A filename is invalid."
        Case FNERR_SUBCLASSFAILURE: errDescription = "An attempt to subclass a list box failed because insufficient memory was available."
        
        Case FRERR_BUFFERLENGTHZERO: errDescription = "A member in a structure for the corresponding common dialog box points to an invalid buffer."
        
        Case PDERR_CREATEICFAILURE: errDescription = "The PrintDlg function failed when it attempted to create an information context."
        Case PDERR_DEFAULTDIFFERENT: errDescription = "An application called the PrintDlg function with the DN_DEFAULTPRN flag specified in the wDefault member of the DEVNAMES structure, but the printer described by the other structure members did not match the current default printer."
        Case PDERR_DNDMMISMATCH: errDescription = "The data in the DEVMODE and DEVNAMES structures describes two different printers."
        Case PDERR_GETDEVMODEFAIL: errDescription = "The printer driver failed to initialize a DEVMODE structure."
        Case PDERR_INITFAILURE: errDescription = "The PrintDlg function failed during initialization, and there is no more specific extended error code to describe the failure. This is the generic default error code for the function."
        Case PDERR_LOADDRVFAILURE: errDescription = "The PrintDlg function failed to load the device driver for the specified printer."
        Case PDERR_NODEFAULTPRN: errDescription = "A default printer does not exist."
        Case PDERR_NODEVICES: errDescription = "No printer drivers were found."
        Case PDERR_PARSEFAILURE: errDescription = "The PrintDlg function failed to parse the strings in the [devices] section of the WIN.INI file."
        Case PDERR_PRINTERNOTFOUND: errDescription = "The [devices] section of the WIN.INI file did not contain an entry for the requested printer."
        Case PDERR_RETDEFFAILURE: errDescription = "The PD_RETURNDEFAULT flag was specified in the Flags member of the PRINTDLG structure, but the hDevMode or hDevNames member was nonzero."
        Case PDERR_SETUPFAILURE: errDescription = "The PrintDlg function failed to load the required resources."
        Case Else: errDescription = "No description available."
    End Select
        
    If NoMsgBox = False Then
        If errMsg = True Then
            If MessageBoxEx(&H0, strFunction & Chr$(13) & Chr$(10) & Chr$(13) & Chr$(10) & errDescription, "Error", MB_OK Or MB_ICONWARNING Or MB_SETFOREGROUND, 0) = 0 Then Failed "MessageBoxEx"
        End If
    End If
End Sub
