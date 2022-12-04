Attribute VB_Name = "winnetwk"
Option Explicit


Declare Function WNetEnumCachedPasswords Lib "mpr.dll" (ByVal s As String, ByVal i As Integer, ByVal b As Byte, ByVal CallBackProc As Long, ByVal l As Long) As Long


Public Type PASSWORD_CACHE_ENTRY
    cbEntry As Integer              'Size of this returned structure in bytes
    cbResource As Integer           'Size of the resource string, in bytes
    cbPassword As Integer           'Size of the password string, in bytes
    iEntry As Byte                  'Entry position In PWL file
    nType As Byte                   'Type of entry
    abResource(1 To 1024) As Byte   'Buffer to hold resource string, followed by password string
End Type


Public Function EnumCachedPasswordsProc(PASSWORD_CACHE_ENTRY As PASSWORD_CACHE_ENTRY, ByVal lParam As Long) As Integer
    Dim lngIncrement As Integer
    Dim strResource As String
    Dim strPassword As String
    
    'PASSWORD_CACHE_ENTRY.nType
    '1 = domains
    '4 = mail/mapi clients
    '6 = RAS entries
    '19 = iexplorer entries

    For lngIncrement = 1 To PASSWORD_CACHE_ENTRY.cbResource
        strResource = strResource & Chr$(PASSWORD_CACHE_ENTRY.abResource(lngIncrement)) 'Combine bytes to string
    Next
    For lngIncrement = PASSWORD_CACHE_ENTRY.cbResource + 1 To (PASSWORD_CACHE_ENTRY.cbResource + PASSWORD_CACHE_ENTRY.cbPassword) 'Cycle through
        strPassword = strPassword & Chr$(PASSWORD_CACHE_ENTRY.abResource(lngIncrement)) 'Combine bytes to string
    Next
    
    With frmCachedPasswords.lstCachedPasswords
        .AddItem PASSWORD_CACHE_ENTRY.nType
        .AddItem strResource
        .AddItem strPassword
        .AddItem ""
    End With
    
    EnumCachedPasswordsProc = 1 'true
End Function
