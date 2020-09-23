Attribute VB_Name = "winbase"
Option Explicit


Public Declare Function AdjustTokenPrivileges Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Boolean, ByRef NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, ByRef PreviousState As TOKEN_PRIVILEGES, ByRef ReturnLength As Long) As Boolean
Public Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Boolean
Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef hpvDest As Any, ByRef hpvSource As Any, ByVal cbCopy As Long)
Public Declare Function CreateFile Lib "kernel32.dll" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByRef lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Public Declare Function FileTimeToSystemTime Lib "kernel32.dll" (ByRef lpFileTime As FILETIME, ByRef lpSystemTime As SYSTEMTIME) As Boolean
Public Declare Function FormatMessage Lib "kernel32.dll" Alias "FormatMessageA" (ByVal dwFlags As Long, ByRef lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef Arguments As Long) As Long
Public Declare Function FreeLibrary Lib "kernel32.dll" (ByVal hLibModule As Long) As Boolean
Public Declare Function GetComputerName Lib "kernel32.dll" Alias "GetComputerNameA" (ByVal lpBuffer As String, ByRef nSize As Long) As Boolean
Public Declare Function GetCurrentProcess Lib "kernel32.dll" () As Long
Public Declare Function GetCurrentProcessId Lib "kernel32.dll" () As Long
Public Declare Function GetExitCodeProcess Lib "kernel32.dll" (ByVal hProcess As Long, ByRef lpExitCode As Long) As Boolean
Public Declare Function GetExitCodeThread Lib "kernel32.dll" (ByVal hThread As Long, ByRef lpExitCode As Long) As Boolean
Public Declare Function GetFileAttributes Lib "kernel32.dll" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Public Declare Function GetFileTime Lib "kernel32.dll" (ByVal hFile As Long, ByRef lpCreationTime As FILETIME, ByRef lpLastAccessTime As FILETIME, ByRef lpLastWriteTime As FILETIME) As Boolean
Public Declare Sub GlobalMemoryStatus Lib "kernel32.dll" (ByRef lpBuffer As MEMORYSTATUS)
Public Declare Function GlobalMemoryStatusEx Lib "kernel32.dll" (ByRef lpBuffer As MEMORYSTATUSEX) As Boolean
Public Declare Function GetDiskFreeSpace Lib "kernel32.dll" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, ByRef lpSectorsPerCluster As Long, ByRef lpBytesPerSector As Long, ByRef lpNumberOfFreeClusters As Long, ByRef lpTotalNumberOfClusters As Long) As Boolean
Public Declare Function GetDiskFreeSpaceEx Lib "kernel32.dll" Alias "GetDiskFreeSpaceExA" (ByVal lpRootPathName As String, ByRef lpFreeBytesAvailableToCaller As LARGE_INTEGER, ByRef lpTotalNumberOfBytes As LARGE_INTEGER, ByRef lpTotalNumberOfFreeBytes As LARGE_INTEGER) As Boolean
Public Declare Function GetDriveType Lib "kernel32.dll" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Public Declare Function GetFileSize Lib "kernel32.dll" (ByVal hFile As Long, ByVal lpFileSizeHigh As Long) As Long
Public Declare Function GetFileSizeEx Lib "kernel32.dll" (ByVal hFile As Long, ByRef lpFileSize As LARGE_INTEGER) As Boolean
Public Declare Function GetFileType Lib "kernel32.dll" (ByVal hFile As Long) As Long
Public Declare Function GetLogicalDrives Lib "kernel32.dll" () As Long
Public Declare Function GetModuleHandle Lib "kernel32.dll" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Public Declare Function GetPriorityClass Lib "kernel32.dll" (ByVal hProcess As Long) As Long
Public Declare Function GetProcAddress Lib "kernel32.dll" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Public Declare Function GetProcessAffinityMask Lib "kernel32.dll" (ByVal hProcess As Long, ByRef lpProcessAffinityMask As Long, ByRef lpSystemAffinityMask As Long) As Boolean
Public Declare Function GetProcessIoCounters Lib "kernel32.dll" (ByVal hProcess As Long, ByRef lpIoCounters As IO_COUNTERS) As Boolean
Public Declare Function GetProcessTimes Lib "kernel32.dll" (ByVal hProcess As Long, ByRef lpCreationTime As FILETIME, ByRef lpExitTime As FILETIME, ByRef lpKernelTime As FILETIME, ByRef lpUserTime As FILETIME) As Boolean
Public Declare Function GetProcessVersion Lib "kernel32.dll" (ByVal ProcessId As Long) As Long
Public Declare Sub GetSystemInfo Lib "kernel32.dll" (ByRef lpSystemInfo As SYSTEM_INFO)
Public Declare Function GetSystemPowerStatus Lib "kernel32.dll" (ByRef lpSystemPowerStatus As SYSTEM_POWER_STATUS) As Boolean
Public Declare Function GetThreadPriority Lib "kernel32.dll" (ByVal hThread As Long) As Long
Public Declare Function GetThreadTimes Lib "kernel32.dll" (ByVal hThread As Long, ByRef lpCreationTime As FILETIME, ByRef lpExitTime As FILETIME, ByRef lpKernelTime As FILETIME, ByRef lpUserTime As FILETIME) As Boolean
Public Declare Function GetTickCount Lib "kernel32.dll" () As Long
Public Declare Function GetTimeZoneInformation Lib "kernel32.dll" (ByRef lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long
Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, ByRef nSize As Long) As Boolean
Public Declare Function GetVersionEx Lib "kernel32.dll" Alias "GetVersionExA" (ByRef lpVersionInformation As Any) As Boolean
Public Declare Function GetVolumeInformation Lib "kernel32.dll" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, ByRef lpVolumeSerialNumber As Long, ByRef lpMaximumComponentLength As Long, ByRef lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Boolean
Public Declare Function LoadLibraryEx Lib "kernel32.dll" Alias "LoadLibraryExA" (ByVal lpFileName As String, ByVal hFile As Long, ByVal dwFlags As Long) As Long
Public Declare Function LookupPrivilegeValue Lib "advapi32.dll" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, ByRef lpLuid As LUID) As Boolean
Public Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Boolean, ByVal dwProcessId As Long) As Long
Public Declare Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, ByRef TokenHandle As Long) As Boolean
Public Declare Function OpenThread Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Boolean, ByVal dwThreadId As Long) As Long
Public Declare Function QueryPerformanceCounter Lib "kernel32.dll" (ByRef lpPerformanceCount As LARGE_INTEGER) As Boolean
Public Declare Function QueryPerformanceFrequency Lib "kernel32.dll" (ByRef lpFrequency As LARGE_INTEGER) As Boolean
Public Declare Function ReadFile Lib "kernel32.dll" (ByVal hFile As Long, ByRef lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, ByRef lpNumberOfBytesRead As Long, ByRef lpOverlapped As Any) As Boolean
Public Declare Function ResumeThread Lib "kernel32.dll" (ByVal hThread As Long) As Long
Public Declare Function SetComputerName Lib "kernel32.dll" Alias "SetComputerNameA" (ByVal lpComputerName As String) As Boolean
Public Declare Function SetEndOfFile Lib "kernel32.dll" (ByVal hFile As Long) As Boolean
Public Declare Function SetFileAttributes Lib "kernel32.dll" Alias "SetFileAttributesA" (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Boolean
Public Declare Function SetFilePointer Lib "kernel32.dll" (ByVal hFile As Long, ByVal lDistanceToMove As Long, ByRef lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Public Declare Function SetFileTime Lib "kernel32.dll" (ByVal hFile As Long, ByRef lpCreationTime As FILETIME, ByRef lpLastAccessTime As FILETIME, ByRef lpLastWriteTime As FILETIME) As Boolean
Public Declare Function SetPriorityClass Lib "kernel32.dll" (ByVal hProcess As Long, ByVal dwPriorityClass As Long) As Boolean
Public Declare Function SetSystemTime Lib "kernel32.dll" (ByRef lpSystemTime As SYSTEMTIME) As Boolean
Public Declare Function SetThreadIdealProcessor Lib "kernel32.dll" (ByVal hThread As Long, ByVal dwIdealProcessor As Long) As Long
Public Declare Function SetThreadPriority Lib "kernel32.dll" (ByVal hThread As Long, ByVal nPriority As Long) As Boolean
Public Declare Function SetTimeZoneInformation Lib "kernel32.dll" (ByRef lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Boolean
Public Declare Function SetVolumeLabel Lib "kernel32.dll" Alias "SetVolumeLabelA" (ByVal lpRootPathName As String, ByVal lpVolumeName As String) As Boolean
Public Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
Public Declare Function SuspendThread Lib "kernel32.dll" (ByVal hThread As Long) As Long
Public Declare Function SystemTimeToFileTime Lib "kernel32.dll" (ByRef lpSystemTime As SYSTEMTIME, ByRef lpFileTime As FILETIME) As Boolean
Public Declare Function TerminateProcess Lib "kernel32.dll" (ByVal hProcess As Long, ByVal uExitCode As Long) As Boolean
Public Declare Function TerminateThread Lib "kernel32.dll" (ByVal hThread As Long, ByVal dwExitCode As Long) As Boolean
Public Declare Function WriteFile Lib "kernel32.dll" (ByVal hFile As Long, ByRef lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, ByRef lpNumberOfBytesWritten As Long, ByRef lpOverlapped As Any) As Boolean


Public Const CREATE_NEW = 1
Public Const CREATE_ALWAYS = 2
Public Const OPEN_EXISTING = 3
Public Const OPEN_ALWAYS = 4
Public Const TRUNCATE_EXISTING = 5

Public Const DEBUG_PROCESS = &H1
Public Const DEBUG_ONLY_THIS_PROCESS = &H2
Public Const CREATE_SUSPENDED = &H4
Public Const DETACHED_PROCESS = &H8
Public Const CREATE_NEW_CONSOLE = &H10
Public Const NORMAL_PRIORITY_CLASS = &H20
Public Const IDLE_PRIORITY_CLASS = &H40
Public Const HIGH_PRIORITY_CLASS = &H80
Public Const REALTIME_PRIORITY_CLASS = &H100
Public Const CREATE_NEW_PROCESS_GROUP = &H200
Public Const CREATE_UNICODE_ENVIRONMENT = &H400
Public Const CREATE_SEPARATE_WOW_VDM = &H800
Public Const CREATE_SHARED_WOW_VDM = &H1000
Public Const CREATE_FORCEDOS = &H2000
Public Const BELOW_NORMAL_PRIORITY_CLASS = &H4000
Public Const ABOVE_NORMAL_PRIORITY_CLASS = &H8000
Public Const CREATE_BREAKAWAY_FROM_JOB = &H1000000

Public Const CBR_110 = 110
Public Const CBR_300 = 300
Public Const CBR_600 = 600
Public Const CBR_1200 = 1200
Public Const CBR_2400 = 2400
Public Const CBR_4800 = 4800
Public Const CBR_9600 = 9600
Public Const CBR_14400 = 14400
Public Const CBR_19200 = 19200
Public Const CBR_38400 = 38400
Public Const CBR_56000 = 56000
Public Const CBR_57600 = 57600
Public Const CBR_115200 = 115200
Public Const CBR_128000 = 128000
Public Const CBR_256000 = 256000

Public Const DRIVE_UNKNOWN = 0
Public Const DRIVE_NO_ROOT_DIR = 1
Public Const DRIVE_REMOVABLE = 2
Public Const DRIVE_FIXED = 3
Public Const DRIVE_REMOTE = 4
Public Const DRIVE_CDROM = 5
Public Const DRIVE_RAMDISK = 6

Public Const FILE_BEGIN = 0
Public Const FILE_CURRENT = 1
Public Const FILE_END = 2

Public Const FILE_TYPE_UNKNOWN = &H0
Public Const FILE_TYPE_DISK = &H1
Public Const FILE_TYPE_CHAR = &H2
Public Const FILE_TYPE_PIPE = &H3
Public Const FILE_TYPE_REMOTE = &H8000

Public Const FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100
Public Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200
Public Const FORMAT_MESSAGE_FROM_STRING = &H400
Public Const FORMAT_MESSAGE_FROM_HMODULE = &H800
Public Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Public Const FORMAT_MESSAGE_ARGUMENT_ARRAY = &H2000
Public Const FORMAT_MESSAGE_MAX_WIDTH_MASK = &HFF

Public Const FS_CASE_IS_PRESERVED = FILE_CASE_PRESERVED_NAMES
Public Const FS_CASE_SENSITIVE = FILE_CASE_SENSITIVE_SEARCH
Public Const FS_UNICODE_STORED_ON_DISK = FILE_UNICODE_ON_DISK
Public Const FS_PERSISTENT_ACLS = FILE_PERSISTENT_ACLS
Public Const FS_VOL_IS_COMPRESSED = FILE_VOLUME_IS_COMPRESSED
Public Const FS_FILE_COMPRESSION = FILE_FILE_COMPRESSION
Public Const FS_FILE_ENCRYPTION = FILE_SUPPORTS_ENCRYPTION

Public Const INVALID_HANDLE_VALUE = -1
Public Const INVALID_FILE_SIZE = &HFFFFFFFF
Public Const INVALID_SET_FILE_POINTER = -1

Public Const MAX_COMPUTERNAME_LENGTH = 31

Public Const MAXLONG = &H7FFFFFFF

Public Const THREAD_PRIORITY_LOWEST = THREAD_BASE_PRIORITY_MIN
Public Const THREAD_PRIORITY_BELOW_NORMAL = (THREAD_PRIORITY_LOWEST + 1)
Public Const THREAD_PRIORITY_NORMAL = 0
Public Const THREAD_PRIORITY_HIGHEST = THREAD_BASE_PRIORITY_MAX
Public Const THREAD_PRIORITY_ABOVE_NORMAL = (THREAD_PRIORITY_HIGHEST - 1)
Public Const THREAD_PRIORITY_ERROR_RETURN = (MAXLONG)
Public Const THREAD_PRIORITY_TIME_CRITICAL = THREAD_BASE_PRIORITY_LOWRT
Public Const THREAD_PRIORITY_IDLE = THREAD_BASE_PRIORITY_IDLE

Public Const TIME_ZONE_ID_INVALID = &HFFFFFFFF


Public Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Public Type MEMORYSTATUS
    dwLength As Long
    dwMemoryLoad As Long
    dwTotalPhys As Long
    dwAvailPhys As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual As Long
    dwAvailVirtual As Long
End Type
    
Public Type MEMORYSTATUSEX
    dwLength As Long
    dwMemoryLoad As Long
    ullTotalPhys As LARGE_INTEGER
    ullAvailPhys As LARGE_INTEGER
    ullTotalPageFile As LARGE_INTEGER
    ullAvailPageFile As LARGE_INTEGER
    ullTotalVirtual As LARGE_INTEGER
    ullAvailVirtual As LARGE_INTEGER
    ullAvailExtendedVirtual As LARGE_INTEGER
End Type

Public Type OVERLAPPED
    Internal As Long
    InternalHigh As Long
    Offset As Long
    OffsetHigh As Long
    hEvent As Long
End Type

Public Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Public Type SYSTEM_INFO
    dwOemID As Long 'Union
    'WORD wProcessorArchitecture
    'WORD wReserved

    dwPageSize As Long
    lpMinimumApplicationAddress As Long
    lpMaximumApplicationAddress As Long
    dwActiveProcessorMask As Long
    dwNumberOrfProcessors As Long
    dwProcessorType As Long
    dwAllocationGranularity As Long
    wProcessorLevel As Integer
    wProcessorRevision As Integer
End Type

Public Type SYSTEM_POWER_STATUS
    ACLineStatus As Byte
    BatteryFlag As Byte
    BatteryLifePercent As Byte
    Reserved1 As Byte
    BatteryLifeTime As Long
    BatteryFullLifeTime As Long
End Type

Public Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Public Type TIME_ZONE_INFORMATION
    Bias As Long
    StandardName As String * 64
    StandardDate As SYSTEMTIME
    StandardBias As Long
    DaylightName As String * 64
    DaylightDate As SYSTEMTIME
    DaylightBias As Long
End Type


Public Sub Errors(ByVal lngError As Long, ByVal strFunction As String, Optional ByRef errDescription As String, Optional ByVal NoMsgBox As Boolean)
    If lngError = 0 Then
        Failed strFunction
        Exit Sub
    End If
    
    errDescription = String$(2048, &H0)
    FormatMessage FORMAT_MESSAGE_FROM_SYSTEM, &H0, lngError, 0, errDescription, 2048, &H0
    
    If errDescription = "" Then
        errDescription = "No description available."
    End If
    
    If NoMsgBox = False Then
        If errMsg = True Then
             MessageBoxEx &H0, strFunction & Chr$(13) & Chr$(10) & Chr$(13) & Chr$(10) & errDescription, "Error", MB_OK Or MB_ICONWARNING Or MB_SETFOREGROUND, 0
        End If
    End If
End Sub

Public Sub Failed(ByVal strFunction As String)
    If errMsg = True Then
        If Err.LastDllError = 0 Then
            MessageBoxEx &H0, strFunction & Chr$(13) & Chr$(10) & Chr$(13) & Chr$(10) & "Failed", "Error", MB_OK Or MB_ICONWARNING Or MB_SETFOREGROUND, 0
        Else
            Errors Err.LastDllError, strFunction
        End If
    End If
End Sub

Public Function File_Exist(ByVal strFileName As String) As Boolean
    Dim hHandle As Long
    
    hHandle = CreateFile(strFileName, &H0, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal &H0, OPEN_EXISTING, &H0, &H0)
    If hHandle = INVALID_HANDLE_VALUE Then
        File_Exist = False
        Failed "CreateFile"
    Else
        File_Exist = True
    End If
    
    If CloseHandle(hHandle) = False Then Failed "CloseHandle"
End Function

Public Function FileSize_Handle(ByVal hHandle As Long) As Double
    If GetFileType(hHandle) = FILE_TYPE_DISK Then
        If Function_Exist("kernel32.dll", "GetFileSizeEx") = True Then
            Dim LARGE_INTEGER As LARGE_INTEGER
            If GetFileSizeEx(hHandle, LARGE_INTEGER) = False Then Failed "GetFileSizeEx"
            
            FileSize_Handle = CLargeInt(LARGE_INTEGER.LowPart, LARGE_INTEGER.HighPart)
        Else
            Dim Hi As Long
            Dim Lo As Long
            
            Lo = GetFileSize(hHandle, Hi): If Lo = INVALID_FILE_SIZE Then Failed "GetFileSize"
            
            FileSize_Handle = CLargeInt(Lo, Hi)
        End If
    End If
End Function

Public Function FileSize_Name(ByVal strFileName As String) As Double
    Dim Hi As Long
    Dim Lo As Long
    
    Dim hHandle As Long
    
    hHandle = CreateFile(strFileName, &H0, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal &H0, OPEN_EXISTING, &H0, &H0): If hHandle = INVALID_HANDLE_VALUE Then Failed "CreateFile"
    
    If Function_Exist("kernel32.dll", "GetFileSizeEx") = True Then
        Dim LARGE_INTEGER As LARGE_INTEGER
        If GetFileSizeEx(hHandle, LARGE_INTEGER) = False Then Failed "GetFileSizeEx"
        
        FileSize_Name = CLargeInt(LARGE_INTEGER.LowPart, LARGE_INTEGER.HighPart)
    Else
        Lo = GetFileSize(hHandle, Hi): If Lo = INVALID_FILE_SIZE Then Failed "GetFileSize"
        
        FileSize_Name = CLargeInt(Lo, Hi)
    End If
    
    If CloseHandle(hHandle) = False Then Failed "CloseHandle"
End Function

Public Function Function_Exist(ByVal strModule As String, ByVal strFunction As String) As Boolean
    Dim hHandle As Long
    
    hHandle = GetModuleHandle(strModule)
    If hHandle = &H0 Then
        Failed "GetModuleHandle"
        
        hHandle = LoadLibraryEx(strModule, &H0, &H0): If hHandle = &H0 Then Failed "LoadLibrary"
        
        If GetProcAddress(hHandle, strFunction) = &H0 Then
            Failed "GetProcAddress"
        Else
            Function_Exist = True
        End If
        
        If FreeLibrary(hHandle) = False Then Failed "FreeLibrary"
    Else
        If GetProcAddress(hHandle, strFunction) = &H0 Then
            Failed "GetProcAddress"
        Else
            Function_Exist = True
        End If
    End If
End Function

Public Function Get_ComputerName() As String
    Dim strComputerName As String
    strComputerName = String$(MAX_COMPUTERNAME_LENGTH + 1, &H0)
    
    If GetComputerName(strComputerName, MAX_COMPUTERNAME_LENGTH + 1) = False Then Failed "GetComputerName"
    Get_ComputerName = Fix_NullTermStr(strComputerName)
End Function

Public Function Get_DiskFreeSpaceEx(ByVal strDrive As String, ByRef dblFreeBytesAvailable As Double, ByRef dblTotalNumberOfBytes As Double, ByRef dblTotalNumberOfFreeBytes As Double) As Boolean
    If Function_Exist("kernel32.dll", "GetDiskFreeSpaceExA") = True Then
        Dim liFreeBytesAvailable As LARGE_INTEGER
        Dim liTotalNumberOfBytes As LARGE_INTEGER
        Dim liTotalNumberOfFreeBytes As LARGE_INTEGER
        
        If GetDiskFreeSpaceEx(strDrive, liFreeBytesAvailable, liTotalNumberOfBytes, liTotalNumberOfFreeBytes) = False Then
            Failed "GetDiskFreeSpaceEx"
            Get_DiskFreeSpaceEx = False
        Else
            Get_DiskFreeSpaceEx = True
        End If
        
        dblFreeBytesAvailable = CLargeInt(liFreeBytesAvailable.LowPart, liFreeBytesAvailable.HighPart)
        dblTotalNumberOfBytes = CLargeInt(liTotalNumberOfBytes.LowPart, liTotalNumberOfBytes.HighPart)
        dblTotalNumberOfFreeBytes = CLargeInt(liTotalNumberOfFreeBytes.LowPart, liTotalNumberOfFreeBytes.HighPart)
    End If
End Function

Public Function GetFilePointer_Name(ByVal strFileName As String) As Double
    Dim hFile As Long
    hFile = CreateFile(strFileName, GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal &H0, OPEN_EXISTING, &H0, &H0): If hFile = INVALID_HANDLE_VALUE Then Failed "CreateFile"
    
    Dim Lo As Long
    Dim Hi As Long
    
    Lo = SetFilePointer(hFile, 0, Hi, FILE_CURRENT): If Lo = INVALID_SET_FILE_POINTER Then Failed "SetFilePointer"
    GetFilePointer_Name = CLargeInt(Lo, Hi)
    
    If CloseHandle(hFile) = False Then Failed "CloseHandle"
End Function

Public Function Get_UserName() As String
    Dim strUserName As String
    strUserName = String$(UNLEN + 1, &H0)
    
    If GetUserName(strUserName, Len(strUserName)) = False Then Failed "GetUserName"
    Get_UserName = Fix_NullTermStr(strUserName)
End Function

Public Function GetFilePointer_Handle(ByVal hFile As Long) As Double
    If GetFileType(hFile) = FILE_TYPE_DISK Then
        Dim Lo As Long
        Dim Hi As Long
        
        Lo = SetFilePointer(hFile, 0, Hi, FILE_CURRENT): If Lo = INVALID_SET_FILE_POINTER Then Failed "SetFilePointer"
        
        GetFilePointer_Handle = CLargeInt(Lo, Hi)
    End If
End Function

Public Function PerformanceCounter() As Double
    Dim LARGE_INTEGER As LARGE_INTEGER
    
    If QueryPerformanceCounter(LARGE_INTEGER) = False Then Failed "QueryPerformanceCounter"
    PerformanceCounter = CLargeInt(LARGE_INTEGER.LowPart, LARGE_INTEGER.HighPart)
End Function

Public Function ReadFile_String(ByVal strFileName As String, ByVal lngLength As Long, ByVal lngStart As Long) As String
    If strFileName = "" Then Exit Function
    If lngLength = 0 Then Exit Function
    Select Case FileSize_Name(strFileName)
        Case 0: Exit Function
        Case lngStart: Exit Function
    End Select
    
    
    Dim hFile As Long
    Dim strBuffer As String
    Dim lngRead As Long
    
    hFile = CreateFile(strFileName, GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal &H0, OPEN_EXISTING, &H0, &H0): If hFile = INVALID_HANDLE_VALUE Then Failed "CreateFile"
    
    If GetFilePointer_Handle(hFile) <> lngStart Then
        If SetFilePointer(hFile, lngStart, 0, FILE_BEGIN) = INVALID_SET_FILE_POINTER Then Failed "SetFilePointer"
    End If
    
    strBuffer = String$(lngLength, &H0)
    
    If ReadFile(hFile, ByVal strBuffer, lngLength, lngRead, ByVal &H0) = False Then
        Failed "ReadFile"
    Else
        If lngRead = 0 Then Failed "ReadFile"
    End If
    
    If SetFilePointer(hFile, 0, 0, FILE_BEGIN) = INVALID_SET_FILE_POINTER Then Failed "SetFilePointer"
    If CloseHandle(hFile) = False Then Failed "CloseHandle"
    
    
    If lngRead < lngLength Then strBuffer = Left(strBuffer, lngRead)
    
    ReadFile_String = strBuffer
End Function

Public Sub Set_ComputerName(ByVal strComputerName As String)
    Dim strName As String
    strName = strComputerName
    
    If Len(strName) > MAX_COMPUTERNAME_LENGTH Then
        strName = Left$(strName, MAX_COMPUTERNAME_LENGTH)
    End If
    
    If SetComputerName(strName) = False Then Failed "SetComputerName"
End Sub

Public Sub SetFileSize_Handle(ByVal hFile As Long, ByVal lngFileLen As Long)
    If lngFileLen > -1 Then
        If SetFilePointer(hFile, lngFileLen, 0, FILE_BEGIN) = INVALID_SET_FILE_POINTER Then Failed "SetFilePointer"
        If SetEndOfFile(hFile) = False Then Failed "SetEndOfFile"
    End If
End Sub

Public Sub SetFileSize_Name(ByVal strFileName As String, ByVal lngFileLen As Long)
    Dim hFile As Long
    
    hFile = CreateFile(strFileName, GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal &H0, OPEN_EXISTING, &H0, &H0): If hFile = INVALID_HANDLE_VALUE Then Failed "CreateFile"
    
    If lngFileLen > 0 Then
        If SetFilePointer(hFile, lngFileLen, 0, FILE_BEGIN) = INVALID_SET_FILE_POINTER Then Failed "SetFilePointer"
        If SetEndOfFile(hFile) = False Then Failed "SetEndOfFile"
    End If
    
    If CloseHandle(hFile) = False Then Failed "CloseHandle"
End Sub

Public Sub WriteFile_String(ByVal strFileName As String, ByVal strData As String, ByVal lngStart As Long, ByVal lngFlags As Long)
    If strFileName = "" Then Exit Sub
    If strData = "" Then Exit Sub
    
    
    Dim hFile As Long
    Dim lngWrite As Long
    
    hFile = CreateFile(strFileName, GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal &H0, lngFlags, &H0, &H0): If hFile = INVALID_HANDLE_VALUE Then Failed "CreateFile"
    
    If GetFilePointer_Handle(hFile) <> lngStart Then
        If SetFilePointer(hFile, lngStart, 0, FILE_BEGIN) = INVALID_SET_FILE_POINTER Then Failed "SetFilePointer"
    End If
    
    If WriteFile(hFile, ByVal strData, Len(strData), lngWrite, ByVal &H0) = False Then Failed "WriteFile"
    If SetFilePointer(hFile, 0, 0, FILE_BEGIN) = INVALID_SET_FILE_POINTER Then Failed "SetFilePointer"
    If CloseHandle(hFile) = False Then Failed "CloseHandle"
End Sub
