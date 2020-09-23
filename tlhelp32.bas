Attribute VB_Name = "tlhelp32"
Option Explicit


Public Declare Function CreateToolhelp32Snapshot Lib "kernel32.dll" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
Public Declare Function Heap32ListFirst Lib "kernel32.dll" (ByVal hSnapShot As Long, ByRef lphl As HEAPLIST32) As Boolean
Public Declare Function Heap32ListNext Lib "kernel32.dll" (ByVal hSnapShot As Long, ByRef lphl As HEAPLIST32) As Boolean
Public Declare Function Heap32First Lib "kernel32.dll" (ByRef lphe As HEAPENTRY32, ByVal th32ProcessID As Long, ByVal th32HeapID As Long) As Boolean
Public Declare Function Heap32Next Lib "kernel32.dll" (ByRef lphe As HEAPENTRY32) As Boolean
Public Declare Function Module32First Lib "kernel32.dll" (ByVal hSnapShot As Long, ByRef lpme As MODULEENTRY32) As Boolean
Public Declare Function Module32Next Lib "kernel32.dll" (ByVal hSnapShot As Long, ByRef lpme As MODULEENTRY32) As Boolean
Public Declare Function Process32First Lib "kernel32.dll" (ByVal hSnapShot As Long, ByRef lppe As PROCESSENTRY32) As Boolean
Public Declare Function Process32Next Lib "kernel32.dll" (ByVal hSnapShot As Long, ByRef lppe As PROCESSENTRY32) As Boolean
Public Declare Function Thread32First Lib "kernel32.dll" (ByVal hSnapShot As Long, ByRef lpte As THREADENTRY32) As Boolean
Public Declare Function Thread32Next Lib "kernel32.dll" (ByVal hSnapShot As Long, ByRef lpte As THREADENTRY32) As Boolean


Public Const HF32_DEFAULT = 1
Public Const HF32_SHARED = 2

Public Const LF32_FIXED = &H1
Public Const LF32_FREE = &H2
Public Const LF32_MOVEABLE = &H4

Public Const MAX_MODULE_NAME32 = 255

Public Const TH32CS_SNAPHEAPLIST = &H1
Public Const TH32CS_SNAPPROCESS = &H2
Public Const TH32CS_SNAPTHREAD = &H4
Public Const TH32CS_SNAPMODULE = &H8
Public Const TH32CS_SNAPALL = (TH32CS_SNAPHEAPLIST Or TH32CS_SNAPPROCESS Or TH32CS_SNAPTHREAD Or TH32CS_SNAPMODULE)
Public Const TH32CS_INHERIT = &H80000000
    
    
Public Type HEAPENTRY32
    dwSize As Long
    hHandle As Long
    dwAddress As Long
    dwBlockSize As Long
    dwFlags As Long
    dwLockCount As Long
    dwResvd As Long
    th32ProcessID As Long
    th32HeapID As Long
End Type

Public Type HEAPLIST32
    dwSize As Long
    th32ProcessID As Long
    th32HeapID As Long
    dwFlags As Long
End Type

Public Type MODULEENTRY32
    dwSize As Long
    th32ModuleID As Long
    th32ProcessID As Long
    GlblcntUsage As Long
    ProccntUsage As Long
    modBaseAddr As Long
    modBaseSize As Long
    hModule As Long
    szModule As String * 256    'MAX_MODULE_NAME32 + 1
    szExePath As String * MAX_PATH
End Type

Public Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID  As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * MAX_PATH
End Type

Public Type THREADENTRY32
    dwSize As Long
    cntUsage As Long
    th32ThreadID As Long
    th32OwnerProcessID As Long
    tpBasePri As Long
    tpDeltaPri As Long
    dwFlags As Long
End Type


Public Function Heap32_Enum(ByRef Heap() As HEAPENTRY32, ByVal lngProcessID As Long, ByVal lngHeapID As Long) As Long
    If Function_Exist("kernel32.dll", "CreateToolhelp32Snapshot") = True Then
        Dim HEAPENTRY32 As HEAPENTRY32
        Dim lngHeap As Long
        
        
        HEAPENTRY32.dwSize = Len(HEAPENTRY32)
        If Heap32First(HEAPENTRY32, lngProcessID, lngHeapID) = False Then
            Heap32_Enum = -1
            Failed "Heap32First"
            
            Exit Function
        Else
            ReDim Heap(lngHeap)
            Heap(lngHeap) = HEAPENTRY32
        End If
        
        Do
            If Heap32Next(HEAPENTRY32) = False Then
                Exit Do
            Else
                lngHeap = lngHeap + 1
                ReDim Preserve Heap(lngHeap)
                Heap(lngHeap) = HEAPENTRY32
            End If
        Loop
        
        
        Heap32_Enum = lngHeap
    End If
End Function

Public Function Heap32List_Enum(ByRef HeapList() As HEAPLIST32, Optional ByVal lngProcessID As Long) As Long
    If Function_Exist("kernel32.dll", "CreateToolhelp32Snapshot") = True Then
        Dim HEAPLIST32 As HEAPLIST32
        Dim hSnapShot As Long
        Dim lngHeapList As Long
        
        hSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPHEAPLIST, lngProcessID): If hSnapShot = -1 Then Failed "CreateToolhelp32Snapshot"
        
        HEAPLIST32.dwSize = Len(HEAPLIST32)
        If Heap32ListFirst(hSnapShot, HEAPLIST32) = False Then
            Heap32List_Enum = -1
            Failed "Heap32ListFirst"
            
            If CloseHandle(hSnapShot) = False Then Failed "CloseHandle"
            Exit Function
        Else
            ReDim HeapList(lngHeapList)
            HeapList(lngHeapList) = HEAPLIST32
        End If
        
        Do
            If Heap32ListNext(hSnapShot, HEAPLIST32) = False Then
                Exit Do
            Else
                lngHeapList = lngHeapList + 1
                ReDim Preserve HeapList(lngHeapList)
                HeapList(lngHeapList) = HEAPLIST32
            End If
        Loop
        
        If CloseHandle(hSnapShot) = False Then Failed "CloseHandle"
        
        Heap32List_Enum = lngHeapList
    End If
End Function

Public Function Module32_Enum(ByRef Module() As MODULEENTRY32, Optional ByVal lngProcessID As Long) As Long
    If Function_Exist("kernel32.dll", "CreateToolhelp32Snapshot") = True Then
        Dim MODULEENTRY32 As MODULEENTRY32
        Dim hSnapShot As Long
        Dim lngModule As Long
        
        hSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPMODULE, lngProcessID): If hSnapShot = -1 Then Failed "CreateToolhelp32Snapshot"
        
        MODULEENTRY32.dwSize = Len(MODULEENTRY32)
        If Module32First(hSnapShot, MODULEENTRY32) = False Then
            Module32_Enum = -1
            Failed "Module32First"
            
            If CloseHandle(hSnapShot) = False Then Failed "CloseHandle"
            Exit Function
        Else
            ReDim Module(lngModule)
            Module(lngModule) = MODULEENTRY32
        End If
        
        Do
            If Module32Next(hSnapShot, MODULEENTRY32) = False Then
                Exit Do
            Else
                lngModule = lngModule + 1
                ReDim Preserve Module(lngModule)
                Module(lngModule) = MODULEENTRY32
            End If
        Loop
        
        If CloseHandle(hSnapShot) = False Then Failed "CloseHandle"
        
        Module32_Enum = lngModule
    End If
End Function

Public Function Process32_Enum(ByRef Process() As PROCESSENTRY32) As Long
    If Function_Exist("kernel32.dll", "CreateToolhelp32Snapshot") = True Then
        Dim PROCESSENTRY32 As PROCESSENTRY32
        Dim hSnapShot As Long
        Dim lngProcess As Long
        
        hSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, &H0): If hSnapShot = -1 Then Failed "CreateToolhelp32Snapshot"
        
        PROCESSENTRY32.dwSize = Len(PROCESSENTRY32)
        If Process32First(hSnapShot, PROCESSENTRY32) = False Then
            Process32_Enum = -1
            Failed "Process32First"
            
            If CloseHandle(hSnapShot) = False Then Failed "CloseHandle"
            Exit Function
        Else
            ReDim Process(lngProcess)
            Process(lngProcess) = PROCESSENTRY32
        End If
    
        Do
            If Process32Next(hSnapShot, PROCESSENTRY32) = False Then
                Exit Do
            Else
                lngProcess = lngProcess + 1
                ReDim Preserve Process(lngProcess)
                Process(lngProcess) = PROCESSENTRY32
            End If
        Loop
        
        If CloseHandle(hSnapShot) = False Then Failed "CloseHandle"
        
        Process32_Enum = lngProcess
    End If
End Function

Public Function Thread32_Enum(ByRef Thread() As THREADENTRY32) As Long
    If Function_Exist("kernel32.dll", "CreateToolhelp32Snapshot") = True Then
        Dim THREADENTRY32 As THREADENTRY32
        Dim hSnapShot As Long
        Dim lngThread As Long
        
        hSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPTHREAD, &H0): If hSnapShot = -1 Then Failed "CreateToolhelp32Snapshot"
        
        THREADENTRY32.dwSize = Len(THREADENTRY32)
        If Thread32First(hSnapShot, THREADENTRY32) = False Then
            Thread32_Enum = -1
            Failed "Thread32First"
            
            If CloseHandle(hSnapShot) = False Then Failed "CloseHandle"
            Exit Function
        Else
            ReDim Thread(lngThread)
            Thread(lngThread) = THREADENTRY32
        End If
        
        Do
            If Thread32Next(hSnapShot, THREADENTRY32) = False Then
                Exit Do
            Else
                lngThread = lngThread + 1
                ReDim Preserve Thread(lngThread)
                Thread(lngThread) = THREADENTRY32
            End If
        Loop
        
        If CloseHandle(hSnapShot) = False Then Failed "CloseHandle"
        
        Thread32_Enum = lngThread
    End If
End Function
