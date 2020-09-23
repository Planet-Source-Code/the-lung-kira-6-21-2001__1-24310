VERSION 5.00
Begin VB.Form frmProcesses 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Processes"
   ClientHeight    =   4155
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6135
   Icon            =   "frmProcesses.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4155
   ScaleWidth      =   6135
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdTerminate 
      Caption         =   "Terminate"
      Height          =   350
      Left            =   2040
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   3720
      Width           =   975
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   350
      Left            =   2040
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   3240
      Width           =   975
   End
   Begin VB.ComboBox cboPriority 
      Height          =   315
      Left            =   120
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2880
      Width           =   2895
   End
   Begin VB.TextBox txtUsage 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox txtThreads 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox txtPrimaryBaseClass 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox txtParentProcessID 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox txtExpectedVersion 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox txtExeFile 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox txtAffinityMask 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox txtUserObjects 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4800
      TabIndex        =   24
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox txtGDIObjects 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4800
      TabIndex        =   22
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox txtOtherTransfer 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4800
      TabIndex        =   36
      Top             =   3720
      Width           =   1215
   End
   Begin VB.TextBox txtWriteTransfer 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4800
      TabIndex        =   34
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox txtReadTransfer 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4800
      TabIndex        =   32
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox txtOtherOperation 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4800
      TabIndex        =   30
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox txtWriteOperation 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4800
      TabIndex        =   28
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox txtReadOperation 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4800
      TabIndex        =   26
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   350
      Left            =   2040
      TabIndex        =   2
      Top             =   2160
      Width           =   975
   End
   Begin VB.ListBox lstProcess 
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1680
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   2895
   End
   Begin VB.Label lblPriority 
      Caption         =   "Priority"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label lblUsage 
      Caption         =   "Usage"
      Height          =   255
      Left            =   3240
      TabIndex        =   19
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label lblThreads 
      Caption         =   "Threads"
      Height          =   255
      Left            =   3240
      TabIndex        =   17
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label lblParentProcessID 
      Caption         =   "Parent Process ID"
      Height          =   255
      Left            =   3240
      TabIndex        =   13
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label lblPrimaryBaseClass 
      Caption         =   "Primary Base Class"
      Height          =   255
      Left            =   3240
      TabIndex        =   15
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label lblExeFile 
      Caption         =   "Exe File"
      Height          =   255
      Left            =   3240
      TabIndex        =   9
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label lblExpectedVersion 
      Caption         =   "Expected Version"
      Height          =   255
      Left            =   3240
      TabIndex        =   11
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label lblAffinityMask 
      Caption         =   "Affinity Mask"
      Height          =   255
      Left            =   3240
      TabIndex        =   7
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label lblGDIObjects 
      Caption         =   "GDI Objects"
      Height          =   255
      Left            =   3240
      TabIndex        =   21
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label lblUserObjects 
      Caption         =   "User Objects"
      Height          =   255
      Left            =   3240
      TabIndex        =   23
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label lblOtherTransfer 
      Caption         =   "Other Transfer"
      Height          =   255
      Left            =   3240
      TabIndex        =   35
      Top             =   3720
      Width           =   1335
   End
   Begin VB.Label lblWriteTransfer 
      Caption         =   "Write Transfer"
      Height          =   255
      Left            =   3240
      TabIndex        =   33
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label lblReadTransfer 
      Caption         =   "Read Transfer"
      Height          =   255
      Left            =   3240
      TabIndex        =   31
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Label lblOtherOperation 
      Caption         =   "Other Operation"
      Height          =   255
      Left            =   3240
      TabIndex        =   29
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label lblWriteOperation 
      Caption         =   "Write Operation"
      Height          =   255
      Left            =   3240
      TabIndex        =   27
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label lblReadOperation 
      Caption         =   "Read Operation"
      Height          =   255
      Left            =   3240
      TabIndex        =   25
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label lblProcess 
      Caption         =   "Process"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmProcesses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Process() As PROCESSENTRY32
Dim lngProcess As Long

Private Sub cmdApply_Click()
    If cboPriority.ListIndex = -1 Then Exit Sub
    If lstProcess.ListIndex = -1 Then Exit Sub
    
    Dim hProcess As Long
    Dim lngPriority As Long
    
    hProcess = OpenProcess(PROCESS_SET_INFORMATION, False, Process(lstProcess.ListIndex).th32ProcessID): If hProcess = &H0 Then Failed "OpenProcess"
    
    Select Case cboPriority.ListIndex
        Case 0: lngPriority = BELOW_NORMAL_PRIORITY_CLASS
        Case 1: lngPriority = NORMAL_PRIORITY_CLASS
        Case 2: lngPriority = ABOVE_NORMAL_PRIORITY_CLASS
        Case 3: lngPriority = REALTIME_PRIORITY_CLASS
        Case 4: lngPriority = IDLE_PRIORITY_CLASS
    End Select
    
    If SetPriorityClass(hProcess, lngPriority) = False Then Failed "SetPriorityClass"
    If CloseHandle(hProcess) = False Then Failed "CloseHandle"
End Sub

Private Sub cmdRefresh_Click()
    lstProcess.Clear
    lngProcess = 0
    Erase Process()
    
    lngProcess = Process32_Enum(Process())
    
    Dim lngIncrement As Long
    For lngIncrement = 0 To lngProcess
        lstProcess.AddItem CStr(Process(lngIncrement).th32ProcessID)
    Next lngIncrement
End Sub

Private Sub cmdTerminate_Click()
    If lstProcess.ListIndex = -1 Then Exit Sub
    
    Dim hProcess As Long
    Dim lngExitCode As Long
    
    hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_TERMINATE, False, Process(lstProcess.ListIndex).th32ProcessID): If hProcess = &H0 Then Failed "OpenProcess"
    
    If GetExitCodeProcess(hProcess, lngExitCode) = False Then Failed "GetExitCodeProcess"
    If TerminateProcess(hProcess, lngExitCode) = False Then Failed "TerminateProcess"
    
    If CloseHandle(hProcess) = False Then Failed "CloseHandle"
End Sub

Private Sub Form_Load()
    With cboPriority
        .AddItem "Below Normal"
        .AddItem "Normal"
        .AddItem "Above Normal"
        .AddItem "Real Time"
        .AddItem "Idle"
    End With
    
    cmdRefresh_Click
    

    If Function_Exist("kernel32.dll", "CreateToolhelp32Snapshot") = False Then
        lblProcess.Enabled = False
        lstProcess.Enabled = False
        lblPriority.Enabled = False
        cboPriority.Enabled = False
        cmdApply.Enabled = False
        lblAffinityMask.Enabled = False
        lblExeFile.Enabled = False
        lblExpectedVersion.Enabled = False
        lblParentProcessID.Enabled = False
        lblPrimaryBaseClass.Enabled = False
        lblThreads.Enabled = False
        lblUsage.Enabled = False
        cmdRefresh.Enabled = False
        cmdTerminate.Enabled = False
    End If
    If Function_Exist("user32.dll", "GetGuiResources") = False Then
        lblGDIObjects.Enabled = False
        lblUserObjects.Enabled = False
    End If
    If Function_Exist("kernel32.dll", "GetProcessIoCounters") = False Then
        lblReadOperation.Enabled = False
        lblWriteOperation.Enabled = False
        lblOtherOperation.Enabled = False
        lblReadTransfer.Enabled = False
        lblWriteTransfer.Enabled = False
        lblOtherTransfer.Enabled = False
    End If
End Sub

Private Sub lstProcess_Click()
    Dim hProcess As Long
    hProcess = OpenProcess(PROCESS_QUERY_INFORMATION, False, Process(lstProcess.ListIndex).th32ProcessID): If hProcess = &H0 Then Failed "OpenProcess"
    
    
    Select Case GetPriorityClass(hProcess)
        Case BELOW_NORMAL_PRIORITY_CLASS: cboPriority.ListIndex = 0
        Case NORMAL_PRIORITY_CLASS: cboPriority.ListIndex = 1
        Case ABOVE_NORMAL_PRIORITY_CLASS: cboPriority.ListIndex = 2
        Case REALTIME_PRIORITY_CLASS: cboPriority.ListIndex = 3
        Case IDLE_PRIORITY_CLASS: cboPriority.ListIndex = 4
        Case 0: Failed "GetPriorityClass"
        Case Else: cboPriority.Text = CStr(GetPriorityClass(hProcess))
    End Select
    
    
    Dim lngProcessAffinityMask As Long
    Dim lngSystemAffinityMask As Long
    If GetProcessAffinityMask(hProcess, lngProcessAffinityMask, lngSystemAffinityMask) = False Then Failed "GetProcessAffinityMask"
    txtAffinityMask.Text = Right$(String$(32, "0") & ltoa_(lngProcessAffinityMask, 2), 32)
    

    With Process(lstProcess.ListIndex)
        txtExeFile.Text = .szExeFile
        
        Dim lngExpectedVersion As Long
        lngExpectedVersion = GetProcessVersion(.th32ProcessID): If lngExpectedVersion = 0 Then Failed "GetProcessVersion"
        txtExpectedVersion.Text = CStr(HI_WORD(lngExpectedVersion)) & "." & CStr(LO_WORD(lngExpectedVersion))
        
        txtParentProcessID.Text = CStr(.th32ParentProcessID)
        txtPrimaryBaseClass.Text = CStr(.pcPriClassBase)
        txtThreads.Text = CStr(.cntThreads)
        txtUsage.Text = CStr(.cntUsage)
    End With
    
    
    If Function_Exist("user32.dll", "GetGuiResources") = True Then
        Dim lngValue As Long
        
        lngValue = GetGuiResources(hProcess, GR_GDIOBJECTS): If lngValue = 0 Then Failed "GetGuiResources"
        txtGDIObjects.Text = CStr(lngValue)
        
        lngValue = GetGuiResources(hProcess, GR_USEROBJECTS): If lngValue = 0 Then Failed "GetGuiResources"
        txtUserObjects.Text = CStr(lngValue)
    End If
    
    If Function_Exist("kernel32.dll", "GetProcessIoCounters") = True Then
        Dim IO_COUNTERS As IO_COUNTERS
        If GetProcessIoCounters(hProcess, IO_COUNTERS) = False Then Failed "GetProcessIoCounters"
        
        With IO_COUNTERS
            txtReadOperation.Text = CLargeInt(.ReadOperationCount.LowPart, .ReadOperationCount.HighPart)
            txtWriteOperation.Text = CLargeInt(.WriteOperationCount.LowPart, .WriteOperationCount.HighPart)
            txtOtherOperation.Text = CLargeInt(.OtherOperationCount.LowPart, .OtherOperationCount.HighPart)
            txtReadTransfer.Text = CLargeInt(.ReadTransferCount.LowPart, .ReadTransferCount.HighPart)
            txtWriteTransfer.Text = CLargeInt(.WriteTransferCount.LowPart, .WriteTransferCount.HighPart)
            txtOtherTransfer.Text = CLargeInt(.OtherTransferCount.LowPart, .OtherTransferCount.HighPart)
        End With
    End If
    
    If CloseHandle(hProcess) = False Then Failed "CloseHandle"
End Sub
