VERSION 5.00
Begin VB.Form frmThreads 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Threads"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6255
   Icon            =   "frmThreads.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   6255
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdTerminate 
      Caption         =   "Terminate"
      Height          =   350
      Left            =   5160
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   3360
      Width           =   975
   End
   Begin VB.CommandButton cmdSuspend 
      Caption         =   "Suspend"
      Height          =   350
      Left            =   5160
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton cmdResume 
      Caption         =   "Resume"
      Height          =   350
      Left            =   5160
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   2760
      Width           =   975
   End
   Begin VB.TextBox txtSuspendCount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   2760
      Width           =   1335
   End
   Begin VB.TextBox txtIdealProcessor 
      Height          =   285
      Left            =   4800
      TabIndex        =   8
      Text            =   "1"
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox txtUsage 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1920
      Width           =   1335
   End
   Begin VB.TextBox txtDeltaPriority 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   350
      Left            =   5160
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1200
      Width           =   975
   End
   Begin VB.ComboBox cboPriority 
      Height          =   315
      Left            =   3240
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   360
      Width           =   2895
   End
   Begin VB.ListBox lstThread 
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
      TabIndex        =   4
      Top             =   2880
      Width           =   2895
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
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   350
      Left            =   2040
      TabIndex        =   2
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label lblSuspendCount 
      Caption         =   "Suspend Count"
      Height          =   255
      Left            =   3240
      TabIndex        =   14
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label lblIdealProcessor 
      Caption         =   "Ideal Processor"
      Height          =   255
      Left            =   3240
      TabIndex        =   7
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label lblUsage 
      Caption         =   "Usage"
      Height          =   255
      Left            =   3240
      TabIndex        =   12
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label lblDeltaPriority 
      Caption         =   "Delta Priority"
      Height          =   255
      Left            =   3240
      TabIndex        =   10
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label lblPriority 
      Caption         =   "Priority"
      Height          =   255
      Left            =   3240
      TabIndex        =   5
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label lblThread 
      Caption         =   "Thread ID"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label lblProcess 
      Caption         =   "Process ID"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmThreads"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Process() As PROCESSENTRY32
Dim lngProcess As Long
Dim Thread() As THREADENTRY32
Dim lngThread As Long

Private Sub cmdApply_Click()
    If lstThread.ListIndex = -1 Then Exit Sub
    
    Dim hThread As Long
    Dim lngPriority As Long
    
    hThread = OpenThread(THREAD_SET_INFORMATION, False, lstThread.List(lstThread.ListIndex)): If hThread = &H0 Then Failed "OpenThread"
    
    
    If cboPriority.ListIndex > -1 Then
        Select Case cboPriority.ListIndex
            Case 0: lngPriority = THREAD_PRIORITY_TIME_CRITICAL
            Case 1: lngPriority = THREAD_PRIORITY_HIGHEST
            Case 2: lngPriority = THREAD_PRIORITY_ABOVE_NORMAL
            Case 3: lngPriority = THREAD_PRIORITY_NORMAL
            Case 4: lngPriority = THREAD_PRIORITY_BELOW_NORMAL
            Case 5: lngPriority = THREAD_PRIORITY_LOWEST
            Case 6: lngPriority = THREAD_PRIORITY_IDLE
        End Select
        
        If SetThreadPriority(hThread, lngPriority) = False Then Failed "SetThreadPriority"
    End If
    
    If Function_Exist("kernel32.dll", "SetThreadIdealProcessor") = True Then
        If SetThreadIdealProcessor(hThread, Val(txtIdealProcessor.Text)) = -1 Then Failed "SetThreadIdealProcessor"
    End If
    
        
    If CloseHandle(hThread) = False Then Failed "CloseHandle"
End Sub

Private Sub cmdRefresh_Click()
    lstProcess.Clear
    lstThread.Clear
    lngProcess = 0
    Erase Process()
    
    lngProcess = Process32_Enum(Process())
    
    Dim lngIncrement As Long
    For lngIncrement = 0 To lngProcess
        lstProcess.AddItem CStr(Process(lngIncrement).th32ProcessID)
    Next lngIncrement
    
    
    cboPriority.ListIndex = -1
    txtDeltaPriority.Text = ""
    txtUsage.Text = ""
    txtSuspendCount.Text = ""
End Sub

Private Sub cmdResume_Click()
    If lstThread.ListIndex = -1 Then Exit Sub
    
    Dim hThread As Long
    Dim lngSuspendCount As Long
    
    hThread = OpenThread(THREAD_SUSPEND_RESUME, False, lstThread.List(lstThread.ListIndex)): If hThread = &H0 Then Failed "OpenThread"
    
    lngSuspendCount = ResumeThread(hThread): If lngSuspendCount = -1 Then Failed "ResumeThread"
    txtSuspendCount.Text = CStr(lngSuspendCount)
    
    If CloseHandle(hThread) = False Then Failed "CloseHandle"
End Sub

Private Sub cmdSuspend_Click()
    If lstThread.ListIndex = -1 Then Exit Sub
    
    Dim hThread As Long
    Dim lngSuspendCount As Long
    
    hThread = OpenThread(THREAD_SUSPEND_RESUME, False, lstThread.List(lstThread.ListIndex)): If hThread = &H0 Then Failed "OpenThread"
    
    lngSuspendCount = SuspendThread(hThread): If lngSuspendCount = -1 Then Failed "SuspendThread"
    txtSuspendCount.Text = CStr(lngSuspendCount)
    
    If CloseHandle(hThread) = False Then Failed "CloseHandle"
End Sub

Private Sub cmdTerminate_Click()
    If lstThread.ListIndex = -1 Then Exit Sub
    
    Dim hThread As Long
    Dim lngExitCode As Long
    
    hThread = OpenThread(THREAD_QUERY_INFORMATION Or THREAD_TERMINATE, False, lstThread.List(lstThread.ListIndex)): If hThread = &H0 Then Failed "OpenThread"
    If GetExitCodeThread(hThread, lngExitCode) = False Then Failed "GetExitCodeThread"
    If TerminateThread(hThread, lngExitCode) = False Then Failed "TerminateThread"
    
    If CloseHandle(hThread) = False Then Failed "CloseHandle"
End Sub

Private Sub Form_Load()
    With cboPriority
        .AddItem "Time Critical"
        .AddItem "Highest"
        .AddItem "Above Normal"
        .AddItem "Normal"
        .AddItem "Below Normal"
        .AddItem "Lowest"
        .AddItem "Idle"
    End With
    
    cmdRefresh_Click
    
    
    If Function_Exist("kernel32.dll", "CreateToolhelp32Snapshot") = False Then
        lblProcess.Enabled = False
        lstProcess.Enabled = False
        cmdRefresh.Enabled = False
        lblThread.Enabled = False
        lstThread.Enabled = False
        lblDeltaPriority.Enabled = False
        txtDeltaPriority.Enabled = False
        lblUsage.Enabled = False
        txtUsage.Enabled = False
    End If
    If Function_Exist("kernel32.dll", "OpenThread") = False Then
        lblPriority.Enabled = False
        cboPriority.Enabled = False
        lblIdealProcessor.Enabled = False
        txtIdealProcessor.Enabled = False
        
        cmdApply.Enabled = False
        cmdResume.Enabled = False
        cmdSuspend.Enabled = False
        cmdTerminate.Enabled = False
    End If
End Sub

Private Sub lstProcess_Click()
    lstThread.Clear
    lngThread = 0
    Erase Thread()
    
    lngThread = Thread32_Enum(Thread())
    
    Dim lngIncrement As Long
    For lngIncrement = 0 To lngThread
        If Thread(lngIncrement).th32OwnerProcessID = lstProcess.List(lstProcess.ListIndex) Then
            lstThread.AddItem CStr(Thread(lngIncrement).th32ThreadID)
        End If
    Next lngIncrement
    
    
    cboPriority.ListIndex = -1
    txtDeltaPriority.Text = ""
    txtUsage.Text = ""
    txtSuspendCount.Text = ""
End Sub

Private Sub lstThread_Click()
    If Function_Exist("kernel32.dll", "OpenThread") = True Then
        Dim hThread As Long
        hThread = OpenThread(THREAD_QUERY_INFORMATION, False, lstThread.List(lstThread.ListIndex)): If hThread = &H0 Then Failed "OpenThread"
        
        
        Dim lngValue As Long
        lngValue = GetThreadPriority(hThread)
        Select Case lngValue
            Case THREAD_PRIORITY_LOWEST: cboPriority.ListIndex = 0
            Case THREAD_PRIORITY_BELOW_NORMAL: cboPriority.ListIndex = 1
            Case THREAD_PRIORITY_NORMAL: cboPriority.ListIndex = 2
            Case THREAD_PRIORITY_HIGHEST: cboPriority.ListIndex = 3
            Case THREAD_PRIORITY_ABOVE_NORMAL: cboPriority.ListIndex = 4
            Case THREAD_PRIORITY_ERROR_RETURN: cboPriority.ListIndex = 5
            Case THREAD_PRIORITY_TIME_CRITICAL: cboPriority.ListIndex = 6
            Case THREAD_PRIORITY_IDLE: cboPriority.ListIndex = 7
            Case THREAD_PRIORITY_ERROR_RETURN: Failed "GetThreadPriority"
            Case Else: cboPriority.Text = CStr(lngValue)
        End Select
        
        If CloseHandle(hThread) = False Then Failed "CloseHandle"
    Else
        Select Case Thread(lstThread.ListIndex).tpBasePri
            Case THREAD_PRIORITY_LOWEST: cboPriority.ListIndex = 0
            Case THREAD_PRIORITY_BELOW_NORMAL: cboPriority.ListIndex = 1
            Case THREAD_PRIORITY_NORMAL: cboPriority.ListIndex = 2
            Case THREAD_PRIORITY_HIGHEST: cboPriority.ListIndex = 3
            Case THREAD_PRIORITY_ABOVE_NORMAL: cboPriority.ListIndex = 4
            Case THREAD_PRIORITY_ERROR_RETURN: cboPriority.ListIndex = 5
            Case THREAD_PRIORITY_TIME_CRITICAL: cboPriority.ListIndex = 6
            Case THREAD_PRIORITY_IDLE: cboPriority.ListIndex = 7
            Case Else: cboPriority.Text = CStr(Thread(lstThread.ListIndex).tpBasePri)
        End Select
    End If
    
        
    With Thread(lstThread.ListIndex)
        txtDeltaPriority.Text = CStr(.tpDeltaPri)
        txtUsage.Text = CStr(.cntUsage)
    End With
End Sub

Private Sub txtIdealProcessor_Change()
    txtIdealProcessor.Text = CStr(Val(Rem_NonNumeric_Chr(txtIdealProcessor.Text)))
    If Val(txtIdealProcessor.Text) < 1 Then txtIdealProcessor.Text = "1"
    If Val(txtIdealProcessor.Text) > MAXIMUM_PROCESSORS Then txtIdealProcessor.Text = CStr(MAXIMUM_PROCESSORS)
End Sub
