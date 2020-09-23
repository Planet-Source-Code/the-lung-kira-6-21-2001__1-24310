VERSION 5.00
Begin VB.Form frmMemoryStatus 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Memory Status"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5295
   Icon            =   "frmMemoryStatus.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   5295
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtAvailableExtendedVirtualMemoryPercentage 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4680
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   1800
      Width           =   495
   End
   Begin VB.TextBox txtAvailableExtendedVirtualMemory 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   1800
      Width           =   1575
   End
   Begin VB.TextBox txtAvailableVirtualMemory 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   1560
      Width           =   1575
   End
   Begin VB.TextBox txtAvailablePageFileMemory 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   960
      Width           =   1575
   End
   Begin VB.TextBox txtAvailablePhysicalMemory 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   360
      Width           =   1575
   End
   Begin VB.Timer timerMemoryStatus 
      Interval        =   1000
      Left            =   1440
      Top             =   2640
   End
   Begin VB.ComboBox cboRound 
      Height          =   315
      Left            =   2880
      TabIndex        =   23
      Top             =   3000
      Width           =   2295
   End
   Begin VB.ComboBox cboOutput 
      Height          =   315
      Left            =   2880
      TabIndex        =   21
      Top             =   2640
      Width           =   2295
   End
   Begin VB.TextBox txtTotalPhysicalMemory 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox txtAvailablePhysicalMemoryPercentage 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4680
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   360
      Width           =   495
   End
   Begin VB.TextBox txtTotalPageFileMemory 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   720
      Width           =   1575
   End
   Begin VB.TextBox txtAvailablePageFileMemoryPercentage 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4680
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   960
      Width           =   495
   End
   Begin VB.TextBox txtTotalVirtualMemory 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   1320
      Width           =   1575
   End
   Begin VB.TextBox txtAvailableVirtualMemoryPercentage 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4680
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   1560
      Width           =   495
   End
   Begin VB.TextBox txtMemoryLoad 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4680
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label lblRound 
      Caption         =   "Round"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label lblOutput 
      Caption         =   "Output"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label lblTotalPhysicalMemory 
      Caption         =   "Total Physical Memory"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label lblAvailablePhysicalMemory 
      Caption         =   "Available Physical Memory"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   2535
   End
   Begin VB.Label lblTotalPageFileMemory 
      Caption         =   "Total Page File Memory"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   2535
   End
   Begin VB.Label lblAvailablePageFileMemory 
      Caption         =   "Available Page File Memory"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   2535
   End
   Begin VB.Label lblTotalVirtualMemory 
      Caption         =   "Total Virtual Memory"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1320
      Width           =   2535
   End
   Begin VB.Label lblAvailableVirtualMemory 
      Caption         =   "Available Virtual Memory"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1560
      Width           =   2535
   End
   Begin VB.Label lblAvailableExtendedVirtualMemory 
      Caption         =   "Available Extended Virtual Memory"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   1800
      Width           =   2535
   End
   Begin VB.Label lblMemoryLoad 
      Caption         =   "Memory Load"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   2160
      Width           =   2535
   End
End
Attribute VB_Name = "frmMemoryStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Ex As Boolean

Private Sub Form_Load()
    With cboOutput
        .AddItem "Bytes"
        .AddItem "Kilobytes"
        .AddItem "Megabytes"
        .AddItem "Gigabytes"
        .AddItem "Terabytes"
    End With
    With cboRound
        .AddItem "0"
        .AddItem "1"
        .AddItem "2"
        .AddItem "3"
        .AddItem "4"
        .AddItem "5"
    End With
    
    
    Ex = Function_Exist("kernel32.dll", "GlobalMemoryStatusEx")
    If Ex = False Then
        lblAvailableExtendedVirtualMemory.Enabled = False
    End If
    
    
    cboOutput.ListIndex = GetRegSetting(HKEY_CURRENT_USER, "Software\Kira\MemoryStatus", "Output")
    cboRound.ListIndex = GetRegSetting(HKEY_CURRENT_USER, "Software\Kira\MemoryStatus", "Round")
    
    timerMemoryStatus_Timer
End Sub

Private Sub Form_Unload(Cancel As Integer)
    timerMemoryStatus.Enabled = False
    
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\MemoryStatus", "Output", cboOutput.ListIndex, REG_DWORD
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\MemoryStatus", "Round", cboRound.ListIndex, REG_DWORD
End Sub

Private Sub timerMemoryStatus_Timer()
    Dim TotalPhysical As Double
    Dim AvailablePhysical As Double
    Dim TotalPageFile As Double
    Dim AvailablePageFile As Double
    Dim TotalVirtual As Double
    Dim AvailableVirtual As Double
    Dim MemoryLoad As Long
    
    If Ex = True Then
        Dim MEMORYSTATUSEX As MEMORYSTATUSEX
        
        MEMORYSTATUSEX.dwLength = Len(MEMORYSTATUSEX)
        If GlobalMemoryStatusEx(MEMORYSTATUSEX) = False Then Failed "GlobalMemoryStatusEx"
        
        Dim AvailableExtendedVirtual As Double
        
        With MEMORYSTATUSEX
            MemoryLoad = .dwMemoryLoad
            TotalPhysical = CLargeInt(.ullTotalPhys.LowPart, .ullTotalPhys.HighPart)
            AvailablePhysical = CLargeInt(.ullAvailPhys.LowPart, .ullAvailPhys.HighPart)
            TotalPageFile = CLargeInt(.ullTotalPageFile.LowPart, .ullTotalPageFile.HighPart)
            AvailablePageFile = CLargeInt(.ullAvailPageFile.LowPart, .ullAvailPageFile.HighPart)
            TotalVirtual = CLargeInt(.ullTotalVirtual.LowPart, .ullTotalVirtual.HighPart)
            AvailableVirtual = CLargeInt(.ullAvailVirtual.LowPart, .ullAvailVirtual.HighPart)
            AvailableExtendedVirtual = CLargeInt(.ullAvailExtendedVirtual.LowPart, .ullAvailExtendedVirtual.HighPart)
        End With
        
        
        txtAvailableExtendedVirtualMemory.Text = CStr(AvailableExtendedVirtual / (1024 ^ cboOutput.ListIndex))
        txtAvailableExtendedVirtualMemoryPercentage.Text = CStr(Percentage(AvailableExtendedVirtual, TotalVirtual, 0)) & "%"
    Else
        Dim MEMORYSTATUS As MEMORYSTATUS
        GlobalMemoryStatus MEMORYSTATUS
        
        With MEMORYSTATUS
            MemoryLoad = .dwMemoryLoad
            TotalPhysical = .dwTotalPhys
            AvailablePhysical = .dwAvailPhys
            TotalPageFile = .dwTotalPageFile
            AvailablePageFile = .dwAvailPageFile
            TotalVirtual = .dwTotalVirtual
            AvailableVirtual = .dwAvailVirtual
        End With
    End If
    
    
    
    txtTotalPhysicalMemory.Text = CStr(Round(TotalPhysical / (1024 ^ cboOutput.ListIndex), cboRound.ListIndex))
    txtAvailablePhysicalMemory.Text = CStr(Round(AvailablePhysical / (1024 ^ cboOutput.ListIndex), cboRound.ListIndex))
    txtTotalPageFileMemory.Text = CStr(Round(TotalPageFile / (1024 ^ cboOutput.ListIndex), cboRound.ListIndex))
    txtAvailablePageFileMemory.Text = CStr(Round(AvailablePageFile / (1024 ^ cboOutput.ListIndex), cboRound.ListIndex))
    txtTotalVirtualMemory.Text = CStr(Round(TotalVirtual / (1024 ^ cboOutput.ListIndex), cboRound.ListIndex))
    txtAvailableVirtualMemory.Text = CStr(Round(AvailableVirtual / (1024 ^ cboOutput.ListIndex), cboRound.ListIndex))
    
    
    txtAvailablePhysicalMemoryPercentage.Text = CStr(Percentage(AvailablePhysical, TotalPhysical, 0)) & "%"
    txtAvailablePageFileMemoryPercentage.Text = CStr(Percentage(AvailablePageFile, TotalPageFile, 0)) & "%"
    txtAvailableVirtualMemoryPercentage.Text = CStr(Percentage(AvailableVirtual, TotalVirtual, 0)) & "%"
    
    
    txtMemoryLoad.Text = CStr(MemoryLoad) & "%"
End Sub
