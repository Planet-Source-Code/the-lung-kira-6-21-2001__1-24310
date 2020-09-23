VERSION 5.00
Begin VB.Form frmCPUInfo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CPU Info"
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4455
   Icon            =   "frmCPUInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   4455
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkLowEndCPU 
      Enabled         =   0   'False
      Height          =   255
      Left            =   4080
      TabIndex        =   9
      Top             =   1080
      Width           =   255
   End
   Begin VB.Timer timerCyclesElapsed 
      Interval        =   1000
      Left            =   1680
      Top             =   960
   End
   Begin VB.TextBox txtSpeed 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   840
      Width           =   2175
   End
   Begin VB.TextBox txtArchitecture 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   600
      Width           =   2175
   End
   Begin VB.TextBox txtActiveProcessorMask 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   360
      Width           =   2175
   End
   Begin VB.TextBox txtProcessors 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   2175
   End
   Begin VB.TextBox txtCyclesElapsed 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   1440
      Width           =   2175
   End
   Begin VB.Label lblActiveProcessorMask 
      Caption         =   "Active Processor Mask"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label lblLowEndCPU 
      Caption         =   "Low End CPU"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label lblProcessors 
      Caption         =   "Number of Processors"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label lblSpeed 
      Caption         =   "Approx Speed"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label lblCyclesElapsed 
      Caption         =   "Cycles Elapsed"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label lblArchitecture 
      Caption         =   "Architecture"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   1815
   End
End
Attribute VB_Name = "frmCPUInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
    DoEvents
    
    Dim Cycles As Double
    Cycles = rdtsc_
    
    Sleep 1000
    
    Cycles = rdtsc_ - Cycles
    txtSpeed.Text = CStr(Cycles \ 1000000) & " Mhz"
End Sub

Private Sub Form_Load()
    Dim SYSTEM_INFO As SYSTEM_INFO
    GetSystemInfo SYSTEM_INFO
    
    txtActiveProcessorMask.Text = StrReverse(Right$(String$(32, "0") & ltoa_(SYSTEM_INFO.dwActiveProcessorMask, 2), 32))
    txtCyclesElapsed.Text = Format$(rdtsc_, "###,###")
    
    Select Case LO_WORD(SYSTEM_INFO.dwOemID)
        Case PROCESSOR_ARCHITECTURE_INTEL: txtArchitecture.Text = "Intel"
        Case PROCESSOR_ARCHITECTURE_MIPS: txtArchitecture.Text = "MIPS"
        Case PROCESSOR_ARCHITECTURE_ALPHA: txtArchitecture.Text = "Alpha"
        Case PROCESSOR_ARCHITECTURE_PPC: txtArchitecture.Text = "PPC"
        Case PROCESSOR_ARCHITECTURE_SHX: txtArchitecture.Text = "SHX"
        Case PROCESSOR_ARCHITECTURE_ARM: txtArchitecture.Text = "ARM"
        Case PROCESSOR_ARCHITECTURE_IA64: txtArchitecture.Text = "IA-64"
        Case PROCESSOR_ARCHITECTURE_ALPHA64: txtArchitecture.Text = "Alpha 64"
        Case PROCESSOR_ARCHITECTURE_MSIL: txtArchitecture.Text = "MSIL"
        Case PROCESSOR_ARCHITECTURE_AMD64: txtArchitecture.Text = "AMD 64"
        Case PROCESSOR_ARCHITECTURE_UNKNOWN: txtArchitecture.Text = "Unknown"
        Case Else: txtArchitecture.Text = "Unknown"
    End Select
    
    txtProcessors.Text = CStr(SYSTEM_INFO.dwNumberOrfProcessors)
    chkLowEndCPU.value = CInt(GetSystemMetrics(SM_SLOWMACHINE))
    
    
    timerCyclesElapsed_Timer
End Sub

Private Sub Form_Unload(Cancel As Integer)
    timerCyclesElapsed.Enabled = False
End Sub

Private Sub timerCyclesElapsed_Timer()
    txtCyclesElapsed.Text = Format$(rdtsc_, "###,###")
End Sub
