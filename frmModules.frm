VERSION 5.00
Begin VB.Form frmModules 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Modules"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6255
   Icon            =   "frmModules.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   6255
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtUsageCount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox txtGlobalUsageCount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox txtExePath 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox txtBaseSize 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox txtBaseAddress 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   120
      Width           =   1215
   End
   Begin VB.ListBox lstModule 
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
      TabStop         =   0   'False
      Top             =   2880
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
   Begin VB.Label lblUsageCount 
      Caption         =   "Usage Count"
      Height          =   255
      Left            =   3240
      TabIndex        =   13
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label lblGlobalUsageCount 
      Caption         =   "Global Usage Count"
      Height          =   255
      Left            =   3240
      TabIndex        =   11
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label lblBaseAddress 
      Caption         =   "Base Address"
      Height          =   255
      Left            =   3240
      TabIndex        =   5
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label lblBaseSize 
      Caption         =   "Base Size"
      Height          =   255
      Left            =   3240
      TabIndex        =   7
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label lblExePath 
      Caption         =   "Exe Path"
      Height          =   255
      Left            =   3240
      TabIndex        =   9
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label lblModule 
      Caption         =   "Module Name"
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
Attribute VB_Name = "frmModules"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Process() As PROCESSENTRY32
Dim lngProcess As Long
Dim Module() As MODULEENTRY32
Dim lngModule As Long

Private Sub cmdRefresh_Click()
    lstProcess.Clear
    lstModule.Clear
    lngProcess = 0
    Erase Process()
    
    lngProcess = Process32_Enum(Process())
    
    Dim lngIncrement As Long
    For lngIncrement = 0 To lngProcess
        lstProcess.AddItem CStr(Process(lngIncrement).th32ProcessID)
    Next lngIncrement
    
    
    txtBaseAddress.Text = ""
    txtBaseSize.Text = ""
    txtExePath.Text = ""
    txtGlobalUsageCount.Text = ""
    txtUsageCount.Text = ""
End Sub

Private Sub Form_Load()
    cmdRefresh_Click
    
    
    If Function_Exist("kernel32.dll", "CreateToolhelp32Snapshot") = False Then
        lblProcess.Enabled = False
        lstProcess.Enabled = False
        cmdRefresh.Enabled = False
        lblModule.Enabled = False
        lstModule.Enabled = False
        lblBaseAddress.Enabled = False
        lblBaseSize.Enabled = False
        lblExePath.Enabled = False
        lblGlobalUsageCount.Enabled = False
        lblUsageCount.Enabled = False
    End If
End Sub

Private Sub lstModule_Click()
    With Module(lstModule.ListIndex)
        txtBaseAddress.Text = CStr(.modBaseAddr)
        txtBaseSize.Text = CStr(.modBaseSize)
        txtExePath.Text = .szExePath
        txtGlobalUsageCount.Text = CStr(.GlblcntUsage)
        txtUsageCount.Text = CStr(.ProccntUsage)
    End With
End Sub

Private Sub lstProcess_Click()
    lstModule.Clear
    lngModule = 0
    Erase Module()
    
    lngModule = Module32_Enum(Module(), Process(lstProcess.ListIndex).th32ProcessID)
    
    Dim lngIncrement As Long
    For lngIncrement = 0 To lngModule
        lstModule.AddItem CStr(Module(lngIncrement).szModule)
    Next lngIncrement
    
    
    txtBaseAddress.Text = ""
    txtBaseSize.Text = ""
    txtExePath.Text = ""
    txtGlobalUsageCount.Text = ""
    txtUsageCount.Text = ""
End Sub
