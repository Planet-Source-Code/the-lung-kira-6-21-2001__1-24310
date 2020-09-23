VERSION 5.00
Begin VB.Form frmHeaps 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Heaps"
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6135
   Icon            =   "frmHeaps.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   6135
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtLockCount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   5640
      Width           =   1215
   End
   Begin VB.TextBox txtHeapHandle 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   5400
      Width           =   1215
   End
   Begin VB.TextBox txtFlags 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   5160
      Width           =   1215
   End
   Begin VB.TextBox txtBlockSize 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CheckBox chkDefaultHeap 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   255
      Left            =   5760
      TabIndex        =   8
      Top             =   2880
      Width           =   255
   End
   Begin VB.ListBox lstHeapList 
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
   Begin VB.ListBox lstHeap 
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
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   4920
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
   Begin VB.Label lblHeapHandle 
      Caption         =   "Heap Handle"
      Height          =   255
      Left            =   3240
      TabIndex        =   13
      Top             =   5400
      Width           =   1335
   End
   Begin VB.Label lblBlockSize 
      Caption         =   "Block Size"
      Height          =   255
      Left            =   3240
      TabIndex        =   9
      Top             =   4920
      Width           =   1335
   End
   Begin VB.Label lblLockCount 
      Caption         =   "Lock Count"
      Height          =   255
      Left            =   3240
      TabIndex        =   15
      Top             =   5640
      Width           =   1335
   End
   Begin VB.Label lblFlags 
      Caption         =   "Flags"
      Height          =   255
      Left            =   3240
      TabIndex        =   11
      Top             =   5160
      Width           =   1335
   End
   Begin VB.Label lblDefaultHeap 
      Caption         =   "Default Heap"
      Height          =   255
      Left            =   3240
      TabIndex        =   7
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Label lblHeapList 
      Caption         =   "Heap List ID"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label lblHeap 
      Caption         =   "Heap Address"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   4680
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
Attribute VB_Name = "frmHeaps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Process() As PROCESSENTRY32
Dim lngProcess As Long
Dim HeapList() As HEAPLIST32
Dim lngHeapList As Long
Dim Heap() As HEAPENTRY32
Dim lngHeap As Long

Private Sub cmdRefresh_Click()
    lstProcess.Clear
    lstHeapList.Clear
    lstHeap.Clear
    lngProcess = 0
    Erase Process()
    
    lngProcess = Process32_Enum(Process())
    
    Dim lngIncrement As Long
    For lngIncrement = 0 To lngProcess
        lstProcess.AddItem CStr(Process(lngIncrement).th32ProcessID)
    Next lngIncrement
    
    chkDefaultHeap.value = 0
    txtBlockSize.Text = ""
    txtFlags.Text = ""
    txtHeapHandle.Text = ""
    txtLockCount.Text = ""
End Sub

Private Sub Form_Load()
    cmdRefresh_Click
    
    
    If Function_Exist("kernel32.dll", "CreateToolhelp32Snapshot") = False Then
        lblProcess.Enabled = False
        lstProcess.Enabled = False
        lblHeapList.Enabled = False
        lstHeapList.Enabled = False
        lblHeap.Enabled = False
        lstHeap.Enabled = False
        lblDefaultHeap.Enabled = False
        lblBlockSize.Enabled = False
        lblFlags.Enabled = False
        lblHeapHandle.Enabled = False
        lblLockCount.Enabled = False
        cmdRefresh.Enabled = False
    End If
End Sub

Private Sub lstHeap_Click()
    With Heap(lstHeap.ListIndex)
        txtBlockSize.Text = CStr(.dwBlockSize)
        
        Select Case .dwFlags
            Case LF32_FIXED: txtFlags.Text = "Fixed"
            Case LF32_FREE: txtFlags.Text = "Free"
            Case LF32_MOVEABLE: txtFlags.Text = "Moveable"
            Case Else: txtFlags.Text = CStr(.dwFlags)
        End Select
        
        txtHeapHandle.Text = CStr(.hHandle)
        txtLockCount.Text = CStr(.dwLockCount)
    End With
End Sub

Private Sub lstHeapList_Click()
    lstHeap.Clear
    lngHeap = 0
    Erase Heap()
    
    lngHeap = Heap32_Enum(Heap(), Process(lstProcess.ListIndex).th32ProcessID, lstHeapList.List(lstHeapList.ListIndex))
    
    Dim lngIncrement As Long
    For lngIncrement = 0 To lngHeap
        lstHeap.AddItem CStr(Heap(lngIncrement).dwAddress)
    Next lngIncrement
    
    If HeapList(lstHeapList.ListIndex).dwFlags And HF32_DEFAULT Then
        chkDefaultHeap.value = 1
    Else
        chkDefaultHeap.value = 0
    End If
    
    txtBlockSize.Text = ""
    txtFlags.Text = ""
    txtHeapHandle.Text = ""
    txtLockCount.Text = ""
End Sub

Private Sub lstProcess_Click()
    lstHeapList.Clear
    lstHeap.Clear
    lngHeapList = 0
    Erase HeapList()
    
    lngHeapList = Heap32List_Enum(HeapList(), Process(lstProcess.ListIndex).th32ProcessID)
    
    Dim lngIncrement As Long
    For lngIncrement = 0 To lngHeapList
        lstHeapList.AddItem CStr(HeapList(lngIncrement).th32HeapID)
    Next lngIncrement
    
    chkDefaultHeap.value = 0
    txtBlockSize.Text = ""
    txtFlags.Text = ""
    txtHeapHandle.Text = ""
    txtLockCount.Text = ""
End Sub
