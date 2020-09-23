VERSION 5.00
Begin VB.Form frmWindowSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Window Settings"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6735
   Icon            =   "frmWindowSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   6735
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDestroy 
      Caption         =   "Destroy"
      Height          =   350
      Left            =   120
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1560
      Width           =   975
   End
   Begin VB.CheckBox chkShowWindow 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   6360
      TabIndex        =   38
      Top             =   3960
      Width           =   255
   End
   Begin VB.CheckBox chkNoZOrder 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   6360
      TabIndex        =   36
      Top             =   3720
      Width           =   255
   End
   Begin VB.CheckBox chkNoSize 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   6360
      TabIndex        =   34
      Top             =   3480
      Width           =   255
   End
   Begin VB.CheckBox chkNoSendChanging 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   6360
      TabIndex        =   32
      Top             =   3240
      Width           =   255
   End
   Begin VB.CheckBox chkNoRedraw 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   6360
      TabIndex        =   30
      Top             =   3000
      Width           =   255
   End
   Begin VB.CheckBox chkNoOwnerZOrder 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   3000
      TabIndex        =   28
      Top             =   4560
      Width           =   255
   End
   Begin VB.CheckBox chkNoMove 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   3000
      TabIndex        =   26
      Top             =   4320
      Width           =   255
   End
   Begin VB.CheckBox chkNoCopyBits 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   3000
      TabIndex        =   24
      Top             =   4080
      Width           =   255
   End
   Begin VB.CheckBox chkNoActivate 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   3000
      TabIndex        =   22
      Top             =   3840
      Width           =   255
   End
   Begin VB.CheckBox chkHideWindow 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   3000
      TabIndex        =   20
      Top             =   3600
      Width           =   255
   End
   Begin VB.CheckBox chkFrameChanged 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   3000
      TabIndex        =   18
      Top             =   3360
      Width           =   255
   End
   Begin VB.CheckBox chkDeferErase 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   3000
      TabIndex        =   16
      Top             =   3120
      Width           =   255
   End
   Begin VB.CheckBox chkAsyncWindowPos 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   3000
      TabIndex        =   14
      Top             =   2880
      Width           =   255
   End
   Begin VB.ComboBox cboZOrder 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5160
      TabIndex        =   12
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CheckBox chkInvertFlash 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   3000
      TabIndex        =   10
      Top             =   2520
      Width           =   255
   End
   Begin VB.TextBox txtWindowText 
      Height          =   285
      Left            =   1800
      TabIndex        =   8
      Top             =   2040
      Width           =   4815
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   350
      Left            =   5640
      TabIndex        =   39
      Top             =   4440
      Width           =   975
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   350
      Left            =   5640
      TabIndex        =   6
      Top             =   1560
      Width           =   975
   End
   Begin VB.ListBox lstWindows 
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1140
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   4
      Top             =   360
      Width           =   6495
   End
   Begin VB.Label lblShowWindow 
      Caption         =   "Show Window"
      Height          =   255
      Left            =   3480
      TabIndex        =   37
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Label lblNoZOrder 
      Caption         =   "No Z Order"
      Height          =   255
      Left            =   3480
      TabIndex        =   35
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Label lblNoSize 
      Caption         =   "No Size"
      Height          =   255
      Left            =   3480
      TabIndex        =   33
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label lblNoSendChanging 
      Caption         =   "No Send Changing"
      Height          =   255
      Left            =   3480
      TabIndex        =   31
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Label lblNoRedraw 
      Caption         =   "No Redraw"
      Height          =   255
      Left            =   3480
      TabIndex        =   29
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label lblNoOwnerZOrder 
      Caption         =   "No Owner Z Order"
      Height          =   255
      Left            =   120
      TabIndex        =   27
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Label lblNoMove 
      Caption         =   "No Move"
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   4320
      Width           =   1455
   End
   Begin VB.Label lblNoCopyBits 
      Caption         =   "No Copy Bits"
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Label lblNoActivate 
      Caption         =   "No Activate"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Label lblHideWindow 
      Caption         =   "Hide Window"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Label lblFrameChanged 
      Caption         =   "Frame Changed"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Label lblDeferErase 
      Caption         =   "Defer Erase"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label lblAsyncWindowPos 
      Caption         =   "Async Window Pos"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label lblZOrder 
      Caption         =   "Z Order"
      Height          =   255
      Left            =   3480
      TabIndex        =   11
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label lblInvertFlash 
      Caption         =   "Invert Flash"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label lblWindow_Text 
      Caption         =   "Window Text"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label lblWindowText 
      Caption         =   "Window Text"
      Height          =   255
      Left            =   3750
      TabIndex        =   3
      Top             =   120
      Width           =   975
   End
   Begin VB.Label lblHandle 
      Caption         =   "Handle"
      Height          =   255
      Left            =   2580
      TabIndex        =   2
      Top             =   120
      Width           =   555
   End
   Begin VB.Label lblThreadID 
      Caption         =   "Thread ID"
      Height          =   255
      Left            =   1350
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lblProcessID 
      Caption         =   "Process ID"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmWindowSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lngHandle As Long

Private Sub cmdApply_Click()
    If lngHandle > 0 Then
        If SetWindowText(lngHandle, txtWindowText.Text) = False Then Failed "SetWindowText"
        FlashWindow lngHandle, CBool(chkInvertFlash.value)
        
        If lstWindows.ListIndex > -1 Then
            Dim lngInsert As Long
            Dim lngFlags As Long
            
            Select Case cboZOrder.ListIndex
                Case 0: lngInsert = HWND_BOTTOM
                Case 1: lngInsert = HWND_NOTOPMOST
                Case 2: lngInsert = HWND_TOP
                Case 3: lngInsert = HWND_TOPMOST
            End Select
            
            If chkAsyncWindowPos.value = 1 Then lngFlags = lngFlags And SWP_ASYNCWINDOWPOS
            If chkDeferErase.value = 1 Then lngFlags = lngFlags And SWP_DEFERERASE
            If chkFrameChanged.value = 1 Then lngFlags = lngFlags And SWP_FRAMECHANGED
            If chkHideWindow.value = 1 Then lngFlags = lngFlags And SWP_HIDEWINDOW
            If chkNoActivate.value = 1 Then lngFlags = lngFlags And SWP_NOACTIVATE
            If chkNoCopyBits.value = 1 Then lngFlags = lngFlags And SWP_NOCOPYBITS
            If chkNoMove.value = 1 Then lngFlags = lngFlags And SWP_NOMOVE
            If chkNoOwnerZOrder.value = 1 Then lngFlags = lngFlags And SWP_NOOWNERZORDER
            If chkNoRedraw.value = 1 Then lngFlags = lngFlags And SWP_NOREDRAW
            If chkNoSendChanging.value = 1 Then lngFlags = lngFlags And SWP_NOSENDCHANGING
            If chkNoSize.value = 1 Then lngFlags = lngFlags And SWP_NOSIZE
            If chkNoZOrder.value = 1 Then lngFlags = lngFlags And SWP_NOZORDER
            If chkShowWindow.value = 1 Then lngFlags = lngFlags And SWP_SHOWWINDOW
            
            
            Dim WINDOWPLACEMENT As WINDOWPLACEMENT
            WINDOWPLACEMENT.Length = Len(WINDOWPLACEMENT)
            If GetWindowPlacement(lngHandle, WINDOWPLACEMENT) = False Then Failed "GetWindowPlacement"
            
            With WINDOWPLACEMENT.rcNormalPosition
                If SetWindowPos(lngHandle, lngInsert, .Left, .Top, .Right - .Left, .Bottom - .Top, lngFlags) = False Then Failed "SetWindowPos"
            End With
        End If
    End If
End Sub

Private Sub cmdDestroy_Click()
    If lngHandle > 0 Then
        If DestroyWindow(lngHandle) = False Then Failed "DestroyWindow"
    End If
End Sub

Private Sub cmdRefresh_Click()
    lstWindows.Clear
    Erase WindowList()
    WindowListNum = 0
    
    
    If EnumWindows(AddressOf EnumWindowsProc, &H0) = False Then Failed "EnumWindows"
    
    Dim lngIncrement As Long
    Dim strWindowTitle As String
    Dim lngProcessID As Long
    Dim lngThread As Long
    
    For lngIncrement = 1 To WindowListNum
        If WindowList(lngIncrement) <> 0 Then
            lngThread = GetWindowThreadProcessId(WindowList(lngIncrement), lngProcessID)
            
            lstWindows.AddItem Left$(CStr(lngProcessID) & Space$(10), 10) & _
                               Left$(CStr(lngThread) & Space$(10), 10) & _
                               Left$(CStr(WindowList(lngIncrement)) & Space$(10), 10) & _
                               Get_WindowText(WindowList(lngIncrement))
        End If
    Next lngIncrement
    
    
    lngHandle = 0
    txtWindowText.Text = ""
    
    
    Erase WindowList()
    WindowListNum = 0
End Sub

Private Sub Form_Load()
    With cboZOrder
        .AddItem "Bottom"
        .AddItem "Not Top Most"
        .AddItem "Top"
        .AddItem "Top Most"
        .ListIndex = 1
    End With
    
    cmdRefresh_Click
End Sub

Private Sub lstWindows_Click()
    lngHandle = CLng(Trim$(Mid$(lstWindows.List(lstWindows.ListIndex), 21, 8)))
    
    txtWindowText.Text = Mid$(lstWindows.List(lstWindows.ListIndex), 31, Len(lstWindows.List(lstWindows.ListIndex)))
End Sub
