VERSION 5.00
Begin VB.Form frmWindowPlacement 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Window Placement"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6735
   Icon            =   "frmWindowPlacement.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   6735
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtNormalPositionTop 
      Height          =   285
      Left            =   1800
      TabIndex        =   23
      Top             =   5160
      Width           =   1455
   End
   Begin VB.TextBox txtNormalPositionBottom 
      Height          =   285
      Left            =   1800
      TabIndex        =   25
      Top             =   5520
      Width           =   1455
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   350
      Left            =   5640
      TabIndex        =   36
      Top             =   5160
      Width           =   975
   End
   Begin VB.TextBox txtWindowText 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   2040
      Width           =   4815
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
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   350
      Left            =   5640
      TabIndex        =   5
      Top             =   1560
      Width           =   975
   End
   Begin VB.CheckBox chkSetMinimizedPosition 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   3000
      TabIndex        =   16
      Top             =   3600
      Width           =   255
   End
   Begin VB.CheckBox chkRestoreToMaximized 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   3000
      TabIndex        =   14
      Top             =   3360
      Width           =   255
   End
   Begin VB.CheckBox chkAsyncWindowPlacement 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   3000
      TabIndex        =   12
      Top             =   3120
      Width           =   255
   End
   Begin VB.ComboBox cboShowWindow 
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
      Left            =   1800
      TabIndex        =   9
      Top             =   2400
      Width           =   4815
   End
   Begin VB.TextBox txtMinimizedPositionX 
      Height          =   285
      Left            =   5160
      TabIndex        =   28
      Top             =   3240
      Width           =   1455
   End
   Begin VB.TextBox txtMinimizedPositionY 
      Height          =   285
      Left            =   5160
      TabIndex        =   30
      Top             =   3600
      Width           =   1455
   End
   Begin VB.TextBox txtMaximizedPositionX 
      Height          =   285
      Left            =   5160
      TabIndex        =   33
      Top             =   4320
      Width           =   1455
   End
   Begin VB.TextBox txtMaximizedPositionY 
      Height          =   285
      Left            =   5160
      TabIndex        =   35
      Top             =   4680
      Width           =   1455
   End
   Begin VB.TextBox txtNormalPositionLeft 
      Height          =   285
      Left            =   1800
      TabIndex        =   19
      Top             =   4440
      Width           =   1455
   End
   Begin VB.TextBox txtNormalPositionRight 
      Height          =   285
      Left            =   1800
      TabIndex        =   21
      Top             =   4800
      Width           =   1455
   End
   Begin VB.Label lblMaximizedPositionX 
      Caption         =   "X"
      Height          =   255
      Left            =   3480
      TabIndex        =   32
      Top             =   4320
      Width           =   1455
   End
   Begin VB.Label lblMaximizedPositionY 
      Caption         =   "Y"
      Height          =   255
      Left            =   3480
      TabIndex        =   34
      Top             =   4680
      Width           =   1455
   End
   Begin VB.Label lblNormalPositionBottom 
      Caption         =   "Bottom"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   5520
      Width           =   1455
   End
   Begin VB.Label lblNormalPositionTop 
      Caption         =   "Top"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   5160
      Width           =   1455
   End
   Begin VB.Label lblNormalPositionRight 
      Caption         =   "Right"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   4800
      Width           =   1455
   End
   Begin VB.Label lblNormalPositionLeft 
      Caption         =   "Left"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   4440
      Width           =   1455
   End
   Begin VB.Label lblWindow_Text 
      Caption         =   "Window Text"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label lblProcessID 
      Caption         =   "Process ID"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
   Begin VB.Label lblThreadID 
      Caption         =   "Thread ID"
      Height          =   255
      Left            =   1350
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lblHandle 
      Caption         =   "Handle"
      Height          =   255
      Left            =   2580
      TabIndex        =   2
      Top             =   120
      Width           =   555
   End
   Begin VB.Label lblWindowText 
      Caption         =   "Window Text"
      Height          =   255
      Left            =   3750
      TabIndex        =   3
      Top             =   120
      Width           =   975
   End
   Begin VB.Label lblMinimizedPositionY 
      Caption         =   "Y"
      Height          =   255
      Left            =   3480
      TabIndex        =   29
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Label lblMinimizedPositionX 
      Caption         =   "X"
      Height          =   255
      Left            =   3480
      TabIndex        =   27
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Label lblSetMinimizedPosition 
      Caption         =   "Set Minimized Position"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   3600
      Width           =   1935
   End
   Begin VB.Label lblRestoreToMaximized 
      Caption         =   "Restore To Maximized"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   3360
      Width           =   1935
   End
   Begin VB.Label lblAsyncWindowPlacement 
      Caption         =   "Async Window Placement"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   3120
      Width           =   1935
   End
   Begin VB.Label lblFlags 
      Caption         =   "Flags"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label lblShowWindow 
      Caption         =   "Show Window"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label lblMinimizedPosition 
      Caption         =   "Minimized Position"
      Height          =   255
      Left            =   3480
      TabIndex        =   26
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label lblMaximizedPosition 
      Caption         =   "Maximized Position"
      Height          =   255
      Left            =   3480
      TabIndex        =   31
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Label lblNormalPosition 
      Caption         =   "Normal Position"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   4080
      Width           =   1455
   End
End
Attribute VB_Name = "frmWindowPlacement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdApply_Click()
    If lstWindows.ListIndex > 0 Then
        Dim lngHandle As Long
        lngHandle = CLng(Trim$(Mid$(lstWindows.List(lstWindows.ListIndex), 21, 8)))
        
        Dim WINDOWPLACEMENT As WINDOWPLACEMENT
        With WINDOWPLACEMENT
            .Length = Len(WINDOWPLACEMENT)
            
            If WinVersion(-1, 5000000, True) = True Then
                If chkAsyncWindowPlacement.value = 1 Then .flags = .flags Or WPF_ASYNCWINDOWPLACEMENT
            End If
            If chkRestoreToMaximized.value = 1 Then .flags = .flags Or WPF_RESTORETOMAXIMIZED
            If chkSetMinimizedPosition.value = 1 Then .flags = .flags Or WPF_SETMINPOSITION
            
            Select Case cboShowWindow.ListIndex
                Case 0: .showCmd = SW_HIDE
                Case 1: .showCmd = SW_SHOWNORMAL
                Case 2: .showCmd = SW_SHOWMINIMIZED
                Case 3: .showCmd = SW_SHOWMAXIMIZED
                Case 4: .showCmd = SW_SHOWNOACTIVATE
                Case 5: .showCmd = SW_SHOW
                Case 6: .showCmd = SW_MINIMIZE
                Case 7: .showCmd = SW_SHOWMINNOACTIVE
                Case 8: .showCmd = SW_SHOWNA
                Case 9: .showCmd = SW_RESTORE
                Case 10: .showCmd = SW_SHOWDEFAULT
                Case 11: .showCmd = SW_FORCEMINIMIZE
            End Select
            
            .ptMinPosition.X = Val(txtMinimizedPositionX.Text)
            .ptMinPosition.Y = Val(txtMinimizedPositionY.Text)
            .ptMaxPosition.X = Val(txtMaximizedPositionX.Text)
            .ptMaxPosition.Y = Val(txtMaximizedPositionY.Text)
            .rcNormalPosition.Left = Val(txtNormalPositionLeft.Text)
            .rcNormalPosition.Right = Val(txtNormalPositionRight.Text)
            .rcNormalPosition.Top = Val(txtNormalPositionTop.Text)
            .rcNormalPosition.Bottom = Val(txtNormalPositionBottom.Text)
        End With
        
        If SetWindowPlacement(lngHandle, WINDOWPLACEMENT) = False Then Failed "SetWindowPlacement"
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
    
    
    txtWindowText.Text = ""
    cboShowWindow.ListIndex = -1
    chkAsyncWindowPlacement.value = 0
    chkRestoreToMaximized.value = 0
    chkSetMinimizedPosition.value = 0
    txtMinimizedPositionX.Text = ""
    txtMinimizedPositionY.Text = ""
    txtMaximizedPositionX.Text = ""
    txtMaximizedPositionY.Text = ""
    txtNormalPositionLeft.Text = ""
    txtNormalPositionRight.Text = ""
    txtNormalPositionTop.Text = ""
    txtNormalPositionBottom.Text = ""
    
    
    Erase WindowList()
    WindowListNum = 0
End Sub

Private Sub Form_Load()
    With cboShowWindow
        .AddItem "Hide"
        .AddItem "Show Normal"
        .AddItem "Show Minimized"
        .AddItem "Show Maximized"
        .AddItem "Show Not Active"
        .AddItem "Show"
        .AddItem "Minimize"
        .AddItem "Show Minimized Not Activated"
        .AddItem "Show NA"
        .AddItem "Restore"
        .AddItem "Show Default"
        .AddItem "Force Minimize"
    End With
    
    
    If WinVersion(-1, 5000000, True) = False Then
        lblAsyncWindowPlacement.Enabled = False
        chkAsyncWindowPlacement.Enabled = False
    End If
    
    cmdRefresh_Click
End Sub

Private Sub lstWindows_Click()
    Dim lngHandle As Long
    lngHandle = CLng(Trim$(Mid$(lstWindows.List(lstWindows.ListIndex), 21, 8)))
    
    txtWindowText.Text = Mid$(lstWindows.List(lstWindows.ListIndex), 31, Len(lstWindows.List(lstWindows.ListIndex)))
    
    
    Dim WINDOWPLACEMENT As WINDOWPLACEMENT
    WINDOWPLACEMENT.Length = Len(WINDOWPLACEMENT)
    If GetWindowPlacement(lngHandle, WINDOWPLACEMENT) = False Then Failed "GetWindowPlacement"
    
    With WINDOWPLACEMENT
        If WinVersion(-1, 5000000, True) = True Then
            If .flags And WPF_ASYNCWINDOWPLACEMENT Then chkAsyncWindowPlacement.value = 1 Else chkAsyncWindowPlacement.value = 0
        End If
        
        If .flags And WPF_RESTORETOMAXIMIZED Then chkRestoreToMaximized.value = 1 Else chkRestoreToMaximized.value = 0
        If .flags And WPF_SETMINPOSITION Then chkSetMinimizedPosition.value = 1 Else chkSetMinimizedPosition.value = 0
        
        Select Case .showCmd
            Case SW_HIDE: cboShowWindow.ListIndex = 0
            Case SW_SHOWNORMAL: cboShowWindow.ListIndex = 1
            Case SW_SHOWMINIMIZED: cboShowWindow.ListIndex = 2
            Case SW_SHOWMAXIMIZED: cboShowWindow.ListIndex = 3
            Case SW_SHOWNOACTIVATE: cboShowWindow.ListIndex = 4
            Case SW_SHOW: cboShowWindow.ListIndex = 5
            Case SW_MINIMIZE: cboShowWindow.ListIndex = 6
            Case SW_SHOWMINNOACTIVE: cboShowWindow.ListIndex = 7
            Case SW_SHOWNA: cboShowWindow.ListIndex = 8
            Case SW_RESTORE: cboShowWindow.ListIndex = 9
            Case SW_SHOWDEFAULT: cboShowWindow.ListIndex = 10
            Case SW_FORCEMINIMIZE: cboShowWindow.ListIndex = 11
            Case Else: cboShowWindow.ListIndex = -1
        End Select
        
        txtMinimizedPositionX.Text = CStr(.ptMinPosition.X)
        txtMinimizedPositionY.Text = CStr(.ptMinPosition.Y)
        txtMaximizedPositionX.Text = CStr(.ptMaxPosition.X)
        txtMaximizedPositionY.Text = CStr(.ptMaxPosition.Y)
        txtNormalPositionLeft.Text = CStr(.rcNormalPosition.Left)
        txtNormalPositionRight.Text = CStr(.rcNormalPosition.Right)
        txtNormalPositionTop.Text = CStr(.rcNormalPosition.Top)
        txtNormalPositionBottom.Text = CStr(.rcNormalPosition.Bottom)
    End With
End Sub

Private Sub txtMaximizedPositionX_Change()
    txtMaximizedPositionX.Text = CStr(Val(Rem_NonNumeric_Chr(txtMaximizedPositionX.Text)))
    If Val(txtMaximizedPositionX.Text) < -2147483648# Then txtMaximizedPositionX.Text = "-2147483648"
    If Val(txtMaximizedPositionX.Text) > 2147483647 Then txtMaximizedPositionX.Text = "2147483647"
End Sub

Private Sub txtMaximizedPositionY_Change()
    txtMaximizedPositionY.Text = CStr(Val(Rem_NonNumeric_Chr(txtMaximizedPositionY.Text)))
    If Val(txtMaximizedPositionY.Text) < -2147483648# Then txtMaximizedPositionY.Text = "-2147483648"
    If Val(txtMaximizedPositionY.Text) > 2147483647 Then txtMaximizedPositionY.Text = "2147483647"
End Sub

Private Sub txtMinimizedPositionX_Change()
    txtMinimizedPositionX.Text = CStr(Val(Rem_NonNumeric_Chr(txtMinimizedPositionX.Text)))
    If Val(txtMinimizedPositionX.Text) < -2147483648# Then txtMinimizedPositionX.Text = "-2147483648"
    If Val(txtMinimizedPositionX.Text) > 2147483647 Then txtMinimizedPositionX.Text = "2147483647"
End Sub

Private Sub txtMinimizedPositionY_Change()
    txtMinimizedPositionY.Text = CStr(Val(Rem_NonNumeric_Chr(txtMinimizedPositionY.Text)))
    If Val(txtMinimizedPositionY.Text) < -2147483648# Then txtMinimizedPositionY.Text = "-2147483648"
    If Val(txtMinimizedPositionY.Text) > 2147483647 Then txtMinimizedPositionY.Text = "2147483647"
End Sub

Private Sub txtNormalPositionBottom_Change()
    txtNormalPositionBottom.Text = CStr(Val(Rem_NonNumeric_Chr(txtNormalPositionBottom.Text)))
    If Val(txtNormalPositionBottom.Text) < -2147483648# Then txtNormalPositionBottom.Text = "-2147483648"
    If Val(txtNormalPositionBottom.Text) > 2147483647 Then txtNormalPositionBottom.Text = "2147483647"
End Sub

Private Sub txtNormalPositionLeft_Change()
    txtNormalPositionLeft.Text = CStr(Val(Rem_NonNumeric_Chr(txtNormalPositionLeft.Text)))
    If Val(txtNormalPositionLeft.Text) < -2147483648# Then txtNormalPositionLeft.Text = "-2147483648"
    If Val(txtNormalPositionLeft.Text) > 2147483647 Then txtNormalPositionLeft.Text = "2147483647"
End Sub

Private Sub txtNormalPositionRight_Change()
    txtNormalPositionRight.Text = CStr(Val(Rem_NonNumeric_Chr(txtNormalPositionRight.Text)))
    If Val(txtNormalPositionRight.Text) < -2147483648# Then txtNormalPositionRight.Text = "-2147483648"
    If Val(txtNormalPositionRight.Text) > 2147483647 Then txtNormalPositionRight.Text = "2147483647"
End Sub

Private Sub txtNormalPositionTop_Change()
    txtNormalPositionTop.Text = CStr(Val(Rem_NonNumeric_Chr(txtNormalPositionTop.Text)))
    If Val(txtNormalPositionTop.Text) < -2147483648# Then txtNormalPositionTop.Text = "-2147483648"
    If Val(txtNormalPositionTop.Text) > 2147483647 Then txtNormalPositionTop.Text = "2147483647"
End Sub
