VERSION 5.00
Begin VB.Form frmWindowInfo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Window Info"
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6735
   Icon            =   "frmWindowInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   6735
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtAssociatedModule 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   2280
      Width           =   4815
   End
   Begin VB.TextBox txtWindowCoordinatesBottom 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5160
      Locked          =   -1  'True
      TabIndex        =   51
      Top             =   4680
      Width           =   1455
   End
   Begin VB.TextBox txtWindowCoordinatesTop 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5160
      Locked          =   -1  'True
      TabIndex        =   49
      Top             =   4440
      Width           =   1455
   End
   Begin VB.TextBox txtWindowCoordinatesRight 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5160
      Locked          =   -1  'True
      TabIndex        =   47
      Top             =   4200
      Width           =   1455
   End
   Begin VB.TextBox txtWindowCoordinatesLeft 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5160
      Locked          =   -1  'True
      TabIndex        =   45
      Top             =   3960
      Width           =   1455
   End
   Begin VB.TextBox txtClientCoordinatesBottom 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5160
      Locked          =   -1  'True
      TabIndex        =   42
      Top             =   3360
      Width           =   1455
   End
   Begin VB.TextBox txtClientCoordinatesTop 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5160
      Locked          =   -1  'True
      TabIndex        =   40
      Top             =   3120
      Width           =   1455
   End
   Begin VB.TextBox txtClientCoordinatesRight 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5160
      Locked          =   -1  'True
      TabIndex        =   38
      Top             =   2880
      Width           =   1455
   End
   Begin VB.TextBox txtZOrderPrevious 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   31
      Top             =   5040
      Width           =   1455
   End
   Begin VB.TextBox txtZOrderNext 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   29
      Top             =   4800
      Width           =   1455
   End
   Begin VB.CheckBox chkUnicode 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   255
      Left            =   3000
      TabIndex        =   27
      Top             =   4560
      Width           =   255
   End
   Begin VB.TextBox txtRootOwner 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   25
      Top             =   4320
      Width           =   1455
   End
   Begin VB.TextBox txtRoot 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   4080
      Width           =   1455
   End
   Begin VB.TextBox txtParent 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   3840
      Width           =   1455
   End
   Begin VB.TextBox txtOwner 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   3600
      Width           =   1455
   End
   Begin VB.TextBox txtCreatorVersion 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   3360
      Width           =   1455
   End
   Begin VB.TextBox txtClassAtom 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   3120
      Width           =   1455
   End
   Begin VB.TextBox txtWindowBorderHeight 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   2880
      Width           =   1455
   End
   Begin VB.ListBox lstExtendedStyles 
      Height          =   645
      Left            =   3480
      Sorted          =   -1  'True
      TabIndex        =   53
      Top             =   5640
      Width           =   3135
   End
   Begin VB.ListBox lstStyles 
      Height          =   645
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   33
      Top             =   5640
      Width           =   3135
   End
   Begin VB.TextBox txtWindowBorderWidth 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   2640
      Width           =   1455
   End
   Begin VB.TextBox txtClientCoordinatesLeft 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5160
      Locked          =   -1  'True
      TabIndex        =   36
      Top             =   2640
      Width           =   1455
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
   Begin VB.Label lblAssociatedModule 
      Caption         =   "Associated Module"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label lblUnicode 
      Caption         =   "Unicode"
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Label lblZOrderPrevious 
      Caption         =   "Z Order Previous"
      Height          =   255
      Left            =   120
      TabIndex        =   30
      Top             =   5040
      Width           =   1455
   End
   Begin VB.Label lblZOrderNext 
      Caption         =   "Z Order Next"
      Height          =   255
      Left            =   120
      TabIndex        =   28
      Top             =   4800
      Width           =   1455
   End
   Begin VB.Label lblOwner 
      Caption         =   "Owner"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Label lblParent 
      Caption         =   "Parent"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Label lblRoot 
      Caption         =   "Root"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Label lblRootOwner 
      Caption         =   "Root Owner"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   4320
      Width           =   1455
   End
   Begin VB.Label lblExtendedStyles 
      Caption         =   "Extended Styles"
      Height          =   255
      Left            =   3480
      TabIndex        =   52
      Top             =   5400
      Width           =   1455
   End
   Begin VB.Label lblStyles 
      Caption         =   "Styles"
      Height          =   255
      Left            =   120
      TabIndex        =   32
      Top             =   5400
      Width           =   1455
   End
   Begin VB.Label lblCreatorVersion 
      Caption         =   "Creator Version"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Label lblClassAtom 
      Caption         =   "Class Atom"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label lblWindowBorderHeight 
      Caption         =   "Border Height"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label lblWindowBorderWidth 
      Caption         =   "Border Width"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label lblClientCoordinatesBottom 
      Caption         =   "Bottom"
      Height          =   255
      Left            =   3480
      TabIndex        =   41
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Label lblClientCoordinatesTop 
      Caption         =   "Top"
      Height          =   255
      Left            =   3480
      TabIndex        =   39
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Label lblClientCoordinatesRight 
      Caption         =   "Right"
      Height          =   255
      Left            =   3480
      TabIndex        =   37
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label lblClientCoordinates 
      Caption         =   "Client Coordinates"
      Height          =   255
      Left            =   3480
      TabIndex        =   34
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label lblClientCoordinatesLeft 
      Caption         =   "Left"
      Height          =   255
      Left            =   3480
      TabIndex        =   35
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label lblWindowCoordinatesBottom 
      Caption         =   "Bottom"
      Height          =   255
      Left            =   3480
      TabIndex        =   50
      Top             =   4920
      Width           =   1455
   End
   Begin VB.Label lblWindowCoordinatesTop 
      Caption         =   "Top"
      Height          =   255
      Left            =   3480
      TabIndex        =   48
      Top             =   4680
      Width           =   1455
   End
   Begin VB.Label lblWindowCoordinatesRight 
      Caption         =   "Right"
      Height          =   255
      Left            =   3480
      TabIndex        =   46
      Top             =   4440
      Width           =   1455
   End
   Begin VB.Label lblWindowCoordinates 
      Caption         =   "Window Coordinates"
      Height          =   255
      Left            =   3480
      TabIndex        =   43
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Label lblWindowCoordinatesLeft 
      Caption         =   "Left"
      Height          =   255
      Left            =   3480
      TabIndex        =   44
      Top             =   4200
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
End
Attribute VB_Name = "frmWindowInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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
    txtOwner.Text = ""
    txtZOrderNext.Text = ""
    txtZOrderPrevious.Text = ""
    chkUnicode.value = 0
    txtParent.Text = ""
    txtRoot.Text = ""
    txtRootOwner.Text = ""
    txtWindowCoordinatesLeft.Text = ""
    txtWindowCoordinatesRight.Text = ""
    txtWindowCoordinatesTop.Text = ""
    txtWindowCoordinatesBottom.Text = ""
    txtClientCoordinatesLeft.Text = ""
    txtClientCoordinatesRight.Text = ""
    txtClientCoordinatesTop.Text = ""
    txtClientCoordinatesBottom.Text = ""
    lstStyles.Clear
    lstExtendedStyles.Clear
    txtWindowBorderWidth.Text = ""
    txtWindowBorderHeight.Text = ""
    txtClassAtom.Text = ""
    txtCreatorVersion.Text = ""
    
    Erase WindowList()
    WindowListNum = 0
End Sub

Private Sub Form_Load()
    If Function_Exist("user32.dll", "GetAncestor") = False Then
        lblParent.Enabled = False
        lblRoot.Enabled = False
        lblRootOwner.Enabled = False
    End If
    If Function_Exist("user32.dll", "GetWindowInfo") = False Then
        lblClassAtom.Enabled = False
        lblClientCoordinates.Enabled = False
        lblClientCoordinatesLeft.Enabled = False
        lblClientCoordinatesRight.Enabled = False
        lblClientCoordinatesTop.Enabled = False
        lblClientCoordinatesBottom.Enabled = False
        lblCreatorVersion.Enabled = False
        lblExtendedStyles.Enabled = False
        lstExtendedStyles.Enabled = False
        lblStyles.Enabled = False
        lstStyles.Enabled = False
        lblWindowBorderWidth.Enabled = False
        lblWindowBorderHeight.Enabled = False
        lblWindowCoordinates.Enabled = False
        lblWindowCoordinatesLeft.Enabled = False
        lblWindowCoordinatesRight.Enabled = False
        lblWindowCoordinatesTop.Enabled = False
        lblWindowCoordinatesBottom.Enabled = False
    End If
    If Function_Exist("user32.dll", "GetWindowModuleFileNameA") = False Then
        lblAssociatedModule.Enabled = False
    End If
    
    cmdRefresh_Click
End Sub

Private Sub lstWindows_Click()
    Dim lngHandle As Long
    lngHandle = CLng(Trim$(Mid$(lstWindows.List(lstWindows.ListIndex), 21, 8)))
    
    txtWindowText.Text = Mid$(lstWindows.List(lstWindows.ListIndex), 31, Len(lstWindows.List(lstWindows.ListIndex)))
    
    
    txtOwner.Text = CStr(GetWindow(lngHandle, GW_OWNER))
    txtZOrderNext.Text = CStr(GetWindow(lngHandle, GW_HWNDNEXT))
    txtZOrderPrevious.Text = CStr(GetWindow(lngHandle, GW_HWNDPREV))
    chkUnicode.value = IsWindowUnicode(lngHandle)
    
    If Function_Exist("user32.dll", "GetAncestor") = True Then
        txtParent.Text = CStr(GetAncestor(lngHandle, GA_PARENT))
        txtRoot.Text = CStr(GetAncestor(lngHandle, GA_ROOT))
        txtRootOwner.Text = CStr(GetAncestor(lngHandle, GA_ROOTOWNER))
    End If
    If Function_Exist("user32.dll", "GetWindowInfo") = True Then
        Dim WINDOWINFO As WINDOWINFO
        WINDOWINFO.cbSize = Len(WINDOWINFO)
        If GetWindowInfo(lngHandle, WINDOWINFO) = False Then Failed "GetWindowInfo"
        
        With WINDOWINFO
            txtWindowCoordinatesLeft.Text = CStr(.rcWindow.Left)
            txtWindowCoordinatesRight.Text = CStr(.rcWindow.Right)
            txtWindowCoordinatesTop.Text = CStr(.rcWindow.Top)
            txtWindowCoordinatesBottom.Text = CStr(.rcWindow.Bottom)
            
            txtClientCoordinatesLeft.Text = CStr(.rcClient.Left)
            txtClientCoordinatesRight.Text = CStr(.rcClient.Right)
            txtClientCoordinatesTop.Text = CStr(.rcClient.Top)
            txtClientCoordinatesBottom.Text = CStr(.rcClient.Bottom)
            
            lstStyles.Clear
            If .dwStyle And WS_OVERLAPPED Then lstStyles.AddItem "Overlapped"
            If .dwStyle And WS_POPUP Then lstStyles.AddItem "Popup"
            If .dwStyle And WS_CHILD Then lstStyles.AddItem "Child"
            If .dwStyle And WS_MINIMIZE Then lstStyles.AddItem "Minimize"
            If .dwStyle And WS_VISIBLE Then lstStyles.AddItem "Visible"
            If .dwStyle And WS_DISABLED Then lstStyles.AddItem "Disabled"
            If .dwStyle And WS_CLIPSIBLINGS Then lstStyles.AddItem "Clip Siblings"
            If .dwStyle And WS_CLIPCHILDREN Then lstStyles.AddItem "Clib Children"
            If .dwStyle And WS_MAXIMIZE Then lstStyles.AddItem "Maximize"
            If .dwStyle And WS_CAPTION Then lstStyles.AddItem "Caption"
            If .dwStyle And WS_BORDER Then lstStyles.AddItem "Border"
            If .dwStyle And WS_DLGFRAME Then lstStyles.AddItem "Dialog Frame"
            If .dwStyle And WS_VSCROLL Then lstStyles.AddItem "Vertical Scroll"
            If .dwStyle And WS_HSCROLL Then lstStyles.AddItem "Horizontal Scroll"
            If .dwStyle And WS_SYSMENU Then lstStyles.AddItem "System Menu"
            If .dwStyle And WS_THICKFRAME Then lstStyles.AddItem "Thick Frame"
            If .dwStyle And WS_GROUP Then lstStyles.AddItem "Group"
            If .dwStyle And WS_TABSTOP Then lstStyles.AddItem "Tab Stop"
            If .dwStyle And WS_MINIMIZEBOX Then lstStyles.AddItem "Minimize Box"
            If .dwStyle And WS_MAXIMIZEBOX Then lstStyles.AddItem "Maximize Box"
            
            lstExtendedStyles.Clear
            If .dwExStyle And WS_EX_DLGMODALFRAME Then lstExtendedStyles.AddItem "Dialog Modal Frame"
            If .dwExStyle And WS_EX_NOPARENTNOTIFY Then lstExtendedStyles.AddItem "No Parent Notify"
            If .dwExStyle And WS_EX_TOPMOST Then lstExtendedStyles.AddItem "Top Most"
            If .dwExStyle And WS_EX_ACCEPTFILES Then lstExtendedStyles.AddItem "Accept Files"
            If .dwExStyle And WS_EX_TRANSPARENT Then lstExtendedStyles.AddItem "Transparent"
            If .dwExStyle And WS_EX_MDICHILD Then lstExtendedStyles.AddItem "MDI Child"
            If .dwExStyle And WS_EX_TOOLWINDOW Then lstExtendedStyles.AddItem "Tool Window"
            If .dwExStyle And WS_EX_WINDOWEDGE Then lstExtendedStyles.AddItem "Window Edge"
            If .dwExStyle And WS_EX_CLIENTEDGE Then lstExtendedStyles.AddItem "Client Edge"
            If .dwExStyle And WS_EX_CONTEXTHELP Then lstExtendedStyles.AddItem "Context Help"
            If .dwExStyle And WS_EX_RIGHT Then lstExtendedStyles.AddItem "Right"
            If .dwExStyle And WS_EX_LEFT Then lstExtendedStyles.AddItem "Left"
            If .dwExStyle And WS_EX_RTLREADING Then lstExtendedStyles.AddItem "Right Reading"
            If .dwExStyle And WS_EX_LEFTSCROLLBAR Then lstExtendedStyles.AddItem "Left Scroll Bar"
            If .dwExStyle And WS_EX_CONTROLPARENT Then lstExtendedStyles.AddItem "Control Parent"
            If .dwExStyle And WS_EX_STATICEDGE Then lstExtendedStyles.AddItem "Static Edge"
            If .dwExStyle And WS_EX_APPWINDOW Then lstExtendedStyles.AddItem "App Window"
            If .dwExStyle And WS_EX_LAYERED Then lstExtendedStyles.AddItem "Layered"
            If .dwExStyle And WS_EX_NOINHERITLAYOUT Then lstExtendedStyles.AddItem "No Inherit Layout"
            If .dwExStyle And WS_EX_LAYOUTRTL Then lstExtendedStyles.AddItem "Layout Right"
            If .dwExStyle And WS_EX_COMPOSITED Then lstExtendedStyles.AddItem "Composited"
            If .dwExStyle And WS_EX_NOACTIVATE Then lstExtendedStyles.AddItem "No Active"
            
            txtWindowBorderWidth.Text = CStr(.cxWindowBorders)
            txtWindowBorderHeight.Text = CStr(.cyWindowBorders)
            txtClassAtom.Text = CStr(.atomWindowType)
            txtCreatorVersion.Text = CStr(.wCreatorVersion)
        End With
    End If
    If Function_Exist("user32.dll", "GetWindowModuleFileNameA") = True Then
        Dim strFileName As String
        strFileName = String$(MAX_PATH, Chr$(0))
        txtAssociatedModule.Text = Left$(strFileName, GetWindowModuleFileName(lngHandle, strFileName, Len(strFileName)))
    End If
End Sub
