VERSION 5.00
Begin VB.Form frmWindowMetrics 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Window Metrics"
   ClientHeight    =   3870
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6975
   Icon            =   "frmWindowMetrics.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   6975
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtFullScreenHeight 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   27
      Top             =   3240
      Width           =   1095
   End
   Begin VB.TextBox txtFullScreenWidth 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   25
      Top             =   3000
      Width           =   1095
   End
   Begin VB.TextBox txtFocusBorderHeight 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   2760
      Width           =   1095
   End
   Begin VB.TextBox txtFocusBorderWidth 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   2520
      Width           =   1095
   End
   Begin VB.TextBox txtDialogBorderHeight 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   2280
      Width           =   1095
   End
   Begin VB.TextBox txtKanjiHeight 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   29
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox txtSizingBorderHeight 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   54
      Top             =   3480
      Width           =   1095
   End
   Begin VB.TextBox txtSizingBorderWidth 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   52
      Top             =   3240
      Width           =   1095
   End
   Begin VB.TextBox txtNormalMinimizedHeight 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   50
      Top             =   3000
      Width           =   1095
   End
   Begin VB.TextBox txtNormalMinimizedWidth 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   48
      Top             =   2760
      Width           =   1095
   End
   Begin VB.TextBox txtMinimumTrackingHeight 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   46
      Top             =   2520
      Width           =   1095
   End
   Begin VB.TextBox txtMinimumTrackingWidth 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   44
      Top             =   2280
      Width           =   1095
   End
   Begin VB.TextBox txtMinimumHeight 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   42
      Top             =   2040
      Width           =   1095
   End
   Begin VB.TextBox txtMinimumWidth 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   40
      Top             =   1800
      Width           =   1095
   End
   Begin VB.TextBox txtMinimizedGridSpaceHeight 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   38
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox txtMinimizedGridSpaceWidth 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   36
      Top             =   1320
      Width           =   1095
   End
   Begin VB.TextBox txtDirection 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   34
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox txtStartingPosition 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   32
      Top             =   720
      Width           =   1095
   End
   Begin VB.TextBox txtDialogBorderWidth 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   2040
      Width           =   1095
   End
   Begin VB.TextBox txtDefaultMaximizedHeight 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   1800
      Width           =   1095
   End
   Begin VB.TextBox txtDefaultMaximizedWidth 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox txtDefaultMaximumHeight 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   1320
      Width           =   1095
   End
   Begin VB.TextBox txtDefaultMaximumWidth 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox txtBorderHeight 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox txtBorderWidth 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   600
      Width           =   1095
   End
   Begin VB.TextBox txt3DBorderHeight 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   360
      Width           =   1095
   End
   Begin VB.TextBox txt3DBorderWidth 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label lblFocusBorderWidth 
      Caption         =   "Focus Border Width"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Label lblFocusBorderHeight 
      Caption         =   "Focus Border Height"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   2760
      Width           =   2055
   End
   Begin VB.Label lblKanjiHeight 
      Caption         =   "Kanji Height"
      Height          =   255
      Left            =   3600
      TabIndex        =   28
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label lblSizingBorderHeight 
      Caption         =   "Sizing Border Height"
      Height          =   255
      Left            =   3600
      TabIndex        =   53
      Top             =   3480
      Width           =   2055
   End
   Begin VB.Label lblSizingBorderWidth 
      Caption         =   "Sizing Border Width"
      Height          =   255
      Left            =   3600
      TabIndex        =   51
      Top             =   3240
      Width           =   2055
   End
   Begin VB.Label lblNormalMinimizedHeight 
      Caption         =   "Normal Minimized Height"
      Height          =   255
      Left            =   3600
      TabIndex        =   49
      Top             =   3000
      Width           =   2055
   End
   Begin VB.Label lblNormalMinimizedWidth 
      Caption         =   "Normal Minimized Width"
      Height          =   255
      Left            =   3600
      TabIndex        =   47
      Top             =   2760
      Width           =   2055
   End
   Begin VB.Label lblMinimumTrackingHeight 
      Caption         =   "Minimum Tracking Height"
      Height          =   255
      Left            =   3600
      TabIndex        =   45
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Label lblMinimumTrackingWidth 
      Caption         =   "Minimum Tracking Width"
      Height          =   255
      Left            =   3600
      TabIndex        =   43
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Label lblMinimumHeight 
      Caption         =   "Minimum Height"
      Height          =   255
      Left            =   3600
      TabIndex        =   41
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Label lblMinimumWidth 
      Caption         =   "Minimum Width"
      Height          =   255
      Left            =   3600
      TabIndex        =   39
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Label lblMinimizedGridSpaceHeight 
      Caption         =   "Minimized GridSpace Height"
      Height          =   255
      Left            =   3600
      TabIndex        =   37
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label lblMinimizedGridSpaceWidth 
      Caption         =   "Minimized GridSpace Width"
      Height          =   255
      Left            =   3600
      TabIndex        =   35
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label lblDirection 
      Caption         =   "Direction"
      Height          =   255
      Left            =   3600
      TabIndex        =   33
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label lblStartingPosition 
      Caption         =   "Starting Position"
      Height          =   255
      Left            =   3600
      TabIndex        =   31
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label lblMinimizedArranging 
      Caption         =   "Minimized Arranging"
      Height          =   255
      Left            =   3600
      TabIndex        =   30
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label lblFullScreenHeight 
      Caption         =   "Full Screen Height"
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   3240
      Width           =   2055
   End
   Begin VB.Label lblFullScreenWidth 
      Caption         =   "Full Screen Width"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   3000
      Width           =   2055
   End
   Begin VB.Label lblDialogBorderHeight 
      Caption         =   "Dialog Border Height"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Label lblDialogBorderWidth 
      Caption         =   "Dialog Border Width"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Label lblDefaultMaximizedHeight 
      Caption         =   "Default Maximized Height"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Label lblDefaultMaximizedWidth 
      Caption         =   "Default Maximized Width"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label lblDefaultMaximumHeight 
      Caption         =   "Default Maximum Height"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label lblDefaultMaximumWidth 
      Caption         =   "Default Maximum Width"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label lblBorderHeight 
      Caption         =   "Border Height"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label lblBorderWidth 
      Caption         =   "Border Width"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label lbl3DBorderHeight 
      Caption         =   "3D Border Height"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label lbl3DBorderWidth 
      Caption         =   "3D Border Width"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "frmWindowMetrics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    txt3DBorderWidth.Text = CStr(GetSystemMetrics(SM_CXEDGE))
    txt3DBorderHeight.Text = CStr(GetSystemMetrics(SM_CYEDGE))
    txtBorderWidth.Text = CStr(GetSystemMetrics(SM_CXBORDER))
    txtBorderHeight.Text = CStr(GetSystemMetrics(SM_CYBORDER))
    txtDefaultMaximumWidth.Text = CStr(GetSystemMetrics(SM_CXMAXTRACK))
    txtDefaultMaximumHeight.Text = CStr(GetSystemMetrics(SM_CYMAXTRACK))
    txtDefaultMaximizedWidth.Text = CStr(GetSystemMetrics(SM_CXMAXIMIZED))
    txtDefaultMaximizedHeight.Text = CStr(GetSystemMetrics(SM_CYMAXIMIZED))
    txtDialogBorderWidth.Text = CStr(GetSystemMetrics(SM_CXFIXEDFRAME))
    txtDialogBorderHeight.Text = CStr(GetSystemMetrics(SM_CYFIXEDFRAME))
    
    
    If WinVersion(-1, 5010000, True) = True Then
        txtFocusBorderWidth.Text = CStr(GetSystemMetrics(SM_CXFOCUSBORDER))
        txtFocusBorderHeight.Text = CStr(GetSystemMetrics(SM_CYFOCUSBORDER))
    Else
        lblFocusBorderWidth.Enabled = False
        lblFocusBorderHeight.Enabled = False
    End If
    
    
    txtFullScreenWidth.Text = CStr(GetSystemMetrics(SM_CXFULLSCREEN))
    txtFullScreenHeight.Text = CStr(GetSystemMetrics(SM_CYFULLSCREEN))
    
    
    If GetSystemMetrics(SM_DBCSENABLED) = 0 Then
        lblKanjiHeight.Enabled = False
    Else
        txtKanjiHeight.Text = CStr(GetSystemMetrics(SM_CYKANJIWINDOW))
    End If
    
    
    Dim lngValue As Long
    lngValue = GetSystemMetrics(SM_ARRANGE)
    Select Case lngValue
        Case (lngValue And ARW_BOTTOMLEFT): txtStartingPosition.Text = "Bottom Left"
        Case (lngValue And ARW_BOTTOMRIGHT): txtStartingPosition.Text = "Bottom Right"
        Case (lngValue And ARW_HIDE): txtStartingPosition.Text = "Hide"
        Case (lngValue And ARW_TOPLEFT): txtStartingPosition.Text = "Top Left"
        Case (lngValue And ARW_TOPRIGHT): txtStartingPosition.Text = "Top Right"
    End Select
    
    
    lngValue = GetSystemMetrics(SM_ARRANGE)
    Select Case lngValue
        Case (lngValue And ARW_DOWN): txtDirection.Text = "Down"
        Case (lngValue And ARW_LEFT): txtDirection.Text = "Left"
        Case (lngValue And ARW_RIGHT): txtDirection.Text = "Right"
        Case (lngValue And ARW_UP): txtDirection.Text = "Up"
    End Select
    
    
    txtMinimizedGridSpaceWidth.Text = CStr(GetSystemMetrics(SM_CXMINSPACING))
    txtMinimizedGridSpaceHeight.Text = CStr(GetSystemMetrics(SM_CYMINSPACING))
    txtMinimumWidth.Text = CStr(GetSystemMetrics(SM_CXMIN))
    txtMinimumHeight.Text = CStr(GetSystemMetrics(SM_CXMIN))
    txtMinimumTrackingWidth.Text = CStr(GetSystemMetrics(SM_CXMINTRACK))
    txtMinimumTrackingHeight.Text = CStr(GetSystemMetrics(SM_CYMINTRACK))
    txtNormalMinimizedWidth.Text = CStr(GetSystemMetrics(SM_CXMINIMIZED))
    txtNormalMinimizedHeight.Text = CStr(GetSystemMetrics(SM_CYMINIMIZED))
    txtSizingBorderWidth.Text = CStr(GetSystemMetrics(SM_CXSIZEFRAME))
    txtSizingBorderHeight.Text = CStr(GetSystemMetrics(SM_CYSIZEFRAME))
End Sub
