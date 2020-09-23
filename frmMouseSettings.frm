VERSION 5.00
Begin VB.Form frmMouseSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mouse Settings"
   ClientHeight    =   5235
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3015
   Icon            =   "frmMouseSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   3015
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkHotTracking 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   2640
      TabIndex        =   11
      Top             =   1920
      Width           =   255
   End
   Begin VB.TextBox txtCursorTrails 
      Height          =   285
      Left            =   1920
      TabIndex        =   3
      Top             =   480
      Width           =   975
   End
   Begin VB.CheckBox chkCursorShadow 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   2640
      TabIndex        =   1
      Top             =   120
      Width           =   255
   End
   Begin VB.TextBox txtSpeed 
      Height          =   285
      Left            =   1920
      TabIndex        =   21
      Top             =   3720
      Width           =   975
   End
   Begin VB.TextBox txtWheelScrollLines 
      Height          =   285
      Left            =   1920
      TabIndex        =   25
      Top             =   4440
      Width           =   975
   End
   Begin VB.CheckBox chkSnapToDefault 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   2640
      TabIndex        =   19
      Top             =   3360
      Width           =   255
   End
   Begin VB.CheckBox chkSwapButton 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   2640
      TabIndex        =   23
      Top             =   4080
      Width           =   255
   End
   Begin VB.TextBox txtHoverTimeHeight 
      Height          =   285
      Left            =   1920
      TabIndex        =   15
      Top             =   2640
      Width           =   975
   End
   Begin VB.TextBox txtHoverTimeWidth 
      Height          =   285
      Left            =   1920
      TabIndex        =   17
      Top             =   3000
      Width           =   975
   End
   Begin VB.TextBox txtDoubleClickWidth 
      Height          =   285
      Left            =   1920
      TabIndex        =   9
      Top             =   1560
      Width           =   975
   End
   Begin VB.TextBox txtDoubleClickHeight 
      Height          =   285
      Left            =   1920
      TabIndex        =   7
      Top             =   1200
      Width           =   975
   End
   Begin VB.TextBox txtHoverTime 
      Height          =   285
      Left            =   1920
      TabIndex        =   13
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   350
      Left            =   1920
      TabIndex        =   26
      Top             =   4800
      Width           =   975
   End
   Begin VB.TextBox txtDoubleClickTime 
      Height          =   285
      Left            =   1920
      TabIndex        =   5
      Top             =   840
      Width           =   975
   End
   Begin VB.Label lblHotTracking 
      Caption         =   "Hot Tracking"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label lblCursorTrails 
      Caption         =   "Cursor Trails"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label lblCursorShadow 
      Caption         =   "Cursor Shadow"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label lblSpeed 
      Caption         =   "Speed"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   3720
      Width           =   1575
   End
   Begin VB.Label lblWheelScrollLines 
      Caption         =   "Wheel Scroll Lines"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Label lblSnapToDefault 
      Caption         =   "Snap To Default"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Label lblSwapButton 
      Caption         =   "Swap Button"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   4080
      Width           =   1575
   End
   Begin VB.Label lblHoverTimeHeight 
      Caption         =   "Hover Time Height"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Label lblHoverTimeWidth 
      Caption         =   "Hover Time Width"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label lblDoubleClickWidth 
      Caption         =   "Double Click Width"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label lblDoubleClickHeight 
      Caption         =   "Double Click Height"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label lblHoverTime 
      Caption         =   "Hover Time"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label lblDoubleClickTime 
      Caption         =   "Double Click Time"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1575
   End
End
Attribute VB_Name = "frmMouseSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdApply_Click()
    If SetDoubleClickTime(CLng(txtDoubleClickTime.Text)) = False Then Failed "SetDoubleClickTime"
    
    
    If WinVersion(-1, 5000000, True) = True Then
        If SystemParametersInfo(SPI_SETCURSORSHADOW, 0, ByVal CBool(chkCursorShadow.value), SPIF_UPDATEINIFILE) = False Then Failed "SystemParametersInfo"
    End If
    If WinVersion(0, 5010000, True) = True Then
        If SystemParametersInfo(SPI_SETMOUSETRAILS, CLng(txtCursorTrails.Text), 0, SPIF_UPDATEINIFILE) = False Then Failed "SystemParametersInfo"
    End If
    
    If SystemParametersInfo(SPI_SETDOUBLECLKHEIGHT, CLng(txtDoubleClickHeight.Text), 0, SPIF_UPDATEINIFILE) = False Then Failed "SystemParametersInfo"
    If SystemParametersInfo(SPI_SETDOUBLECLKWIDTH, CLng(txtDoubleClickWidth.Text), 0, SPIF_UPDATEINIFILE) = False Then Failed "SystemParametersInfo"
    
    If WinVersion(4010000, 0, True) = True Then
        If SystemParametersInfo(SPI_SETMOUSEHOVERTIME, CLng(txtHoverTime.Text), 0, SPIF_UPDATEINIFILE) = False Then Failed "SystemParametersInfo"
        If SystemParametersInfo(SPI_SETMOUSEHOVERHEIGHT, CLng(txtHoverTimeHeight.Text), 0, SPIF_UPDATEINIFILE) = False Then Failed "SystemParametersInfo"
        If SystemParametersInfo(SPI_SETMOUSEHOVERWIDTH, CLng(txtHoverTimeWidth.Text), 0, SPIF_UPDATEINIFILE) = False Then Failed "SystemParametersInfo"
        
        If SystemParametersInfo(SPI_SETSNAPTODEFBUTTON, chkSnapToDefault.value, 0, SPIF_UPDATEINIFILE) = False Then Failed "SystemParametersInfo"
        
        If SystemParametersInfo(SPI_SETWHEELSCROLLLINES, CLng(txtWheelScrollLines.Text), 0, SPIF_UPDATEINIFILE) = False Then Failed "SystemParametersInfo"
    End If
    If WinVersion(4010000, 5000000, True) = True Then
        If SystemParametersInfo(SPI_SETHOTTRACKING, 0, ByVal CBool(chkHotTracking.value), SPIF_UPDATEINIFILE) = False Then Failed "SystemParametersInfo"
        If SystemParametersInfo(SPI_SETMOUSESPEED, 0, ByVal CLng(txtSpeed.Text), SPIF_UPDATEINIFILE) = False Then Failed "SystemParametersInfo"
    End If
    
    If SystemParametersInfo(SPI_SETMOUSEBUTTONSWAP, chkSwapButton.value, 0, SPIF_UPDATEINIFILE) = False Then Failed "SystemParametersInfo"
End Sub

Private Sub Form_Load()
    Dim lngValue As Long
    
    If WinVersion(-1, 5000000, True) = True Then
        If SystemParametersInfo(SPI_GETCURSORSHADOW, 0, lngValue, 0) = False Then Failed "SystemParametersInfo"
        chkCursorShadow.value = lngValue
    Else
        lblCursorShadow.Enabled = False
        chkCursorShadow.Enabled = False
    End If
    If WinVersion(0, 5010000, True) = True Then
        If SystemParametersInfo(SPI_GETMOUSETRAILS, 0, lngValue, 0) = False Then Failed "SystemParametersInfo"
        txtCursorTrails.Text = CStr(lngValue)
    Else
        lblCursorTrails.Enabled = False
        txtCursorTrails.Enabled = False
    End If
    
    txtDoubleClickTime.Text = CStr(GetDoubleClickTime)
    txtDoubleClickHeight.Text = CStr(GetSystemMetrics(SM_CYDOUBLECLK))
    txtDoubleClickWidth.Text = CStr(GetSystemMetrics(SM_CXDOUBLECLK))
    
    If WinVersion(4010000, 0, True) = True Then
        If SystemParametersInfo(SPI_GETMOUSEHOVERTIME, 0, lngValue, 0) = False Then Failed "SystemParametersInfo"
        txtHoverTime.Text = CStr(lngValue)
        If SystemParametersInfo(SPI_GETMOUSEHOVERHEIGHT, 0, lngValue, 0) = False Then Failed "SystemParametersInfo"
        txtHoverTimeHeight.Text = CStr(lngValue)
        If SystemParametersInfo(SPI_GETMOUSEHOVERWIDTH, 0, lngValue, 0) = False Then Failed "SystemParametersInfo"
        txtHoverTimeWidth.Text = CStr(lngValue)
        
        If SystemParametersInfo(SPI_GETSNAPTODEFBUTTON, 0, lngValue, 0) = False Then Failed "SystemParametersInfo"
        chkSnapToDefault.value = lngValue
        
        If SystemParametersInfo(SPI_GETWHEELSCROLLLINES, 0, lngValue, 0) = False Then Failed "SystemParametersInfo"
        txtWheelScrollLines.Text = CStr(lngValue)
    Else
        lblHoverTime.Enabled = False
        txtHoverTime.Enabled = False
        lblHoverTimeHeight.Enabled = False
        txtHoverTimeHeight.Enabled = False
        lblHoverTimeWidth.Enabled = False
        txtHoverTimeWidth.Enabled = False
        lblSnapToDefault.Enabled = False
        chkSnapToDefault.Enabled = False
        lblWheelScrollLines.Enabled = False
        txtWheelScrollLines.Enabled = False
    End If
    If WinVersion(4010000, 5000000, True) = True Then
        Dim boolHotTrack As Boolean
        If SystemParametersInfo(SPI_GETHOTTRACKING, 0, boolHotTrack, 0) = False Then Failed "SystemParametersInfo"
        chkHotTracking.value = boolHotTrack
        
        If SystemParametersInfo(SPI_GETMOUSESPEED, 0, lngValue, 0) = False Then Failed "SystemParametersInfo"
        txtSpeed.Text = CStr(lngValue)
    Else
        lblSpeed.Enabled = False
        txtSpeed.Enabled = False
    End If
    
    chkSwapButton.value = GetSystemMetrics(SM_SWAPBUTTON)
End Sub

Private Sub txtCursorTrails_Change()
    txtCursorTrails.Text = CStr(Val(Rem_NonNumeric_Chr(txtCursorTrails.Text)))
    If Val(txtCursorTrails.Text) < 0 Then txtCursorTrails.Text = "0"
    If Val(txtCursorTrails.Text) > 16 Then txtCursorTrails.Text = "16"
End Sub

Private Sub txtDoubleClickHeight_Change()
    txtDoubleClickHeight.Text = CStr(Val(Rem_NonNumeric_Chr(txtDoubleClickHeight.Text)))
    If Val(txtDoubleClickHeight.Text) < 0 Then txtDoubleClickHeight.Text = "0"
    If Val(txtDoubleClickHeight.Text) > 65535 Then txtDoubleClickHeight.Text = "65535"
End Sub

Private Sub txtDoubleClickTime_Change()
    txtDoubleClickTime.Text = CStr(Val(Rem_NonNumeric_Chr(txtDoubleClickTime.Text)))
    If Val(txtDoubleClickTime.Text) < 0 Then txtDoubleClickTime.Text = "0"
    If Val(txtDoubleClickTime.Text) > 5000 Then txtDoubleClickTime.Text = "5000"
End Sub

Private Sub txtDoubleClickWidth_Change()
    txtDoubleClickWidth.Text = CStr(Val(Rem_NonNumeric_Chr(txtDoubleClickWidth.Text)))
    If Val(txtDoubleClickWidth.Text) < 0 Then txtDoubleClickWidth.Text = "0"
    If Val(txtDoubleClickWidth.Text) > 65535 Then txtDoubleClickWidth.Text = "65535"
End Sub

Private Sub txtHoverTime_Change()
    txtHoverTime.Text = CStr(Val(Rem_NonNumeric_Chr(txtHoverTime.Text)))
    If Val(txtHoverTime.Text) < 0 Then txtHoverTime.Text = "0"
    If Val(txtHoverTime.Text) > 65535 Then txtHoverTime.Text = "65535"
End Sub

Private Sub txtHoverTimeHeight_Change()
    txtHoverTimeHeight.Text = CStr(Val(Rem_NonNumeric_Chr(txtHoverTimeHeight.Text)))
    If Val(txtHoverTimeHeight.Text) < 0 Then txtHoverTimeHeight.Text = "0"
    If Val(txtHoverTimeHeight.Text) > 65535 Then txtHoverTimeHeight.Text = "65535"
End Sub

Private Sub txtHoverTimeWidth_Change()
    txtHoverTimeWidth.Text = CStr(Val(Rem_NonNumeric_Chr(txtHoverTimeWidth.Text)))
    If Val(txtHoverTimeWidth.Text) < 0 Then txtHoverTimeWidth.Text = "0"
    If Val(txtHoverTimeWidth.Text) > 65535 Then txtHoverTimeWidth.Text = "65535"
End Sub

Private Sub txtSpeed_Change()
    txtSpeed.Text = CStr(Val(Rem_NonNumeric_Chr(txtSpeed.Text)))
    If Val(txtSpeed.Text) < 1 Then txtSpeed.Text = "1"
    If Val(txtSpeed.Text) > 20 Then txtSpeed.Text = "20"
End Sub

Private Sub txtWheelScrollLines_Change()
    txtWheelScrollLines.Text = CStr(Val(Rem_NonNumeric_Chr(txtWheelScrollLines.Text)))
    If Val(txtWheelScrollLines.Text) < 0 Then txtWheelScrollLines.Text = "0"
    If Val(txtWheelScrollLines.Text) > 65535 Then txtWheelScrollLines.Text = "65535"
End Sub
