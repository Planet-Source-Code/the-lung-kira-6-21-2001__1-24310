VERSION 5.00
Begin VB.Form frmLocalesTime 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Locales - Time"
   ClientHeight    =   3420
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6615
   Icon            =   "frmLocalesTime.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   6615
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtTimeMarkerUse 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   2160
      Width           =   1935
   End
   Begin VB.TextBox txtTimeFormatting 
      Height          =   285
      Left            =   4560
      TabIndex        =   14
      Top             =   1560
      Width           =   1935
   End
   Begin VB.TextBox txtTimeSeperator 
      Height          =   285
      Left            =   4560
      TabIndex        =   20
      Top             =   2520
      Width           =   1935
   End
   Begin VB.TextBox txtPMDesignator 
      Height          =   285
      Left            =   4560
      TabIndex        =   10
      Top             =   840
      Width           =   1935
   End
   Begin VB.TextBox txtAMDesignator 
      Height          =   285
      Left            =   4560
      TabIndex        =   6
      Top             =   120
      Width           =   1935
   End
   Begin VB.CheckBox chkHourLeadingZeros 
      Height          =   255
      Left            =   6240
      TabIndex        =   8
      Top             =   480
      Width           =   255
   End
   Begin VB.TextBox txtTimeMarkerPosition 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   1920
      Width           =   1935
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   350
      Left            =   5520
      TabIndex        =   21
      Top             =   3000
      Width           =   975
   End
   Begin VB.ComboBox cboTimeFormat 
      Height          =   315
      Left            =   4560
      TabIndex        =   12
      Top             =   1200
      Width           =   1935
   End
   Begin VB.ListBox lstLocales 
      Height          =   1035
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   1935
   End
   Begin VB.ComboBox cboDisplay 
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   1800
      Width           =   1935
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   350
      Left            =   1080
      TabIndex        =   4
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label lblTimeFormatting 
      Caption         =   "Time Formatting"
      Height          =   255
      Left            =   2280
      TabIndex        =   13
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label lblTimeSeperator 
      Caption         =   "Time Seperator"
      Height          =   255
      Left            =   2280
      TabIndex        =   19
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Label lblPMDesignator 
      Caption         =   "PM Designator"
      Height          =   255
      Left            =   2280
      TabIndex        =   9
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label lblAMDesignator 
      Caption         =   "AM Designator"
      Height          =   255
      Left            =   2280
      TabIndex        =   5
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label lblHourLeadingZeros 
      Caption         =   "Hour Leading Zeros"
      Height          =   255
      Left            =   2280
      TabIndex        =   7
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label lblTimeMarkerUse 
      Caption         =   "Time Marker Use"
      Height          =   255
      Left            =   2280
      TabIndex        =   17
      Top             =   2160
      Width           =   2055
   End
   Begin VB.Label lblTimeMarkerPosition 
      Caption         =   "Time Marker Position"
      Height          =   255
      Left            =   2280
      TabIndex        =   15
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label lblTimeFormat 
      Caption         =   "Time Format"
      Height          =   255
      Left            =   2280
      TabIndex        =   11
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label lblLocales 
      Caption         =   "Locales"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label lblDisplay 
      Caption         =   "Display"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   1815
   End
End
Attribute VB_Name = "frmLocalesTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboDisplay_Click()
    cmdRefresh_Click
End Sub

Private Sub cmdApply_Click()
    If lstLocales.ListIndex = -1 Then Exit Sub
    
    
    Dim lngLocale As Long
    lngLocale = strtoul_(lstLocales.List(lstLocales.ListIndex), 16)
    
    
    If cboTimeFormat.ListIndex > -1 Then
        If SetLocaleInfo(lngLocale, LOCALE_ITIME, CStr(cboTimeFormat.ListIndex)) = False Then Failed "SetLocaleInfo"
    End If
    If SetLocaleInfo(lngLocale, LOCALE_S1159, txtAMDesignator.Text) = False Then Failed "SetLocaleInfo"
    If SetLocaleInfo(lngLocale, LOCALE_S2359, txtPMDesignator.Text) = False Then Failed "SetLocaleInfo"
    If SetLocaleInfo(lngLocale, LOCALE_STIME, txtTimeSeperator.Text) = False Then Failed "SetLocaleInfo"
    If SetLocaleInfo(lngLocale, LOCALE_STIMEFORMAT, txtTimeFormatting.Text) = False Then Failed "SetLocaleInfo"
End Sub

Private Sub cmdRefresh_Click()
    LocaleListNum = 0
    Erase LocaleList
    lstLocales.Clear
    
    cboTimeFormat.ListIndex = -1
    txtTimeMarkerPosition.Text = ""
    'txtTimeMarkerUse.Text = ""
    chkHourLeadingZeros.value = 0
    txtAMDesignator.Text = ""
    txtPMDesignator.Text = ""
    txtTimeSeperator.Text = ""
    txtTimeFormatting.Text = ""
    
    
    Dim lngFlags As Long
    Select Case cboDisplay.ListIndex
        Case 0: lngFlags = LCID_INSTALLED
        Case 1: lngFlags = LCID_SUPPORTED
        Case 2: lngFlags = LCID_ALTERNATE_SORTS
        Case 3: lngFlags = LCID_ALTERNATE_SORTS Or LCID_INSTALLED
        Case 4: lngFlags = LCID_ALTERNATE_SORTS Or LCID_SUPPORTED
    End Select
    
    If EnumSystemLocales(AddressOf EnumLocalesProc, lngFlags) = False Then Failed "EnumSystemLocales"
    
    Dim lngIncrement As Long
    For lngIncrement = 1 To LocaleListNum
        lstLocales.AddItem LocaleList(lngIncrement)
    Next lngIncrement
    
    LocaleListNum = 0
    Erase LocaleList
End Sub

Private Sub Form_Load()
    With cboDisplay
        .AddItem "Installed"
        .AddItem "Supported"
        .AddItem "Alternate Sorts"
        .AddItem "Alternate Sorts + Installed"
        .AddItem "Alternate Sorts + Supported"
    End With
    With cboTimeFormat
        .AddItem "AM / PM 12-hour format"
        .AddItem "24-hour format"
    End With
    
    
    lblTimeMarkerUse.Enabled = False
    
    cmdRefresh_Click
End Sub

Private Sub lstLocales_Click()
    Dim lngLocale As Long
    lngLocale = strtoul_(lstLocales.List(lstLocales.ListIndex), 16)
    
    
    cboTimeFormat.ListIndex = Val(Get_LocaleInfo(lngLocale, LOCALE_ITIME))
    
    Select Case Val(Get_LocaleInfo(lngLocale, LOCALE_ITIMEMARKPOSN))
        Case 0: txtTimeMarkerPosition.Text = "Use as suffix"
        Case 1: txtTimeMarkerPosition.Text = "Use as prefix"
        Case Else: txtTimeMarkerPosition.Text = ""
    End Select
    'Select Case Val(Get_LocaleInfo(lngLocale, LOCALE_ITIMEMARKERUSE))
    '    Case 0: txtTimeMarkerUse.Text = "Use with 12-hour clock"
    '    Case 1: txtTimeMarkerUse.Text = "Use with 24-hour clock"
    '    Case 2: txtTimeMarkerUse.Text = "Use with both 12-hour and 24-hour clocks"
    '    Case 3: txtTimeMarkerUse.Text = "Never use"
    '    Case Else: txtTimeMarkerUse.Text = ""
    'End Select
    
    chkHourLeadingZeros.value = Val(Get_LocaleInfo(lngLocale, LOCALE_ITLZERO))
    txtAMDesignator.Text = Get_LocaleInfo(lngLocale, LOCALE_S1159)
    txtPMDesignator.Text = Get_LocaleInfo(lngLocale, LOCALE_S2359)
    txtTimeSeperator.Text = Get_LocaleInfo(lngLocale, LOCALE_STIME)
    txtTimeFormatting.Text = Get_LocaleInfo(lngLocale, LOCALE_STIMEFORMAT)
End Sub

Private Sub txtAMDesignator_Change()
    If Len(txtAMDesignator.Text) > 9 Then
        txtAMDesignator.Text = Left$(txtAMDesignator.Text, 9)
    End If
End Sub

Private Sub txtPMDesignator_Change()
    If Len(txtPMDesignator.Text) > 9 Then
        txtPMDesignator.Text = Left$(txtPMDesignator.Text, 9)
    End If
End Sub

Private Sub txtTimeFormatting_Change()
    If Len(txtTimeFormatting.Text) > 80 Then
        txtTimeFormatting.Text = Left$(txtTimeFormatting.Text, 80)
    End If
End Sub

Private Sub txtTimeSeperator_Change()
    If Len(txtTimeSeperator.Text) > 4 Then
        txtTimeSeperator.Text = Left$(txtTimeSeperator.Text, 4)
    End If
End Sub
