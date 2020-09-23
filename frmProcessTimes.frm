VERSION 5.00
Begin VB.Form frmProcessTimes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Process Times"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6015
   Icon            =   "frmProcessTimes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   6015
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtMillisecondsK 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   44
      Top             =   3840
      Width           =   615
   End
   Begin VB.TextBox txtSecondsK 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   41
      Top             =   3600
      Width           =   615
   End
   Begin VB.TextBox txtMinutesK 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   38
      Top             =   3360
      Width           =   615
   End
   Begin VB.TextBox txtHoursK 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   35
      Top             =   3120
      Width           =   615
   End
   Begin VB.TextBox txtDaysK 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   32
      Top             =   2880
      Width           =   615
   End
   Begin VB.TextBox txtMillisecondsU 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   45
      Top             =   3840
      Width           =   615
   End
   Begin VB.TextBox txtSecondsU 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   42
      Top             =   3600
      Width           =   615
   End
   Begin VB.TextBox txtMinutesU 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   39
      Top             =   3360
      Width           =   615
   End
   Begin VB.TextBox txtHoursU 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   36
      Top             =   3120
      Width           =   615
   End
   Begin VB.TextBox txtDaysU 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   33
      Top             =   2880
      Width           =   615
   End
   Begin VB.TextBox txtMillisecondE 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5280
      TabIndex        =   28
      Top             =   2160
      Width           =   615
   End
   Begin VB.TextBox txtSecondE 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5280
      TabIndex        =   25
      Top             =   1920
      Width           =   615
   End
   Begin VB.TextBox txtMinuteE 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5280
      TabIndex        =   22
      Top             =   1680
      Width           =   615
   End
   Begin VB.TextBox txtHourE 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5280
      TabIndex        =   19
      Top             =   1440
      Width           =   615
   End
   Begin VB.TextBox txtDayE 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5280
      TabIndex        =   16
      Top             =   1200
      Width           =   615
   End
   Begin VB.TextBox txtDayOfWeekE 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   960
      Width           =   615
   End
   Begin VB.TextBox txtMonthE 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5280
      TabIndex        =   10
      Top             =   720
      Width           =   615
   End
   Begin VB.TextBox txtYearE 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5280
      TabIndex        =   7
      Top             =   480
      Width           =   615
   End
   Begin VB.TextBox txtMillisecondC 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4440
      TabIndex        =   27
      Top             =   2160
      Width           =   615
   End
   Begin VB.TextBox txtSecondC 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4440
      TabIndex        =   24
      Top             =   1920
      Width           =   615
   End
   Begin VB.TextBox txtMinuteC 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4440
      TabIndex        =   21
      Top             =   1680
      Width           =   615
   End
   Begin VB.TextBox txtHourC 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4440
      TabIndex        =   18
      Top             =   1440
      Width           =   615
   End
   Begin VB.TextBox txtDayC 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4440
      TabIndex        =   15
      Top             =   1200
      Width           =   615
   End
   Begin VB.TextBox txtDayOfWeekC 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   960
      Width           =   615
   End
   Begin VB.TextBox txtMonthC 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4440
      TabIndex        =   9
      Top             =   720
      Width           =   615
   End
   Begin VB.TextBox txtYearC 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4440
      TabIndex        =   6
      Top             =   480
      Width           =   615
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
   Begin VB.Label lblMillisecond 
      Caption         =   "Millisecond"
      Height          =   255
      Left            =   3240
      TabIndex        =   26
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label lblSecond 
      Caption         =   "Second"
      Height          =   255
      Left            =   3240
      TabIndex        =   23
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label lblMinute 
      Caption         =   "Minute"
      Height          =   255
      Left            =   3240
      TabIndex        =   20
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label lblHour 
      Caption         =   "Hour"
      Height          =   255
      Left            =   3240
      TabIndex        =   17
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label lblDay 
      Caption         =   "Day"
      Height          =   255
      Left            =   3240
      TabIndex        =   14
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label lblDayOfWeek 
      Caption         =   "Day of Week"
      Height          =   255
      Left            =   3240
      TabIndex        =   11
      Top             =   960
      Width           =   975
   End
   Begin VB.Label lblMonth 
      Caption         =   "Month"
      Height          =   255
      Left            =   3240
      TabIndex        =   8
      Top             =   720
      Width           =   975
   End
   Begin VB.Label lblYear 
      Caption         =   "Year"
      Height          =   255
      Left            =   3240
      TabIndex        =   5
      Top             =   480
      Width           =   975
   End
   Begin VB.Label lblExit 
      Caption         =   "Exit"
      Height          =   255
      Left            =   5280
      TabIndex        =   4
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lblCreation 
      Caption         =   "Creation"
      Height          =   255
      Left            =   4440
      TabIndex        =   3
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lblSeconds 
      Caption         =   "Seconds"
      Height          =   255
      Left            =   3240
      TabIndex        =   40
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label lblMinutes 
      Caption         =   "Minutes"
      Height          =   255
      Left            =   3240
      TabIndex        =   37
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label lblHours 
      Caption         =   "Hours"
      Height          =   255
      Left            =   3240
      TabIndex        =   34
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label lblDays 
      Caption         =   "Days"
      Height          =   255
      Left            =   3240
      TabIndex        =   31
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label lblKernel 
      Caption         =   "Kernel"
      Height          =   255
      Left            =   4440
      TabIndex        =   29
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label lblMilliseconds 
      Caption         =   "Milliseconds"
      Height          =   255
      Left            =   3240
      TabIndex        =   43
      Top             =   3840
      Width           =   975
   End
   Begin VB.Label lblUser 
      Caption         =   "User"
      Height          =   255
      Left            =   5280
      TabIndex        =   30
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label lblProcess 
      Caption         =   "Process"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmProcessTimes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Process() As PROCESSENTRY32
Dim lngProcess As Long

Private Sub cmdRefresh_Click()
    lstProcess.Clear
    lngProcess = 0
    Erase Process()
    
    lngProcess = Process32_Enum(Process())
    
    Dim lngIncrement As Long
    For lngIncrement = 0 To lngProcess
        lstProcess.AddItem CStr(Process(lngIncrement).th32ProcessID)
    Next lngIncrement
    
    
    txtYearC.Text = ""
    txtMonthC.Text = ""
    txtDayOfWeekC.Text = ""
    txtDayC.Text = ""
    txtHourC.Text = ""
    txtMinuteC.Text = ""
    txtSecondC.Text = ""
    txtMillisecondC.Text = ""
    txtYearE.Text = ""
    txtMonthE.Text = ""
    txtDayOfWeekE.Text = ""
    txtDayE.Text = ""
    txtHourE.Text = ""
    txtMinuteE.Text = ""
    txtSecondE.Text = ""
    txtMillisecondE.Text = ""
    txtMillisecondsK.Text = ""
    txtSecondsK.Text = ""
    txtMinutesK.Text = ""
    txtHoursK.Text = ""
    txtDaysK.Text = ""
    txtMillisecondsU.Text = ""
    txtSecondsU.Text = ""
    txtMinutesU.Text = ""
    txtHoursU.Text = ""
    txtDaysU.Text = ""
End Sub

Private Sub Form_Load()
    cmdRefresh_Click
    
    
    If Function_Exist("kernel32.dll", "CreateToolhelp32Snapshot") = False Then
        lblProcess.Enabled = False
        lstProcess.Enabled = False
        cmdRefresh.Enabled = False
    End If
    If Function_Exist("kernel32.dll", "OpenThread") = False Then
        lblYear.Enabled = False
        lblMonth.Enabled = False
        lblDayOfWeek.Enabled = False
        lblDay.Enabled = False
        lblHour.Enabled = False
        lblMinute.Enabled = False
        lblSecond.Enabled = False
        lblMillisecond.Enabled = False
        lblMilliseconds.Enabled = False
        lblSeconds.Enabled = False
        lblMinutes.Enabled = False
        lblHours.Enabled = False
        lblDays.Enabled = False
    End If
End Sub

Private Sub lstProcess_Click()
    If Function_Exist("kernel32.dll", "GetProcessTimes") = True Then
        Dim hProcess As Long
        hProcess = OpenProcess(PROCESS_QUERY_INFORMATION, False, Process(lstProcess.ListIndex).th32ProcessID): If hProcess = &H0 Then Failed "OpenProcess"
        
        
        Dim ftCreation As FILETIME
        Dim ftExit As FILETIME
        Dim ftKernel As FILETIME
        Dim ftUser As FILETIME
        If GetProcessTimes(hProcess, ftCreation, ftExit, ftKernel, ftUser) = False Then Failed "GetProcessTimes"
        
        
        Dim stCreation As SYSTEMTIME
        Dim stExit As SYSTEMTIME
        If FileTimeToSystemTime(ftCreation, stCreation) = False Then Failed "FiletimeToSystemTime"
        If FileTimeToSystemTime(ftExit, stExit) = False Then Failed "FiletimeToSystemTime"
        
        With stCreation
            txtYearC.Text = CStr(.wYear)
            txtMonthC.Text = CStr(.wMonth)
            txtDayOfWeekC.Text = CStr(.wDayOfWeek)
            txtDayC.Text = CStr(.wDay)
            txtHourC.Text = CStr(.wHour)
            txtMinuteC.Text = CStr(.wMinute)
            txtSecondC.Text = CStr(.wSecond)
            txtMillisecondC.Text = CStr(.wMilliseconds)
        End With
        With stExit
            txtYearE.Text = CStr(.wYear)
            txtMonthE.Text = CStr(.wMonth)
            txtDayOfWeekE.Text = CStr(.wDayOfWeek)
            txtDayE.Text = CStr(.wDay)
            txtHourE.Text = CStr(.wHour)
            txtMinuteE.Text = CStr(.wMinute)
            txtSecondE.Text = CStr(.wSecond)
            txtMillisecondE.Text = CStr(.wMilliseconds)
        End With
        
        
        Dim dblKernel As Double
        Dim dblUser As Double
        dblKernel = CLargeInt(ftKernel.dwLowDateTime, ftKernel.dwHighDateTime)
        dblUser = CLargeInt(ftUser.dwLowDateTime, ftUser.dwHighDateTime)
        
        
        dblKernel = Round(dblKernel / 10000, 0)
        txtMillisecondsK.Text = CStr(dblKernel - ((dblKernel \ 1000) * 1000))
        dblKernel = Round(dblKernel / 1000, 0)
        txtSecondsK.Text = CStr(dblKernel - (dblKernel \ 60) * 60)
        dblKernel = Round(dblKernel / 60, 0)
        txtMinutesK.Text = CStr(dblKernel - (dblKernel \ 60) * 60)
        dblKernel = Round(dblKernel / 60, 0)
        txtHoursK.Text = CStr(dblKernel - (dblKernel \ 24) * 24)
        dblKernel = Round(dblKernel / 24, 0)
        txtDaysK.Text = CStr(dblKernel)
        
        dblUser = Round(dblUser / 10000, 0)
        txtMillisecondsU.Text = CStr(dblUser - (dblUser \ 1000) * 1000)
        dblUser = Round(dblUser / 1000, 0)
        txtSecondsU.Text = CStr(dblUser - (dblUser \ 60) * 60)
        dblUser = Round(dblUser / 60, 0)
        txtMinutesU.Text = CStr(dblUser - (dblUser \ 60) * 60)
        dblUser = Round(dblUser / 60, 0)
        txtHoursU.Text = CStr(dblUser - (dblUser \ 24) * 24)
        dblUser = Round(dblUser / 24, 0)
        txtDaysU.Text = CStr(dblUser)
        
        
        If CloseHandle(hProcess) = False Then Failed "CloseHandle"
    End If
End Sub
