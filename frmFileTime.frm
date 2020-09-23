VERSION 5.00
Begin VB.Form frmFileTime 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "File Time"
   ClientHeight    =   5250
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7215
   Icon            =   "frmFileTime.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   7215
   StartUpPosition =   3  'Windows Default
   Begin VB.DriveListBox drvFileTime 
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   4800
      Width           =   2175
   End
   Begin VB.TextBox txtSelected 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   360
      Width           =   6975
   End
   Begin VB.DirListBox dirFileTime 
      Height          =   1890
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   2175
   End
   Begin VB.FileListBox fileFileTime 
      Height          =   2040
      Hidden          =   -1  'True
      Left            =   120
      System          =   -1  'True
      TabIndex        =   3
      Top             =   2760
      Width           =   2175
   End
   Begin VB.TextBox txtMillisecondLW 
      Height          =   285
      Left            =   6120
      TabIndex        =   39
      Top             =   4440
      Width           =   975
   End
   Begin VB.TextBox txtMillisecondLA 
      Height          =   285
      Left            =   4920
      TabIndex        =   38
      Top             =   4440
      Width           =   975
   End
   Begin VB.TextBox txtMillisecondCT 
      Height          =   285
      Left            =   3720
      TabIndex        =   37
      Top             =   4440
      Width           =   975
   End
   Begin VB.TextBox txtSecondLW 
      Height          =   285
      Left            =   6120
      TabIndex        =   35
      Top             =   4080
      Width           =   975
   End
   Begin VB.TextBox txtSecondLA 
      Height          =   285
      Left            =   4920
      TabIndex        =   34
      Top             =   4080
      Width           =   975
   End
   Begin VB.TextBox txtSecondCT 
      Height          =   285
      Left            =   3720
      TabIndex        =   33
      Top             =   4080
      Width           =   975
   End
   Begin VB.TextBox txtMinuteLW 
      Height          =   285
      Left            =   6120
      TabIndex        =   31
      Top             =   3720
      Width           =   975
   End
   Begin VB.TextBox txtMinuteLA 
      Height          =   285
      Left            =   4920
      TabIndex        =   30
      Top             =   3720
      Width           =   975
   End
   Begin VB.TextBox txtMinuteCT 
      Height          =   285
      Left            =   3720
      TabIndex        =   29
      Top             =   3720
      Width           =   975
   End
   Begin VB.TextBox txtHourLW 
      Height          =   285
      Left            =   6120
      TabIndex        =   27
      Top             =   3360
      Width           =   975
   End
   Begin VB.TextBox txtHourLA 
      Height          =   285
      Left            =   4920
      TabIndex        =   26
      Top             =   3360
      Width           =   975
   End
   Begin VB.TextBox txtHourCT 
      Height          =   285
      Left            =   3720
      TabIndex        =   25
      Top             =   3360
      Width           =   975
   End
   Begin VB.TextBox txtDayLW 
      Height          =   285
      Left            =   6120
      TabIndex        =   23
      Top             =   3000
      Width           =   975
   End
   Begin VB.TextBox txtDayLA 
      Height          =   285
      Left            =   4920
      TabIndex        =   22
      Top             =   3000
      Width           =   975
   End
   Begin VB.TextBox txtDayCT 
      Height          =   285
      Left            =   3720
      TabIndex        =   21
      Top             =   3000
      Width           =   975
   End
   Begin VB.TextBox txtDayOfWeekLW 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6120
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   2640
      Width           =   975
   End
   Begin VB.TextBox txtDayOfWeekLA 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   2640
      Width           =   975
   End
   Begin VB.TextBox txtDayOfWeekCT 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   2640
      Width           =   975
   End
   Begin VB.TextBox txtMonthLW 
      Height          =   285
      Left            =   6120
      TabIndex        =   15
      Top             =   2280
      Width           =   975
   End
   Begin VB.TextBox txtMonthLA 
      Height          =   285
      Left            =   4920
      TabIndex        =   14
      Top             =   2280
      Width           =   975
   End
   Begin VB.TextBox txtMonthCT 
      Height          =   285
      Left            =   3720
      TabIndex        =   13
      Top             =   2280
      Width           =   975
   End
   Begin VB.TextBox txtYearLW 
      Height          =   285
      Left            =   6120
      TabIndex        =   11
      Top             =   1920
      Width           =   975
   End
   Begin VB.TextBox txtYearLA 
      Height          =   285
      Left            =   4920
      TabIndex        =   10
      Top             =   1920
      Width           =   975
   End
   Begin VB.TextBox txtYearCT 
      Height          =   285
      Left            =   3720
      TabIndex        =   9
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   350
      Left            =   6120
      TabIndex        =   40
      Top             =   4800
      Width           =   975
   End
   Begin VB.Label lblSelected 
      Caption         =   "Selected"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label lblMillisecond 
      Caption         =   "Millisecond"
      Height          =   255
      Left            =   2520
      TabIndex        =   36
      Top             =   4440
      Width           =   975
   End
   Begin VB.Label lblSecond 
      Caption         =   "Second"
      Height          =   255
      Left            =   2520
      TabIndex        =   32
      Top             =   4080
      Width           =   975
   End
   Begin VB.Label lblMinute 
      Caption         =   "Minute"
      Height          =   255
      Left            =   2520
      TabIndex        =   28
      Top             =   3720
      Width           =   975
   End
   Begin VB.Label lblHour 
      Caption         =   "Hour"
      Height          =   255
      Left            =   2520
      TabIndex        =   24
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label lblDay 
      Caption         =   "Day"
      Height          =   255
      Left            =   2520
      TabIndex        =   20
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label lblDayOfWeek 
      Caption         =   "Day of Week"
      Height          =   255
      Left            =   2520
      TabIndex        =   16
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label lblMonth 
      Caption         =   "Month"
      Height          =   255
      Left            =   2520
      TabIndex        =   12
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label lblYear 
      Caption         =   "Year"
      Height          =   255
      Left            =   2520
      TabIndex        =   8
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label lblLastWrite 
      Caption         =   "Last Write"
      Height          =   255
      Left            =   6120
      TabIndex        =   7
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label lblLastAccess 
      Caption         =   "Last Access"
      Height          =   255
      Left            =   4920
      TabIndex        =   6
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label lblCreation 
      Caption         =   "Creation"
      Height          =   255
      Left            =   3720
      TabIndex        =   5
      Top             =   1560
      Width           =   975
   End
End
Attribute VB_Name = "frmFileTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdApply_Click()
    If txtSelected.Text = "" Then Exit Sub
    

    Dim hFile As Long
    
    Dim ftCreationTime As FILETIME
    Dim ftLastAccess As FILETIME
    Dim ftLastWrite As FILETIME
    Dim stCreationTime As SYSTEMTIME
    Dim stLastAccess As SYSTEMTIME
    Dim stLastWrite As SYSTEMTIME
    
    
    hFile = CreateFile(txtSelected.Text, GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal &H0, OPEN_EXISTING, &H0, &H0): If hFile = INVALID_HANDLE_VALUE Then Failed "CreateFile"
    
    
    With stCreationTime
        .wYear = Val(txtYearCT.Text)
        .wMonth = Val(txtMonthCT.Text)
        .wDayOfWeek = Val(txtDayOfWeekCT.Text)
        .wDay = Val(txtDayCT.Text)
        .wHour = Val(txtHourCT.Text)
        .wMinute = Val(txtMinuteCT.Text)
        .wSecond = Val(txtSecondCT.Text)
        .wMilliseconds = Val(txtMillisecondCT.Text)
    End With
    With stLastAccess
        .wYear = Val(txtYearLA.Text)
        .wMonth = Val(txtMonthLA.Text)
        .wDayOfWeek = Val(txtDayOfWeekLA.Text)
        .wDay = Val(txtDayLA.Text)
        .wHour = Val(txtHourLA.Text)
        .wMinute = Val(txtMinuteLA.Text)
        .wSecond = Val(txtSecondLA.Text)
        .wMilliseconds = Val(txtMillisecondLA.Text)
    End With
    With stLastWrite
        .wYear = Val(txtYearLW.Text)
        .wMonth = Val(txtMonthLW.Text)
        .wDayOfWeek = Val(txtDayOfWeekLW.Text)
        .wDay = Val(txtDayLW.Text)
        .wHour = Val(txtHourLW.Text)
        .wMinute = Val(txtMinuteLW.Text)
        .wSecond = Val(txtSecondLW.Text)
        .wMilliseconds = Val(txtMillisecondLW.Text)
    End With
    
    
    If SystemTimeToFileTime(stCreationTime, ftCreationTime) = False Then Failed "SystemTimeToFileTime"
    If SystemTimeToFileTime(stLastAccess, ftLastAccess) = False Then Failed "SystemTimeToFileTime"
    If SystemTimeToFileTime(stLastWrite, ftLastWrite) = False Then Failed "SystemTimeToFileTime"
    
    If SetFileTime(hFile, ftCreationTime, ftLastAccess, ftLastWrite) = False Then Failed "SetFileTime"
    If CloseHandle(hFile) = False Then Failed "CloseHandle"
End Sub

Private Sub dirFileTime_Change()
    fileFileTime.Path = dirFileTime.Path
End Sub

Private Sub drvFileTime_Change()
    On Error Resume Next
    dirFileTime.Path = drvFileTime.Drive
    On Error GoTo 0
End Sub

Private Sub fileFileTime_Click()
    txtSelected.Text = Fix_Dir(dirFileTime.Path) & "\" & fileFileTime.FileName
    Process txtSelected.Text
End Sub

Private Sub Process(strFileName As String)
    Dim hFile As Long
    
    Dim ftCreationTime As FILETIME
    Dim ftLastAccess As FILETIME
    Dim ftLastWrite As FILETIME
    Dim stCreationTime As SYSTEMTIME
    Dim stLastAccess As SYSTEMTIME
    Dim stLastWrite As SYSTEMTIME


    hFile = CreateFile(strFileName, GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal &H0, OPEN_EXISTING, &H0, &H0): If hFile = INVALID_HANDLE_VALUE Then Failed "CreateFile"
    
    If GetFileTime(hFile, ftCreationTime, ftLastAccess, ftLastWrite) = False Then Failed "GetFileTime"
    
    If FileTimeToSystemTime(ftCreationTime, stCreationTime) = False Then Failed "FiletimeToSystemTime"
    If FileTimeToSystemTime(ftLastAccess, stLastAccess) = False Then Failed "FiletimeToSystemTime"
    If FileTimeToSystemTime(ftLastWrite, stLastWrite) = False Then Failed "FiletimeToSystemTime"


    With stCreationTime
        txtYearCT.Text = CStr(.wYear)
        txtMonthCT.Text = CStr(.wMonth)
        txtDayOfWeekCT.Text = CStr(.wDayOfWeek)
        txtDayCT.Text = CStr(.wDay)
        txtHourCT.Text = CStr(.wHour)
        txtMinuteCT.Text = CStr(.wMinute)
        txtSecondCT.Text = CStr(.wSecond)
        txtMillisecondCT.Text = CStr(.wMilliseconds)
    End With
    With stLastAccess
        txtYearLA.Text = CStr(.wYear)
        txtMonthLA.Text = CStr(.wMonth)
        txtDayOfWeekLA.Text = CStr(.wDayOfWeek)
        txtDayLA.Text = CStr(.wDay)
        txtHourLA.Text = CStr(.wHour)
        txtMinuteLA.Text = CStr(.wMinute)
        txtSecondLA.Text = CStr(.wSecond)
        txtMillisecondLA.Text = CStr(.wMilliseconds)
    End With
    With stLastWrite
        txtYearLW.Text = CStr(.wYear)
        txtMonthLW.Text = CStr(.wMonth)
        txtDayOfWeekLW.Text = CStr(.wDayOfWeek)
        txtDayLW.Text = CStr(.wDay)
        txtHourLW.Text = CStr(.wHour)
        txtMinuteLW.Text = CStr(.wMinute)
        txtSecondLW.Text = CStr(.wSecond)
        txtMillisecondLW.Text = CStr(.wMilliseconds)
    End With
    
    
    If CloseHandle(hFile) = False Then Failed "CloseHandle"
End Sub

Private Sub txtDayCT_Change()
    txtDayCT.Text = CStr(Val(Rem_NonNumeric_Chr(txtDayCT.Text)))
    If Val(txtDayCT.Text) < 0 Then txtDayCT.Text = "0"
    If Val(txtDayCT.Text) > 31 Then txtDayCT.Text = "31"
End Sub

Private Sub txtDayLA_Change()
    txtDayLA.Text = CStr(Val(Rem_NonNumeric_Chr(txtDayLA.Text)))
    If Val(txtDayLA.Text) < 0 Then txtDayLA.Text = "0"
    If Val(txtDayLA.Text) > 31 Then txtDayLA.Text = "31"
End Sub

Private Sub txtDayLW_Change()
    txtDayLW.Text = CStr(Val(Rem_NonNumeric_Chr(txtDayLW.Text)))
    If Val(txtDayLW.Text) < 0 Then txtDayLW.Text = "0"
    If Val(txtDayLW.Text) > 31 Then txtDayLW.Text = "31"
End Sub

Private Sub txtHourCT_Change()
    txtHourCT.Text = CStr(Val(Rem_NonNumeric_Chr(txtHourCT.Text)))
    If Val(txtHourCT.Text) < 0 Then txtHourCT.Text = "0"
    If Val(txtHourCT.Text) > 23 Then txtHourCT.Text = "23"
End Sub

Private Sub txtHourLA_Change()
    txtHourLA.Text = CStr(Val(Rem_NonNumeric_Chr(txtHourLA.Text)))
    If Val(txtHourLA.Text) < 0 Then txtHourLA.Text = "0"
    If Val(txtHourLA.Text) > 23 Then txtHourLA.Text = "23"
End Sub

Private Sub txtHourLW_Change()
    txtHourLW.Text = CStr(Val(Rem_NonNumeric_Chr(txtHourLW.Text)))
    If Val(txtHourLW.Text) < 0 Then txtHourLW.Text = "0"
    If Val(txtHourLW.Text) > 23 Then txtHourLW.Text = "23"
End Sub

Private Sub txtMillisecondCT_Change()
    txtMillisecondCT.Text = CStr(Val(Rem_NonNumeric_Chr(txtMillisecondCT.Text)))
    If Val(txtMillisecondCT.Text) < 0 Then txtMillisecondCT.Text = "0"
    If Val(txtMillisecondCT.Text) > 999 Then txtMillisecondCT.Text = "999"
End Sub

Private Sub txtMillisecondLA_Change()
    txtMillisecondLA.Text = CStr(Val(Rem_NonNumeric_Chr(txtMillisecondLA.Text)))
    If Val(txtMillisecondLA.Text) < 0 Then txtMillisecondLA.Text = "0"
    If Val(txtMillisecondLA.Text) > 999 Then txtMillisecondLA.Text = "999"
End Sub

Private Sub txtMillisecondLW_Change()
    txtMillisecondLW.Text = CStr(Val(Rem_NonNumeric_Chr(txtMillisecondLW.Text)))
    If Val(txtMillisecondLW.Text) < 0 Then txtMillisecondLW.Text = "0"
    If Val(txtMillisecondLW.Text) > 999 Then txtMillisecondLW.Text = "999"
End Sub

Private Sub txtMinuteCT_Change()
    txtMinuteCT.Text = CStr(Val(Rem_NonNumeric_Chr(txtMinuteCT.Text)))
    If Val(txtMinuteCT.Text) < 0 Then txtMinuteCT.Text = "0"
    If Val(txtMinuteCT.Text) > 59 Then txtMinuteCT.Text = "59"
End Sub

Private Sub txtMinuteLA_Change()
    txtMinuteLA.Text = CStr(Val(Rem_NonNumeric_Chr(txtMinuteLA.Text)))
    If Val(txtMinuteLA.Text) < 0 Then txtMinuteLA.Text = "0"
    If Val(txtMinuteLA.Text) > 59 Then txtMinuteLA.Text = "59"
End Sub

Private Sub txtMinuteLW_Change()
    txtMinuteLW.Text = CStr(Val(Rem_NonNumeric_Chr(txtMinuteLW.Text)))
    If Val(txtMinuteLW.Text) < 0 Then txtMinuteLW.Text = "0"
    If Val(txtMinuteLW.Text) > 59 Then txtMinuteLW.Text = "59"
End Sub

Private Sub txtMonthCT_Change()
    txtMonthCT.Text = CStr(Val(Rem_NonNumeric_Chr(txtMonthCT.Text)))
    If Val(txtMonthCT.Text) < 0 Then txtMonthCT.Text = "0"
    If Val(txtMonthCT.Text) > 12 Then txtMonthCT.Text = "12"
End Sub

Private Sub txtMonthLA_Change()
    txtMonthLA.Text = CStr(Val(Rem_NonNumeric_Chr(txtMonthLA.Text)))
    If Val(txtMonthLA.Text) < 0 Then txtMonthLA.Text = "0"
    If Val(txtMonthLA.Text) > 12 Then txtMonthLA.Text = "12"
End Sub

Private Sub txtMonthLW_Change()
    txtMonthLW.Text = CStr(Val(Rem_NonNumeric_Chr(txtMonthLW.Text)))
    If Val(txtMonthLW.Text) < 0 Then txtMonthLW.Text = "0"
    If Val(txtMonthLW.Text) > 12 Then txtMonthLW.Text = "12"
End Sub

Private Sub txtSecondCT_Change()
    txtSecondCT.Text = CStr(Val(Rem_NonNumeric_Chr(txtSecondCT.Text)))
    If Val(txtSecondCT.Text) < 0 Then txtSecondCT.Text = "0"
    If Val(txtSecondCT.Text) > 59 Then txtSecondCT.Text = "59"
End Sub

Private Sub txtSecondLA_Change()
    txtSecondLA.Text = CStr(Val(Rem_NonNumeric_Chr(txtSecondLA.Text)))
    If Val(txtSecondLA.Text) < 0 Then txtSecondLA.Text = "0"
    If Val(txtSecondLA.Text) > 59 Then txtSecondLA.Text = "59"
End Sub

Private Sub txtSecondLW_Change()
    txtSecondLW.Text = CStr(Val(Rem_NonNumeric_Chr(txtSecondLW.Text)))
    If Val(txtSecondLW.Text) < 0 Then txtSecondLW.Text = "0"
    If Val(txtSecondLW.Text) > 59 Then txtSecondLW.Text = "59"
End Sub

Private Sub txtYearCT_Change()
    txtYearCT.Text = CStr(Val(Rem_NonNumeric_Chr(txtYearCT.Text)))
    If Val(txtYearCT.Text) < 0 Then txtYearCT.Text = "0"
    If Val(txtYearCT.Text) > 65535 Then txtYearCT.Text = "65535"
End Sub

Private Sub txtYearLA_Change()
    txtYearLA.Text = CStr(Val(Rem_NonNumeric_Chr(txtYearLA.Text)))
    If Val(txtYearLA.Text) < 0 Then txtYearLA.Text = "0"
    If Val(txtYearLA.Text) > 65535 Then txtYearLA.Text = "65535"
End Sub

Private Sub txtYearLW_Change()
    txtYearLW.Text = CStr(Val(Rem_NonNumeric_Chr(txtYearLW.Text)))
    If Val(txtYearLW.Text) < 0 Then txtYearLW.Text = "0"
    If Val(txtYearLW.Text) > 65535 Then txtYearLW.Text = "65535"
End Sub
