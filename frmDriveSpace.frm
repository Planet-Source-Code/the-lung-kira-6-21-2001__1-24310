VERSION 5.00
Begin VB.Form frmDriveSpace 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Drive Space"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4455
   Icon            =   "frmDriveSpace.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   4455
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtFreeClustersAvailable 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   2280
      Width           =   1575
   End
   Begin VB.TextBox txtTotalFreeClusters 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   2040
      Width           =   1575
   End
   Begin VB.TextBox txtFreeSectorsAvailable 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   1440
      Width           =   1575
   End
   Begin VB.TextBox txtTotalFreeSectors 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   1200
      Width           =   1575
   End
   Begin VB.TextBox txtTotalSectors 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   960
      Width           =   1575
   End
   Begin VB.TextBox txtTotalClusters 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   1800
      Width           =   1575
   End
   Begin VB.TextBox txtFreeSpaceAvailablePercentage 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   600
      Width           =   495
   End
   Begin VB.TextBox txtFreeSpaceAvailable 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   600
      Width           =   1575
   End
   Begin VB.TextBox txtTotalFreeSpace 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   360
      Width           =   1575
   End
   Begin VB.TextBox txtTotalSpace 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin VB.ComboBox cboDrive 
      Height          =   315
      Left            =   2520
      TabIndex        =   21
      Top             =   2760
      Width           =   1815
   End
   Begin VB.ComboBox cboRound 
      Height          =   315
      Left            =   2520
      TabIndex        =   25
      Top             =   3480
      Width           =   1815
   End
   Begin VB.ComboBox cboOutput 
      Height          =   315
      Left            =   2520
      TabIndex        =   23
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Timer timerDriveSpace 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1440
      Top             =   2760
   End
   Begin VB.TextBox txtTotalFreeSpacePercentage 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   360
      Width           =   495
   End
   Begin VB.Label lblTotalClusters 
      Caption         =   "Total Clusters"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label lblTotalSectors 
      Caption         =   "Total Sectors"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label lblFreeSectorsAvailable 
      Caption         =   "Free Sectors Available"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label lblTotalFreeSectors 
      Caption         =   "Total Free Sectors"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label lblFreeClustersAvailable 
      Caption         =   "Free Clusters Available"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label lblTotalFreeClusters 
      Caption         =   "Total Free Clusters"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label lblTotalSpace 
      Caption         =   "Total Space"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label lblTotalFreeSpace 
      Caption         =   "Total Free Space"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label lblFreeSpaceAvailable 
      Caption         =   "Free Space Available"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label lblDrive 
      Caption         =   "Drive"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label lblRound 
      Caption         =   "Round"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Label lblOutput 
      Caption         =   "Output"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   3120
      Width           =   1095
   End
End
Attribute VB_Name = "frmDriveSpace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboDrive_Click()
    timerDriveSpace.Enabled = True
    timerDriveSpace_Timer
End Sub

Private Sub Form_Load()
    Dim strDrives As String
    Dim lngIncrement As Long
    
    strDrives = Left$(StrReverse(ltoa_(GetLogicalDrives, 2)) & String$(32, "0"), 32)
    
    With cboDrive
        For lngIncrement = 1 To Len(strDrives)
            If Mid$(strDrives, lngIncrement, 1) = "1" Then
                .AddItem Chr$(&H40 + lngIncrement) & ":\"
            End If
        Next lngIncrement
    End With
    
    
    With cboOutput
        .AddItem "Bytes"
        .AddItem "Kilobytes"
        .AddItem "Megabytes"
        .AddItem "Gigabytes"
        .AddItem "Terabytes"
    End With
    
    With cboRound
        .AddItem "0"
        .AddItem "1"
        .AddItem "2"
        .AddItem "3"
        .AddItem "4"
        .AddItem "5"
    End With
    
    
    cboOutput.ListIndex = GetRegSetting(HKEY_CURRENT_USER, "Software\Kira\DriveSpace", "Output")
    cboRound.ListIndex = GetRegSetting(HKEY_CURRENT_USER, "Software\Kira\DriveSpace", "Round")
    
    
    If Function_Exist("kernel32.dll", "GetDiskFreeSpaceExA") = False Then
        lblTotalSpace.Enabled = False
        lblTotalFreeSpace.Enabled = False
        lblFreeSpaceAvailable.Enabled = False
        lblTotalSectors.Enabled = False
        lblTotalFreeSectors.Enabled = False
        lblFreeSectorsAvailable.Enabled = False
        lblTotalClusters.Enabled = False
        lblTotalFreeClusters.Enabled = False
        lblFreeClustersAvailable.Enabled = False
        lblDrive.Enabled = False
        cboDrive.Enabled = False
        lblOutput.Enabled = False
        cboOutput.Enabled = False
        lblRound.Enabled = False
        cboRound.Enabled = False
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    timerDriveSpace.Enabled = False

    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\DriveSpace", "Output", cboOutput.ListIndex, REG_DWORD
    SaveRegSetting HKEY_CURRENT_USER, "Software\Kira\DriveSpace", "Round", cboRound.ListIndex, REG_DWORD
End Sub

Private Sub timerDriveSpace_Timer()
    Dim dblFreeBytesAvailable As Double
    Dim dblTotalNumberOfBytes As Double
    Dim dblTotalNumberOfFreeBytes As Double
    
    If Len(cboDrive.List(cboDrive.ListIndex)) > 0 Then
        If Get_DiskFreeSpaceEx(cboDrive.List(cboDrive.ListIndex), dblFreeBytesAvailable, dblTotalNumberOfBytes, dblTotalNumberOfFreeBytes) = False Then
            timerDriveSpace.Enabled = False
        Else
            timerDriveSpace.Enabled = True
        End If
    End If
    
    
    txtTotalSpace.Text = CStr(Round(dblTotalNumberOfBytes / (1024 ^ cboOutput.ListIndex), cboRound.ListIndex))
    txtTotalFreeSpace.Text = CStr(Round(dblTotalNumberOfFreeBytes / (1024 ^ cboOutput.ListIndex), cboRound.ListIndex))
    txtFreeSpaceAvailable.Text = CStr(Round(dblFreeBytesAvailable / (1024 ^ cboOutput.ListIndex), cboRound.ListIndex))
    
    txtTotalFreeSpacePercentage.Text = CStr(Percentage(dblTotalNumberOfFreeBytes, dblTotalNumberOfBytes, 0)) & "%"
    txtFreeSpaceAvailablePercentage.Text = CStr(Percentage(dblFreeBytesAvailable, dblTotalNumberOfBytes, 0)) & "%"
    
    
    Dim lngSectorsPerCluster As Long
    Dim lngBytesPerSector As Long
    Dim lngNumberOfFreeClusters As Long
    Dim lngTotalNumberOfClusters As Long
    
    If GetDiskFreeSpace(cboDrive.List(cboDrive.ListIndex), lngSectorsPerCluster, lngBytesPerSector, lngNumberOfFreeClusters, lngTotalNumberOfClusters) = False Then
        Failed "GetDiskFreeSpace"
        timerDriveSpace.Enabled = False
    Else
        timerDriveSpace.Enabled = True
    End If
    
    txtFreeSectorsAvailable.Text = "0"
    txtFreeClustersAvailable.Text = "0"
    txtTotalFreeSectors.Text = "0"
    txtTotalFreeClusters.Text = "0"
    txtTotalSectors.Text = "0"
    txtTotalClusters.Text = "0"
    
    If dblFreeBytesAvailable > 0 Then
        If lngBytesPerSector > 0 Then
            txtFreeSectorsAvailable.Text = CStr(Round(dblFreeBytesAvailable / lngBytesPerSector, 0))
            
            If lngSectorsPerCluster > 0 Then
                txtFreeClustersAvailable.Text = CStr(Round((dblFreeBytesAvailable / lngBytesPerSector) / lngSectorsPerCluster, 0))
            End If
        End If
    End If
    If dblTotalNumberOfFreeBytes > 0 Then
        If lngBytesPerSector > 0 Then
            txtTotalFreeSectors.Text = CStr(Round(dblTotalNumberOfFreeBytes / lngBytesPerSector, 0))
            
            If lngSectorsPerCluster > 0 Then
                txtTotalFreeClusters.Text = CStr(Round((dblTotalNumberOfFreeBytes / lngBytesPerSector) / lngSectorsPerCluster, 0))
            End If
        End If
    End If
    If dblTotalNumberOfBytes > 0 Then
        If lngBytesPerSector > 0 Then
            txtTotalSectors.Text = CStr(Round(dblTotalNumberOfBytes / lngBytesPerSector, 0))
            
            If lngSectorsPerCluster > 0 Then
                txtTotalClusters.Text = CStr(Round((dblTotalNumberOfBytes / lngBytesPerSector) / lngSectorsPerCluster, 0))
            End If
        End If
    End If
End Sub
