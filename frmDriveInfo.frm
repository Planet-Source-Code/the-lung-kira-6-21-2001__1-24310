VERSION 5.00
Begin VB.Form frmDriveInfo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Drive Info"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6735
   Icon            =   "frmDriveInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   6735
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtFileSystemName 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   1680
      Width           =   1575
   End
   Begin VB.TextBox txtMaximumComponentLength 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   1440
      Width           =   1575
   End
   Begin VB.TextBox txtVolumeSerialNumber 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   1200
      Width           =   1575
   End
   Begin VB.TextBox txtBytesPerSector 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   960
      Width           =   1575
   End
   Begin VB.TextBox txtSectorsPerCluster 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   720
      Width           =   1575
   End
   Begin VB.TextBox txtVolumeName 
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   2400
      TabIndex        =   15
      Top             =   2040
      Width           =   1575
   End
   Begin VB.CheckBox chkCaseIsPreserved 
      Enabled         =   0   'False
      Height          =   255
      Left            =   6360
      TabIndex        =   18
      Top             =   120
      Width           =   255
   End
   Begin VB.CheckBox chkCaseSensitive 
      Enabled         =   0   'False
      Height          =   255
      Left            =   6360
      TabIndex        =   20
      Top             =   360
      Width           =   255
   End
   Begin VB.CheckBox chkUnicodeStoredonDisk 
      Enabled         =   0   'False
      Height          =   255
      Left            =   6360
      TabIndex        =   22
      Top             =   600
      Width           =   255
   End
   Begin VB.CheckBox chkPersistantACLS 
      Enabled         =   0   'False
      Height          =   255
      Left            =   6360
      TabIndex        =   24
      Top             =   840
      Width           =   255
   End
   Begin VB.CheckBox chkFileCompression 
      Enabled         =   0   'False
      Height          =   255
      Left            =   6360
      TabIndex        =   26
      Top             =   1080
      Width           =   255
   End
   Begin VB.CheckBox chkVolumeisCompressed 
      Enabled         =   0   'False
      Height          =   255
      Left            =   6360
      TabIndex        =   28
      Top             =   1320
      Width           =   255
   End
   Begin VB.CheckBox chkNamedStreams 
      Enabled         =   0   'False
      Height          =   255
      Left            =   6360
      TabIndex        =   30
      Top             =   1560
      Width           =   255
   End
   Begin VB.CheckBox chkReadOnlyVolume 
      Enabled         =   0   'False
      Height          =   255
      Left            =   6360
      TabIndex        =   32
      Top             =   1800
      Width           =   255
   End
   Begin VB.CheckBox chkSupportsEncryption 
      Enabled         =   0   'False
      Height          =   255
      Left            =   6360
      TabIndex        =   34
      Top             =   2040
      Width           =   255
   End
   Begin VB.CheckBox chkSupportsObjectIDs 
      Enabled         =   0   'False
      Height          =   255
      Left            =   6360
      TabIndex        =   36
      Top             =   2280
      Width           =   255
   End
   Begin VB.CheckBox chkSupportsReparsePoints 
      Enabled         =   0   'False
      Height          =   255
      Left            =   6360
      TabIndex        =   38
      Top             =   2520
      Width           =   255
   End
   Begin VB.CheckBox chkSupportsSparseFiles 
      Enabled         =   0   'False
      Height          =   255
      Left            =   6360
      TabIndex        =   40
      Top             =   2760
      Width           =   255
   End
   Begin VB.CheckBox chkVolumeQuotas 
      Enabled         =   0   'False
      Height          =   255
      Left            =   6360
      TabIndex        =   42
      Top             =   3000
      Width           =   255
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   350
      Left            =   3000
      TabIndex        =   16
      Top             =   2400
      Width           =   975
   End
   Begin VB.TextBox txtDriveType 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   480
      Width           =   1575
   End
   Begin VB.ComboBox cboDrive 
      Height          =   315
      Left            =   2400
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label lblVolumeName 
      Caption         =   "Volume Name"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Label lblVolumeSerialNumber 
      Caption         =   "Volume Serial Number"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label lblFileSystemName 
      Caption         =   "File System Name"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Label lblMaximumComponentLength 
      Caption         =   "Maximum Component Length"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Label lblCaseIsPreserved 
      Caption         =   "Case Is Preserved"
      Height          =   255
      Left            =   4200
      TabIndex        =   17
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label lblCaseSensitive 
      Caption         =   "Case Sensitive"
      Height          =   255
      Left            =   4200
      TabIndex        =   19
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label lblUnicodeStoredonDisk 
      Caption         =   "Unicode Stored on Disk"
      Height          =   255
      Left            =   4200
      TabIndex        =   21
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label lblPersistantACLS 
      Caption         =   "Persistant ACLS"
      Height          =   255
      Left            =   4200
      TabIndex        =   23
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label lblFileCompression 
      Caption         =   "File Compression"
      Height          =   255
      Left            =   4200
      TabIndex        =   25
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label lblVolumeisCompressed 
      Caption         =   "Volume is Compressed"
      Height          =   255
      Left            =   4200
      TabIndex        =   27
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label lblNamedStreams 
      Caption         =   "Named Streams"
      Height          =   255
      Left            =   4200
      TabIndex        =   29
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Label lblReadOnlyVolume 
      Caption         =   "Read Only Volume"
      Height          =   255
      Left            =   4200
      TabIndex        =   31
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label lblSupportsEncryption 
      Caption         =   "Supports Encryption"
      Height          =   255
      Left            =   4200
      TabIndex        =   33
      Top             =   2040
      Width           =   1935
   End
   Begin VB.Label lblSupportsObjectIDs 
      Caption         =   "Supports Object IDs"
      Height          =   255
      Left            =   4200
      TabIndex        =   35
      Top             =   2280
      Width           =   1935
   End
   Begin VB.Label lblSupportsReparsePoints 
      Caption         =   "Supports Reparse Points"
      Height          =   255
      Left            =   4200
      TabIndex        =   37
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Label lblSupportsSparseFiles 
      Caption         =   "Supports Sparse Files"
      Height          =   255
      Left            =   4200
      TabIndex        =   39
      Top             =   2760
      Width           =   1935
   End
   Begin VB.Label lblVolumeQuotas 
      Caption         =   "Volume Quotas"
      Height          =   255
      Left            =   4200
      TabIndex        =   41
      Top             =   3000
      Width           =   1935
   End
   Begin VB.Label lblSectorsPerCluster 
      Caption         =   "Sectors Per Cluster"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label lblBytesPerSector 
      Caption         =   "Bytes Per Sector"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label lblDriveType 
      Caption         =   "Drive Type"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label lblDrive 
      Caption         =   "Drive"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "frmDriveInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboDrive_Click()
    Select Case GetDriveType(cboDrive.List(cboDrive.ListIndex))
        Case DRIVE_UNKNOWN: txtDriveType.Text = "Unknown"
        Case DRIVE_NO_ROOT_DIR: txtDriveType.Text = "No Root Directory"
        Case DRIVE_REMOVABLE: txtDriveType.Text = "Removable"
        Case DRIVE_FIXED: txtDriveType.Text = "Fixed"
        Case DRIVE_REMOTE: txtDriveType.Text = "Remote"
        Case DRIVE_CDROM: txtDriveType.Text = "CDROM"
        Case DRIVE_RAMDISK: txtDriveType.Text = "RAM Disk"
        Case Else: txtDriveType.Text = "Unknown"
    End Select
    
    
    Dim lngSectorsPerCluster As Long
    Dim lngBytesPerSector As Long
    Dim lngNumberOfFreeClusters As Long
    Dim lngTotalNumberOfClusters As Long
    
    If GetDiskFreeSpace(cboDrive.List(cboDrive.ListIndex), lngSectorsPerCluster, lngBytesPerSector, lngNumberOfFreeClusters, lngTotalNumberOfClusters) = False Then Failed "GetDiskFreeSpace"
    
    txtSectorsPerCluster.Text = CStr(lngSectorsPerCluster)
    txtBytesPerSector.Text = CStr(lngBytesPerSector)
    
    
    Dim strVolumeName As String
    Dim lngVolumeSerialNumber As Long
    Dim lngMaximumComponentLength As Long
    Dim lngFileSystemFlags As Long
    Dim strFileSystemName As String
    
    strVolumeName = String$(11, &H0)
    strFileSystemName = String$(256, &H0)
    
    If GetVolumeInformation(cboDrive.List(cboDrive.ListIndex), strVolumeName, Len(strVolumeName), lngVolumeSerialNumber, lngMaximumComponentLength, lngFileSystemFlags, strFileSystemName, Len(strFileSystemName)) = False Then Failed "GetVolumeInformation"
    
    txtVolumeName.Text = strVolumeName
    txtVolumeSerialNumber.Text = Right$("00000000" & ltoa_(lngVolumeSerialNumber, 16), 8)
    txtMaximumComponentLength.Text = CStr(lngMaximumComponentLength)
    txtFileSystemName.Text = strFileSystemName
    
    If lngFileSystemFlags And FS_CASE_IS_PRESERVED Then chkCaseIsPreserved.value = 1 Else chkCaseIsPreserved.value = 0
    If lngFileSystemFlags And FS_CASE_SENSITIVE Then chkCaseSensitive.value = 1 Else: chkCaseSensitive.value = 0
    If lngFileSystemFlags And FS_UNICODE_STORED_ON_DISK Then chkUnicodeStoredonDisk.value = 1 Else: chkUnicodeStoredonDisk.value = 0
    If lngFileSystemFlags And FS_PERSISTENT_ACLS Then chkPersistantACLS.value = 1 Else chkPersistantACLS.value = 0
    If lngFileSystemFlags And FS_FILE_COMPRESSION Then chkFileCompression.value = 1 Else chkFileCompression.value = 0
    If lngFileSystemFlags And FS_VOL_IS_COMPRESSED Then chkVolumeisCompressed.value = 1 Else chkVolumeisCompressed.value = 0
    If lngFileSystemFlags And FILE_NAMED_STREAMS Then chkNamedStreams.value = 1 Else chkNamedStreams.value = 0
    If lngFileSystemFlags And FILE_SUPPORTS_ENCRYPTION Then chkSupportsEncryption.value = 1 Else chkSupportsEncryption.value = 0
    If lngFileSystemFlags And FILE_SUPPORTS_OBJECT_IDS Then chkSupportsObjectIDs.value = 1 Else chkSupportsObjectIDs.value = 0
    If lngFileSystemFlags And FILE_SUPPORTS_REPARSE_POINTS Then chkSupportsReparsePoints.value = 1 Else chkSupportsReparsePoints.value = 0
    If lngFileSystemFlags And FILE_SUPPORTS_SPARSE_FILES Then chkSupportsSparseFiles.value = 1 Else chkSupportsSparseFiles.value = 0
    If lngFileSystemFlags And FILE_VOLUME_QUOTAS Then chkVolumeQuotas.value = 1 Else: chkVolumeQuotas.value = 0
End Sub

Private Sub cmdApply_Click()
    If SetVolumeLabel(cboDrive.List(cboDrive.ListIndex), txtVolumeName.Text) = False Then Failed "SetVolumeLabel"
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
End Sub

Private Sub txtVolumeName_Change()
    txtVolumeName.Text = Rem_NonFat_Chr(txtVolumeName.Text)
    If Len(txtVolumeName.Text) > 11 Then txtVolumeName.Text = Left$(txtVolumeName.Text, 11)
End Sub
