VERSION 5.00
Begin VB.Form frmMZ 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MZ"
   ClientHeight    =   4860
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4815
   Icon            =   "frmMZ.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4860
   ScaleWidth      =   4815
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtSelected 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   360
      Width           =   4575
   End
   Begin VB.TextBox txtOverlay 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   29
      Top             =   4080
      Width           =   1095
   End
   Begin VB.TextBox txtRelocationOffset 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   27
      Top             =   3840
      Width           =   1095
   End
   Begin VB.TextBox txtInitialCS 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   25
      Top             =   3600
      Width           =   1095
   End
   Begin VB.TextBox txtInitialIP 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   3360
      Width           =   1095
   End
   Begin VB.TextBox txtChecksum 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   3120
      Width           =   1095
   End
   Begin VB.TextBox txtInitialSP 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   2880
      Width           =   1095
   End
   Begin VB.TextBox txtInitialSS 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   2640
      Width           =   1095
   End
   Begin VB.TextBox txtMaxPara 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   2400
      Width           =   1095
   End
   Begin VB.TextBox txtMinPara 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   2160
      Width           =   1095
   End
   Begin VB.TextBox txt16Para 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   1920
      Width           =   1095
   End
   Begin VB.TextBox txtRelocationTables 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   1680
      Width           =   1095
   End
   Begin VB.TextBox txt512Pages 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1440
      Width           =   1095
   End
   Begin VB.TextBox txtSizeMod 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CheckBox chkSignature 
      Enabled         =   0   'False
      Height          =   255
      Left            =   4440
      TabIndex        =   3
      Top             =   840
      Width           =   255
   End
   Begin VB.CommandButton cmdChoose 
      Caption         =   "Choose"
      Height          =   350
      Left            =   3720
      TabIndex        =   30
      Top             =   4440
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
   Begin VB.Label lblOverlay 
      Caption         =   "Overlay Number"
      Height          =   255
      Left            =   120
      TabIndex        =   28
      Top             =   4080
      Width           =   3255
   End
   Begin VB.Label lblRelocationOffset 
      Caption         =   "Relocation Offset"
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   3840
      Width           =   3255
   End
   Begin VB.Label lblInitialIP 
      Caption         =   "Initial IP Value"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   3360
      Width           =   3255
   End
   Begin VB.Label lblInitialCS 
      Caption         =   "Initial Relative CS Value (Paragraphs)"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   3600
      Width           =   3255
   End
   Begin VB.Label lblChecksum 
      Caption         =   "Checksum"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   3120
      Width           =   3255
   End
   Begin VB.Label lblInitialSP 
      Caption         =   "Initial SP Value"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   2880
      Width           =   3255
   End
   Begin VB.Label lblInitialSS 
      Caption         =   "Initial Relative SS Value (Paragraphs)"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   2640
      Width           =   3255
   End
   Begin VB.Label lblMinPara 
      Caption         =   "Minimum Number of Paragraphs"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   2160
      Width           =   3255
   End
   Begin VB.Label lblMaxPara 
      Caption         =   "Maximum Number of Paragraphs"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   2400
      Width           =   3255
   End
   Begin VB.Label lbl16Para 
      Caption         =   "16b Paragraphs for Header/Relocation Table"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1920
      Width           =   3255
   End
   Begin VB.Label lblRelocationTables 
      Caption         =   "Relocation Tables"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1680
      Width           =   3255
   End
   Begin VB.Label lbl512Pages 
      Caption         =   "512b Pages"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   3255
   End
   Begin VB.Label lblSizeMod 
      Caption         =   "Image Size Mod 512"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   3255
   End
   Begin VB.Label lblSignature 
      Caption         =   "Signature"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   3255
   End
End
Attribute VB_Name = "frmMZ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdChoose_Click()
    chkSignature.value = 0
    txtSizeMod.Text = ""
    txt512Pages.Text = ""
    txtRelocationTables.Text = ""
    txt16Para.Text = ""
    txtMinPara.Text = ""
    txtMaxPara.Text = ""
    txtInitialSS.Text = ""
    txtInitialSP.Text = ""
    txtChecksum.Text = ""
    txtInitialIP.Text = ""
    txtInitialCS.Text = ""
    txtRelocationOffset.Text = ""
    txtOverlay.Text = ""
    
    Dim strFileName As String
    strFileName = GetOpenName(frmMZ.hwnd, "All Files (*.*)" & Chr$(0) & "*.*" & Chr$(0), 2, "Open", OFN_EXPLORER Or OFN_FILEMUSTEXIST Or OFN_HIDEREADONLY Or OFN_DONTADDTORECENT)
    txtSelected.Text = strFileName
    If strFileName = "" Then Exit Sub
    
    If Not FileSize_Name(strFileName) > 0 Then
        If MessageBoxEx(&H0, "File size is 0.", "Error", MB_OK Or MB_ICONWARNING Or MB_SETFOREGROUND, 0) = 0 Then Failed "MessageBoxEx"
        Exit Sub
    End If
    
    
    Dim strFileContents As String
    
    strFileContents = ReadFile_String(strFileName, 28, 0)
    If Len(strFileContents) < 28 Then
        strFileContents = strFileContents & String$(28 - Len(strFileContents), &H0)
    End If
    
    
    If Mid$(strFileContents, 1, 2) <> "MZ" Then
        If Mid$(strFileContents, 1, 2) <> "ZM" Then
            Exit Sub
        End If
    End If
    
    chkSignature.value = 1
    
    txtSizeMod.Text = CStr(strtoul_(Right$("00" & ltoa_(Asc(Mid$(strFileContents, 4, 1)), 16), 2) & _
                            Right$("00" & ltoa_(Asc(Mid$(strFileContents, 3, 1)), 16), 2), 16))
    txt512Pages.Text = CStr(strtoul_(Right$("00" & ltoa_(Asc(Mid$(strFileContents, 6, 1)), 16), 2) & _
                            Right$("00" & ltoa_(Asc(Mid$(strFileContents, 5, 1)), 16), 2), 16))
    txtRelocationTables.Text = CStr(strtoul_(Right$("00" & ltoa_(Asc(Mid$(strFileContents, 8, 1)), 16), 2) & _
                            Right$("00" & ltoa_(Asc(Mid$(strFileContents, 7, 1)), 16), 2), 16))
    txt16Para.Text = CStr(strtoul_(Right$("00" & ltoa_(Asc(Mid$(strFileContents, 10, 1)), 16), 2) & _
                            Right$("00" & ltoa_(Asc(Mid$(strFileContents, 9, 1)), 16), 2), 16))
    txtMinPara.Text = CStr(strtoul_(Right$("00" & ltoa_(Asc(Mid$(strFileContents, 12, 1)), 16), 2) & _
                            Right$("00" & ltoa_(Asc(Mid$(strFileContents, 11, 1)), 16), 2), 16))
    txtMaxPara.Text = CStr(strtoul_(Right$("00" & ltoa_(Asc(Mid$(strFileContents, 14, 1)), 16), 2) & _
                            Right$("00" & ltoa_(Asc(Mid$(strFileContents, 13, 1)), 16), 2), 16))
    txtInitialSS.Text = CStr(strtoul_(Right$("00" & ltoa_(Asc(Mid$(strFileContents, 16, 1)), 16), 2) & _
                            Right$("00" & ltoa_(Asc(Mid$(strFileContents, 15, 1)), 16), 2), 16))
    txtInitialSP.Text = CStr(strtoul_(Right$("00" & ltoa_(Asc(Mid$(strFileContents, 18, 1)), 16), 2) & _
                            Right$("00" & ltoa_(Asc(Mid$(strFileContents, 17, 1)), 16), 2), 16))
    txtChecksum.Text = CStr(strtoul_(Right$("00" & ltoa_(Asc(Mid$(strFileContents, 20, 1)), 16), 2) & _
                            Right$("00" & ltoa_(Asc(Mid$(strFileContents, 19, 1)), 16), 2), 16))
    txtInitialIP.Text = CStr(strtoul_(Right$("00" & ltoa_(Asc(Mid$(strFileContents, 22, 1)), 16), 2) & _
                            Right$("00" & ltoa_(Asc(Mid$(strFileContents, 21, 1)), 16), 2), 16))
    txtInitialCS.Text = CStr(strtoul_(Right$("00" & ltoa_(Asc(Mid$(strFileContents, 24, 1)), 16), 2) & _
                            Right$("00" & ltoa_(Asc(Mid$(strFileContents, 23, 1)), 16), 2), 16))
    txtRelocationOffset.Text = CStr(strtoul_(Right$("00" & ltoa_(Asc(Mid$(strFileContents, 26, 1)), 16), 2) & _
                            Right$("00" & ltoa_(Asc(Mid$(strFileContents, 25, 1)), 16), 2), 16))
    txtOverlay.Text = CStr(strtoul_(Right$("00" & ltoa_(Asc(Mid$(strFileContents, 28, 1)), 16), 2) & _
                            Right$("00" & ltoa_(Asc(Mid$(strFileContents, 27, 1)), 16), 2), 16))
End Sub

