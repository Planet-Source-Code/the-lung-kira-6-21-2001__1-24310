VERSION 5.00
Begin VB.Form frmSoundSentry 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sound Sentry"
   ClientHeight    =   4755
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3615
   Icon            =   "frmSoundSentry.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   3615
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtWindowsEffectDLL 
      Height          =   285
      Left            =   2280
      TabIndex        =   23
      Top             =   3960
      Width           =   1215
   End
   Begin VB.TextBox txtWindowsEffectDuration 
      Height          =   285
      Left            =   2280
      TabIndex        =   21
      Top             =   3600
      Width           =   1215
   End
   Begin VB.ComboBox cboWindowsEffect 
      Height          =   315
      Left            =   2280
      TabIndex        =   19
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox txtGraphicEffectRGB 
      Height          =   285
      Left            =   2280
      TabIndex        =   11
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox txtGraphicEffectDuration 
      Height          =   285
      Left            =   2280
      TabIndex        =   9
      Top             =   1440
      Width           =   1215
   End
   Begin VB.ComboBox cboGraphicEffect 
      Height          =   315
      Left            =   2280
      TabIndex        =   7
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox txtTextEffectRGB 
      Height          =   285
      Left            =   2280
      TabIndex        =   17
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox txtTextEffectDuration 
      Height          =   285
      Left            =   2280
      TabIndex        =   15
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CheckBox chkSoundSentryOn 
      Height          =   255
      Left            =   3240
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   720
      Width           =   255
   End
   Begin VB.CheckBox chkIndicator 
      Height          =   255
      Left            =   3240
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   480
      Width           =   255
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   350
      Left            =   2520
      TabIndex        =   24
      Top             =   4320
      Width           =   975
   End
   Begin VB.CheckBox chkAvailable 
      Enabled         =   0   'False
      Height          =   255
      Left            =   3240
      TabIndex        =   1
      Top             =   120
      Width           =   255
   End
   Begin VB.ComboBox cboTextEffect 
      Height          =   315
      Left            =   2280
      TabIndex        =   13
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label lblWindowsEffectDLL 
      Caption         =   "Windows Effect DLL"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   3960
      Width           =   1935
   End
   Begin VB.Label lblWindowsEffectDuration 
      Caption         =   "Windows Effect Duration"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   3600
      Width           =   1935
   End
   Begin VB.Label lblWindowsEffect 
      Caption         =   "Windows Effect"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   3240
      Width           =   1935
   End
   Begin VB.Label lblGraphicEffectRGB 
      Caption         =   "Graphic Effect RGB"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label lblGraphicEffectDuration 
      Caption         =   "Graphic Effect Duration"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label lblGraphicEffect 
      Caption         =   "Graphic Effect"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label lblTextEffectRGB 
      Caption         =   "Text Effect RGB"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Label lblTextEffectDuration 
      Caption         =   "Text Effect Duration"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Label lblSoundSentryOn 
      Caption         =   "Sound Sentry On"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label lblIndicator 
      Caption         =   "Indicator"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label lblAvailable 
      Caption         =   "Available"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label lblTextEffect 
      Caption         =   "Text Effect"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   2160
      Width           =   1935
   End
End
Attribute VB_Name = "frmSoundSentry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdApply_Click()
    Dim SOUNDSENTRY As SOUNDSENTRY
    SOUNDSENTRY.cbSize = Len(SOUNDSENTRY)
    
    Dim Indicator As Long
    Dim SoundSentryOn As Long
    
    If chkIndicator.value = 1 Then Indicator = SERKF_INDICATOR
    If chkSoundSentryOn.value = 1 Then SoundSentryOn = SSF_SOUNDSENTRYON
    
    
    With SOUNDSENTRY
        .dwFlags = Indicator Or SoundSentryOn
        
        Select Case cboGraphicEffect.ListIndex
            Case 0: .iFSGrafEffect = SSGF_DISPLAY
            Case 1: .iFSGrafEffect = SSGF_NONE
        End Select
        .iFSGrafEffectColor = Val(txtGraphicEffectRGB.Text)
        .iFSGrafEffectMSec = Val(txtGraphicEffectDuration.Text)
        
        Select Case cboTextEffect.ListIndex
            Case 0: .iFSTextEffect = SSTF_BORDER
            Case 1: .iFSTextEffect = SSTF_CHARS
            Case 2: .iFSTextEffect = SSTF_DISPLAY
            Case 3: .iFSTextEffect = SSTF_NONE
        End Select
        .iFSTextEffectColorBits = Val(txtTextEffectRGB.Text)
        .iFSTextEffectMSec = Val(txtTextEffectDuration.Text)
        
        Select Case cboWindowsEffect.ListIndex
            Case 0: .iWindowsEffect = SSWF_CUSTOM
            Case 1: .iWindowsEffect = SSWF_DISPLAY
            Case 2: .iWindowsEffect = SSWF_NONE
            Case 3: .iWindowsEffect = SSWF_TITLE
            Case 4: .iWindowsEffect = SSWF_WINDOW
        End Select
        .iWindowsEffectMSec = Val(txtWindowsEffectDuration.Text)
        .lpszWindowsEffectDLL = txtWindowsEffectDLL.Text
    End With
    
    If SystemParametersInfo(SPI_SETSOUNDSENTRY, SOUNDSENTRY.cbSize, SOUNDSENTRY, SPIF_UPDATEINIFILE) = False Then Failed "SystemParametersInfo"
End Sub

Private Sub Form_Load()
    With cboGraphicEffect
        .AddItem "Display"
        .AddItem "None"
    End With
    With cboTextEffect
        .AddItem "Flash Border"
        .AddItem "Flash Characters"
        .AddItem "Flash Display"
        .AddItem "None"
    End With
    With cboWindowsEffect
        .AddItem "Custom"
        .AddItem "Flash Display"
        .AddItem "None"
        .AddItem "Flash Title Bar"
        .AddItem "Flash Window"
    End With
    
    
    Dim SOUNDSENTRY As SOUNDSENTRY
    SOUNDSENTRY.cbSize = Len(SOUNDSENTRY)
    
    If SystemParametersInfo(SPI_GETSOUNDSENTRY, SOUNDSENTRY.cbSize, SOUNDSENTRY, 0) = False Then Failed "SystemParametersInfo"
    
    If SOUNDSENTRY.dwFlags And SSF_AVAILABLE Then
        With SOUNDSENTRY
            If .dwFlags And SSF_AVAILABLE Then chkAvailable.value = 1
            If .dwFlags And SSF_INDICATOR Then chkIndicator.value = 1
            If .dwFlags And SSF_SOUNDSENTRYON Then chkSoundSentryOn.value = 1
            
            
            Select Case .iWindowsEffect
                Case SSWF_CUSTOM: cboWindowsEffect.ListIndex = 0
                Case SSWF_DISPLAY: cboWindowsEffect.ListIndex = 1
                Case SSWF_NONE: cboWindowsEffect.ListIndex = 2
                Case SSWF_TITLE: cboWindowsEffect.ListIndex = 3
                Case SSWF_WINDOW: cboWindowsEffect.ListIndex = 4
            End Select
            txtWindowsEffectDLL.Text = .lpszWindowsEffectDLL
            
            
            If WinVersion(0, -1, True) = True Then
                Select Case .iFSGrafEffect
                    Case SSGF_DISPLAY: cboGraphicEffect.ListIndex = 0
                    Case SSGF_NONE: cboGraphicEffect.ListIndex = 0
                End Select
                txtGraphicEffectDuration.Text = CStr(.iFSGrafEffectMSec)
                txtGraphicEffectRGB.Text = CStr(.iFSGrafEffectColor)
                
                Select Case .iFSTextEffect
                    Case SSTF_BORDER: cboTextEffect.ListIndex = 0
                    Case SSTF_CHARS: cboTextEffect.ListIndex = 1
                    Case SSTF_DISPLAY: cboTextEffect.ListIndex = 2
                    Case SSTF_NONE: cboTextEffect.ListIndex = 3
                End Select
                txtTextEffectDuration.Text = CStr(.iFSTextEffectMSec)
                txtTextEffectRGB.Text = CStr(.iFSTextEffectColorBits)
                
                txtWindowsEffectDuration.Text = CStr(.iWindowsEffectMSec)
            Else
                lblGraphicEffect.Enabled = False
                cboGraphicEffect.Enabled = False
                lblGraphicEffectDuration.Enabled = False
                txtGraphicEffectDuration.Enabled = False
                lblGraphicEffectRGB.Enabled = False
                txtGraphicEffectRGB.Enabled = False
                lblTextEffect.Enabled = False
                cboTextEffect.Enabled = False
                lblTextEffectDuration.Enabled = False
                txtTextEffectDuration.Enabled = False
                lblTextEffectRGB.Enabled = False
                txtTextEffectRGB.Enabled = False
                lblWindowsEffectDuration.Enabled = False
                txtWindowsEffectDuration.Enabled = False
            End If
        End With
    Else
        lblIndicator.Enabled = False
        chkIndicator.Enabled = False
        lblSoundSentryOn.Enabled = False
        chkSoundSentryOn.Enabled = False
        lblGraphicEffect.Enabled = False
        cboGraphicEffect.Enabled = False
        lblGraphicEffectDuration.Enabled = False
        txtGraphicEffectDuration.Enabled = False
        lblGraphicEffectRGB.Enabled = False
        txtGraphicEffectRGB.Enabled = False
        lblTextEffect.Enabled = False
        cboTextEffect.Enabled = False
        lblTextEffectDuration.Enabled = False
        txtTextEffectDuration.Enabled = False
        lblTextEffectRGB.Enabled = False
        txtTextEffectRGB.Enabled = False
        lblWindowsEffect.Enabled = False
        cboWindowsEffect.Enabled = False
        lblWindowsEffectDLL.Enabled = False
        txtWindowsEffectDLL.Enabled = False
        lblWindowsEffectDuration.Enabled = False
        txtWindowsEffectDuration.Enabled = False
        
        cmdApply.Enabled = False
    End If
End Sub

Private Sub txtGraphicEffectDuration_Change()
    txtGraphicEffectDuration.Text = CStr(Val(Rem_NonNumeric_Chr(txtGraphicEffectDuration.Text)))
    If Val(txtGraphicEffectDuration.Text) < 0 Then txtGraphicEffectDuration.Text = "0"
    If Val(txtGraphicEffectDuration.Text) > 2147483647 Then txtGraphicEffectDuration.Text = "2147483647"
End Sub

Private Sub txtGraphicEffectRGB_Change()
    txtGraphicEffectRGB.Text = CStr(Val(Rem_NonNumeric_Chr(txtGraphicEffectRGB.Text)))
    If Val(txtGraphicEffectRGB.Text) < -2147483648# Then txtGraphicEffectRGB.Text = "-2147483648"
    If Val(txtGraphicEffectRGB.Text) > 2147483647 Then txtGraphicEffectRGB.Text = "2147483647"
End Sub

Private Sub txtTextEffectDuration_Change()
    txtTextEffectDuration.Text = CStr(Val(Rem_NonNumeric_Chr(txtTextEffectDuration.Text)))
    If Val(txtTextEffectDuration.Text) < 0 Then txtTextEffectDuration.Text = "0"
    If Val(txtTextEffectDuration.Text) > 2147483647 Then txtTextEffectDuration.Text = "2147483647"
End Sub

Private Sub txtTextEffectRGB_Change()
    txtTextEffectRGB.Text = CStr(Val(Rem_NonNumeric_Chr(txtTextEffectRGB.Text)))
    If Val(txtTextEffectRGB.Text) < -2147483648# Then txtTextEffectRGB.Text = "-2147483648"
    If Val(txtTextEffectRGB.Text) > 2147483647 Then txtTextEffectRGB.Text = "2147483647"
End Sub

Private Sub txtWindowsEffectDLL_Change()
    If Len(txtWindowsEffectDLL.Text) > MAX_PATH Then txtWindowsEffectDLL.Text = Left$(txtWindowsEffectDLL.Text, MAX_PATH)
End Sub

Private Sub txtWindowsEffectDuration_Change()
    txtWindowsEffectDuration.Text = CStr(Val(Rem_NonNumeric_Chr(txtWindowsEffectDuration.Text)))
    If Val(txtWindowsEffectDuration.Text) < 0 Then txtWindowsEffectDuration.Text = "0"
    If Val(txtWindowsEffectDuration.Text) > 2147483647 Then txtWindowsEffectDuration.Text = "2147483647"
End Sub
