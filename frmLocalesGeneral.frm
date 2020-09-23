VERSION 5.00
Begin VB.Form frmLocalesGeneral 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Locales - General"
   ClientHeight    =   4080
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8895
   Icon            =   "frmLocalesGeneral.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   8895
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtDefaultMacCodePage 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   24
      Top             =   3720
      Width           =   1935
   End
   Begin VB.TextBox txtDefaultLanguage 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   3480
      Width           =   1935
   End
   Begin VB.TextBox txtDefaultEBCDICCodePage 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   3240
      Width           =   1935
   End
   Begin VB.TextBox txtDefaultCountry 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   3000
      Width           =   1935
   End
   Begin VB.TextBox txtDefaultCodePage 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   2760
      Width           =   1935
   End
   Begin VB.TextBox txtDefaultANSICodePage 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   2520
      Width           =   1935
   End
   Begin VB.TextBox txtCountry 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   2280
      Width           =   1935
   End
   Begin VB.TextBox txtAbbreviatedLanguageName 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   2040
      Width           =   1935
   End
   Begin VB.TextBox txtAbbrevISOCountryName 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   1800
      Width           =   1935
   End
   Begin VB.TextBox txtNativeLanguageName 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6840
      Locked          =   -1  'True
      TabIndex        =   44
      Top             =   2520
      Width           =   1935
   End
   Begin VB.TextBox txtLanguageID 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6840
      Locked          =   -1  'True
      TabIndex        =   38
      Top             =   1560
      Width           =   1935
   End
   Begin VB.TextBox txtFullLanguageName 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6840
      Locked          =   -1  'True
      TabIndex        =   36
      Top             =   1320
      Width           =   1935
   End
   Begin VB.TextBox txtFullISOCountryName 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6840
      Locked          =   -1  'True
      TabIndex        =   34
      Top             =   1080
      Width           =   1935
   End
   Begin VB.TextBox txtFullEnglishLanguageName 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6840
      Locked          =   -1  'True
      TabIndex        =   32
      Top             =   840
      Width           =   1935
   End
   Begin VB.TextBox txtFullEnglishCountryName 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6840
      Locked          =   -1  'True
      TabIndex        =   30
      Top             =   600
      Width           =   1935
   End
   Begin VB.TextBox txtFullCountryName 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6840
      Locked          =   -1  'True
      TabIndex        =   28
      Top             =   360
      Width           =   1935
   End
   Begin VB.TextBox txtFontSignature 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6840
      Locked          =   -1  'True
      TabIndex        =   26
      Top             =   120
      Width           =   1935
   End
   Begin VB.TextBox txtSortName 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6840
      Locked          =   -1  'True
      TabIndex        =   48
      Top             =   3240
      Width           =   1935
   End
   Begin VB.TextBox txtNativeCountryName 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6840
      Locked          =   -1  'True
      TabIndex        =   42
      Top             =   2280
      Width           =   1935
   End
   Begin VB.TextBox txtAbbreviatedCountryName 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   1560
      Width           =   1935
   End
   Begin VB.ComboBox cboPaperSize 
      Height          =   315
      Left            =   6840
      TabIndex        =   46
      Top             =   2880
      Width           =   1935
   End
   Begin VB.ComboBox cboMeasurement 
      Height          =   315
      Left            =   6840
      TabIndex        =   40
      Top             =   1920
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
      Left            =   2280
      TabIndex        =   3
      Top             =   360
      Width           =   1935
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   350
      Left            =   3240
      TabIndex        =   4
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   350
      Left            =   7800
      TabIndex        =   49
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label lblSortName 
      Caption         =   "Sort Name"
      Height          =   255
      Left            =   4560
      TabIndex        =   47
      Top             =   3240
      Width           =   2055
   End
   Begin VB.Label lblNativeLanguageName 
      Caption         =   "Native Language Name"
      Height          =   255
      Left            =   4560
      TabIndex        =   43
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Label lblNativeCountryName 
      Caption         =   "Native Country Name"
      Height          =   255
      Left            =   4560
      TabIndex        =   41
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Label lblFullLanguageName 
      Caption         =   "Full Language Name"
      Height          =   255
      Left            =   4560
      TabIndex        =   35
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label lblAbbrevISOCountryName 
      Caption         =   "Abbrev ISO Country Name"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Label lblFullISOCountryName 
      Caption         =   "Full ISO Country Name"
      Height          =   255
      Left            =   4560
      TabIndex        =   33
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label lblFullEnglishLanguageName 
      Caption         =   "Full English Language Name"
      Height          =   255
      Left            =   4560
      TabIndex        =   31
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label lblFullEnglishCountryName 
      Caption         =   "Full English Country Name"
      Height          =   255
      Left            =   4560
      TabIndex        =   29
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label lblFullCountryName 
      Caption         =   "Full Country Name"
      Height          =   255
      Left            =   4560
      TabIndex        =   27
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label lblAbbreviatedLanguageName 
      Caption         =   "Abbrev Language Name"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Label lblAbbreviatedCountryName 
      Caption         =   "Abbrev Country Name"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label lblPaperSize 
      Caption         =   "Paper Size"
      Height          =   255
      Left            =   4560
      TabIndex        =   45
      Top             =   2880
      Width           =   2055
   End
   Begin VB.Label lblMeasurement 
      Caption         =   "Measurement"
      Height          =   255
      Left            =   4560
      TabIndex        =   39
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label lblLanguageID 
      Caption         =   "Language ID"
      Height          =   255
      Left            =   4560
      TabIndex        =   37
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label lblDefaultMacCodePage 
      Caption         =   "Default Mac Code Page"
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   3720
      Width           =   2055
   End
   Begin VB.Label lblDefaultLanguage 
      Caption         =   "Default Language"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   3480
      Width           =   2055
   End
   Begin VB.Label lblDefaultEBCDICCodePage 
      Caption         =   "Default EBCDIC Code Page"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   3240
      Width           =   2055
   End
   Begin VB.Label lblDefaultCountry 
      Caption         =   "Default Country"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   3000
      Width           =   2055
   End
   Begin VB.Label lblDefaultCodePage 
      Caption         =   "Default Code Page"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   2760
      Width           =   2055
   End
   Begin VB.Label lblDefaultANSICodePage 
      Caption         =   "Default ANSI Code Page"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Label lblCountry 
      Caption         =   "Country"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Label lblFontSignature 
      Caption         =   "Font Signature"
      Height          =   255
      Left            =   4560
      TabIndex        =   25
      Top             =   120
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
      Left            =   2280
      TabIndex        =   2
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "frmLocalesGeneral"
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
    
    
    If cboMeasurement.ListIndex > -1 Then
        If SetLocaleInfo(lngLocale, LOCALE_IMEASURE, CStr(cboMeasurement.ListIndex)) = False Then Failed "SetLocaleInfo"
    End If
    If WinVersion(-1, 5000000, True) = True Then
        If cboPaperSize.ListIndex > -1 Then
            Dim strValue As String
            
            Select Case cboPaperSize.ListIndex
                Case 0: strValue = "1"
                Case 1: strValue = "5"
                Case 2: strValue = "8"
                Case 3: strValue = "9"
            End Select
            
            If SetLocaleInfo(lngLocale, LOCALE_IPAPERSIZE, strValue) = False Then Failed "SetLocaleInfo"
        End If
    End If
End Sub

Private Sub cmdRefresh_Click()
    LocaleListNum = 0
    Erase LocaleList
    lstLocales.Clear
    
    txtFontSignature.Text = ""
    txtCountry.Text = ""
    txtDefaultANSICodePage.Text = ""
    txtDefaultCodePage.Text = ""
    txtDefaultCountry.Text = ""
    txtDefaultEBCDICCodePage.Text = ""
    txtDefaultLanguage.Text = ""
    txtDefaultMacCodePage.Text = ""
    txtLanguageID.Text = ""
    cboMeasurement.ListIndex = -1
    cboPaperSize.ListIndex = -1
    txtAbbreviatedCountryName.Text = ""
    txtAbbreviatedLanguageName.Text = ""
    txtFullCountryName.Text = ""
    txtFullEnglishCountryName.Text = ""
    txtFullEnglishLanguageName.Text = ""
    txtFullISOCountryName.Text = ""
    txtAbbrevISOCountryName.Text = ""
    txtFullLanguageName.Text = ""
    txtNativeCountryName.Text = ""
    txtNativeLanguageName.Text = ""
    txtSortName.Text = ""
    
    
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
    With cboMeasurement
        .AddItem "Metric"
        .AddItem "U.S."
    End With
    With cboPaperSize
        .AddItem "US Letter"
        .AddItem "US legal"
        .AddItem "A3"
        .AddItem "A4"
    End With
    
    
    If WinVersion(0, -1, True) = False Then
        lblFontSignature.Enabled = False
        lblFullISOCountryName.Enabled = False
        lblAbbrevISOCountryName.Enabled = False
    End If
    If WinVersion(-1, 5000000, True) = False Then
        lblDefaultEBCDICCodePage.Enabled = False
        lblPaperSize.Enabled = False
        lblSortName.Enabled = False
    End If
    
    cmdRefresh_Click
End Sub

Private Sub lstLocales_Click()
    Dim lngLocale As Long
    lngLocale = strtoul_(lstLocales.List(lstLocales.ListIndex), 16)
    
    
    If WinID = VER_PLATFORM_WIN32_NT Then
        txtFontSignature.Text = Get_LocaleInfo(lngLocale, LOCALE_FONTSIGNATURE)
    End If
    
    txtCountry.Text = Get_LocaleInfo(lngLocale, LOCALE_ICOUNTRY)
    txtDefaultANSICodePage.Text = Get_LocaleInfo(lngLocale, LOCALE_IDEFAULTANSICODEPAGE)
    txtDefaultCodePage.Text = Get_LocaleInfo(lngLocale, LOCALE_IDEFAULTCODEPAGE)
    txtDefaultCountry.Text = Get_LocaleInfo(lngLocale, LOCALE_IDEFAULTCOUNTRY)
    txtDefaultEBCDICCodePage.Text = Get_LocaleInfo(lngLocale, LOCALE_IDEFAULTEBCDICCODEPAGE)
    txtDefaultLanguage.Text = Get_LocaleInfo(lngLocale, LOCALE_IDEFAULTLANGUAGE)
    txtDefaultMacCodePage.Text = Get_LocaleInfo(lngLocale, LOCALE_IDEFAULTMACCODEPAGE)
    txtLanguageID.Text = LangIdent(strtoul_(Get_LocaleInfo(lngLocale, LOCALE_ILANGUAGE), 16))
    cboMeasurement.ListIndex = Val(Get_LocaleInfo(lngLocale, LOCALE_IMEASURE))
    
    If WinVersion(-1, 5000000, True) = True Then
        Select Case Val(Get_LocaleInfo(lngLocale, LOCALE_IPAPERSIZE))
            Case 1: cboPaperSize.ListIndex = 0
            Case 5: cboPaperSize.ListIndex = 1
            Case 8: cboPaperSize.ListIndex = 2
            Case 9: cboPaperSize.ListIndex = 3
            Case Else: cboPaperSize.ListIndex = -1
        End Select
    End If
    
    txtAbbreviatedCountryName.Text = Get_LocaleInfo(lngLocale, LOCALE_SABBREVCTRYNAME)
    txtAbbreviatedLanguageName.Text = Get_LocaleInfo(lngLocale, LOCALE_SABBREVLANGNAME)
    txtFullCountryName.Text = Get_LocaleInfo(lngLocale, LOCALE_SCOUNTRY)
    txtFullEnglishCountryName.Text = Get_LocaleInfo(lngLocale, LOCALE_SENGCOUNTRY)
    txtFullEnglishLanguageName.Text = Get_LocaleInfo(lngLocale, LOCALE_SENGLANGUAGE)
    
    If WinID = VER_PLATFORM_WIN32_NT Then
        txtFullISOCountryName.Text = Get_LocaleInfo(lngLocale, LOCALE_SISO3166CTRYNAME)
        txtAbbrevISOCountryName.Text = Get_LocaleInfo(lngLocale, LOCALE_SISO639LANGNAME)
    End If
    
    txtFullLanguageName.Text = Get_LocaleInfo(lngLocale, LOCALE_SLANGUAGE)
    txtNativeCountryName.Text = Get_LocaleInfo(lngLocale, LOCALE_SNATIVECTRYNAME)
    txtNativeLanguageName.Text = Get_LocaleInfo(lngLocale, LOCALE_SNATIVELANGNAME)
    
    If WinVersion(-1, 5000000, True) = True Then
        txtSortName.Text = Get_LocaleInfo(lngLocale, LOCALE_SSORTNAME)
    End If
End Sub
