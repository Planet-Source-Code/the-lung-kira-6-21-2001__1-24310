VERSION 5.00
Begin VB.Form frmLocalesCurrency 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Locales - Currency"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8895
   Icon            =   "frmLocalesCurrency.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   8895
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtPositivePosition 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6840
      Locked          =   -1  'True
      TabIndex        =   38
      Top             =   2760
      Width           =   1935
   End
   Begin VB.TextBox txtPositivePositionFormat 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6840
      Locked          =   -1  'True
      TabIndex        =   36
      Top             =   2520
      Width           =   1935
   End
   Begin VB.CheckBox chkPositiveSpaceSeperation 
      Enabled         =   0   'False
      Height          =   255
      Left            =   8520
      TabIndex        =   34
      Top             =   2280
      Width           =   255
   End
   Begin VB.TextBox txtNegativePosition 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6840
      Locked          =   -1  'True
      TabIndex        =   32
      Top             =   2040
      Width           =   1935
   End
   Begin VB.TextBox txtNegativePositionFormat 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6840
      Locked          =   -1  'True
      TabIndex        =   30
      Top             =   1800
      Width           =   1935
   End
   Begin VB.TextBox txtNativeCurrencyName 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6840
      Locked          =   -1  'True
      TabIndex        =   24
      Top             =   840
      Width           =   1935
   End
   Begin VB.TextBox txtIntlCurrencySymbol 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6840
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   600
      Width           =   1935
   End
   Begin VB.TextBox txtInternationalCurrencyDigits 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6840
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   360
      Width           =   1935
   End
   Begin VB.TextBox txtCurrencyGroupSize 
      Height          =   285
      Left            =   2400
      TabIndex        =   14
      Top             =   3000
      Width           =   1935
   End
   Begin VB.TextBox txtCurrencyDecimalSeperator 
      Height          =   285
      Left            =   2400
      TabIndex        =   10
      Top             =   2280
      Width           =   1935
   End
   Begin VB.TextBox txtCurrencyGroupSeperator 
      Height          =   285
      Left            =   2400
      TabIndex        =   12
      Top             =   2640
      Width           =   1935
   End
   Begin VB.TextBox txtFullEnglishCurrencyName 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6840
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   120
      Width           =   1935
   End
   Begin VB.TextBox txtCurrencySymbol 
      Height          =   285
      Left            =   2400
      TabIndex        =   16
      Top             =   3360
      Width           =   1935
   End
   Begin VB.CheckBox chkNegativeSpaceSeperation 
      Enabled         =   0   'False
      Height          =   255
      Left            =   8520
      TabIndex        =   28
      Top             =   1560
      Width           =   255
   End
   Begin VB.ComboBox cboNegativeCurrency 
      Height          =   315
      Left            =   6840
      TabIndex        =   26
      Top             =   1200
      Width           =   1935
   End
   Begin VB.ComboBox cboCurrency 
      Height          =   315
      Left            =   2400
      TabIndex        =   8
      Top             =   1920
      Width           =   1935
   End
   Begin VB.TextBox txtCurrencyDigits 
      Height          =   285
      Left            =   2400
      TabIndex        =   6
      Top             =   1560
      Width           =   1935
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   350
      Left            =   7800
      TabIndex        =   39
      Top             =   3120
      Width           =   975
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
   Begin VB.Label lblNativeCurrencyName 
      Caption         =   "Native Currency Namel"
      Height          =   255
      Left            =   4560
      TabIndex        =   23
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label lblCurrencyGroupSize 
      Caption         =   "Currency Group Size"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   3000
      Width           =   2055
   End
   Begin VB.Label lblCurrencyDecimalSeperator 
      Caption         =   "Currency Decimal Seperator"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Label lblIntlCurrencySymbol 
      Caption         =   "Intl Currency Symbol"
      Height          =   255
      Left            =   4560
      TabIndex        =   21
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label lblCurrencyGroupSeperator 
      Caption         =   "Currency Group Seperator"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   2640
      Width           =   2055
   End
   Begin VB.Label lblFullEnglishCurrencyName 
      Caption         =   "Full English Currency Name"
      Height          =   255
      Left            =   4560
      TabIndex        =   17
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label lblCurrencySymbol 
      Caption         =   "Currency Symbol"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   3360
      Width           =   2055
   End
   Begin VB.Label lblPositivePosition 
      Caption         =   "Positive Position"
      Height          =   255
      Left            =   4560
      TabIndex        =   37
      Top             =   2760
      Width           =   2055
   End
   Begin VB.Label lblPositivePositionFormat 
      Caption         =   "Positive Position Format"
      Height          =   255
      Left            =   4560
      TabIndex        =   35
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Label lblPositiveSpaceSeperation 
      Caption         =   "Positive Space Seperation"
      Height          =   255
      Left            =   4560
      TabIndex        =   33
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Label lblNegativePosition 
      Caption         =   "Negative Position"
      Height          =   255
      Left            =   4560
      TabIndex        =   31
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Label lblNegativePositionFormat 
      Caption         =   "Negative Position Format"
      Height          =   255
      Left            =   4560
      TabIndex        =   29
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Label lblNegativeSpaceSeperation 
      Caption         =   "Negative Space Seperation"
      Height          =   255
      Left            =   4560
      TabIndex        =   27
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label lblNegativeCurrency 
      Caption         =   "Negative Currency"
      Height          =   255
      Left            =   4560
      TabIndex        =   25
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label lblInternationalCurrencyDigits 
      Caption         =   "International Currency Digits"
      Height          =   255
      Left            =   4560
      TabIndex        =   19
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label lblCurrencyDigits 
      Caption         =   "Currency Digits"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label lblCurrency 
      Caption         =   "Currency"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1920
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
Attribute VB_Name = "frmLocalesCurrency"
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
    
    
    If SetLocaleInfo(lngLocale, LOCALE_ICURRDIGITS, txtCurrencyDigits.Text) = False Then Failed "SetLocaleInfo"
    If cboCurrency.ListIndex > -1 Then
        If SetLocaleInfo(lngLocale, LOCALE_ICURRENCY, CStr(cboCurrency.ListIndex)) = False Then Failed "SetLocaleInfo"
    End If
    If cboNegativeCurrency.ListIndex > -1 Then
        If SetLocaleInfo(lngLocale, LOCALE_INEGCURR, CStr(cboNegativeCurrency.ListIndex)) = False Then Failed "SetLocaleInfo"
    End If
    If SetLocaleInfo(lngLocale, LOCALE_SCURRENCY, txtCurrencySymbol.Text) = False Then Failed "SetLocaleInfo"
    If SetLocaleInfo(lngLocale, LOCALE_SMONTHOUSANDSEP, txtCurrencyGroupSeperator.Text) = False Then Failed "SetLocaleInfo"
    If SetLocaleInfo(lngLocale, LOCALE_SMONDECIMALSEP, txtCurrencyDecimalSeperator.Text) = False Then Failed "SetLocaleInfo"
    If SetLocaleInfo(lngLocale, LOCALE_SMONGROUPING, txtCurrencyGroupSize.Text) = False Then Failed "SetLocaleInfo"
End Sub

Private Sub cmdRefresh_Click()
    LocaleListNum = 0
    Erase LocaleList
    lstLocales.Clear
    
    txtCurrencyDigits.Text = ""
    cboCurrency.ListIndex = -1
    txtInternationalCurrencyDigits.Text = ""
    cboNegativeCurrency.ListIndex = -1
    chkNegativeSpaceSeperation.value = 0
    txtNegativePositionFormat.Text = ""
    txtNegativePosition.Text = ""
    chkPositiveSpaceSeperation.value = 0
    txtPositivePositionFormat.Text = ""
    txtPositivePosition.Text = ""
    txtCurrencySymbol.Text = ""
    txtFullEnglishCurrencyName.Text = ""
    txtCurrencyGroupSeperator.Text = ""
    txtIntlCurrencySymbol.Text = ""
    txtCurrencyDecimalSeperator.Text = ""
    txtCurrencyGroupSize.Text = ""
    txtNativeCurrencyName.Text = ""
    
    
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
    With cboCurrency
        .AddItem "Prefix, no separation"
        .AddItem "Suffix, no separation"
        .AddItem "Prefix, 1-character separation"
        .AddItem "Suffix, 1-character separation"
    End With
    With cboNegativeCurrency
        .AddItem "Left parenthesis,monetary symbol,number,right parenthesis"
        .AddItem "Negative sign, monetary symbol, number"
        .AddItem "Monetary symbol, negative sign, number"
        .AddItem "Monetary symbol, number, negative sign"
        .AddItem "Left parenthesis, number, monetary symbol, right parenthesis"
        .AddItem "Negative sign, number, monetary symbol"
        .AddItem "Number, negative sign, monetary symbol"
        .AddItem "Number, monetary symbol, negative sign"
        .AddItem "Negative sign, number, space, monetary symbol"
        .AddItem "Negative sign, monetary symbol, space, number"
        .AddItem "Number, space, monetary symbol, negative sign"
        .AddItem "Monetary symbol, space, number, negative sign"
        .AddItem "Monetary symbol, space, negative sign, number"
        .AddItem "Number, negative sign, space, monetary symbol"
        .AddItem "Left parenthesis, monetary symbol, space, number, right parenthesis"
        .AddItem "Left parenthesis, number, space, monetary symbol, right parenthesis"
    End With
    
    
    If WinVersion(-1, 5000000, True) = False Then
        lblFullEnglishCurrencyName.Enabled = False
        lblNativeCurrencyName.Enabled = False
    End If
    
    cmdRefresh_Click
End Sub

Private Sub lstLocales_Click()
    Dim lngLocale As Long
    lngLocale = strtoul_(lstLocales.List(lstLocales.ListIndex), 16)
    
    
    txtCurrencyDigits.Text = Get_LocaleInfo(lngLocale, LOCALE_ICURRDIGITS)
    cboCurrency.ListIndex = Val(Get_LocaleInfo(lngLocale, LOCALE_ICURRENCY))
    txtInternationalCurrencyDigits.Text = Get_LocaleInfo(lngLocale, LOCALE_IINTLCURRDIGITS)
    cboNegativeCurrency.ListIndex = Val(Get_LocaleInfo(lngLocale, LOCALE_INEGCURR))
    chkNegativeSpaceSeperation.value = Val(Get_LocaleInfo(lngLocale, LOCALE_INEGSEPBYSPACE))

    Select Case Val(Get_LocaleInfo(lngLocale, LOCALE_INEGSIGNPOSN))
        Case 0: txtNegativePositionFormat.Text = "Parentheses surround the amount and the monetary symbol"
        Case 1: txtNegativePositionFormat.Text = "The sign precedes the number"
        Case 2: txtNegativePositionFormat.Text = "The sign follows the number"
        Case 3: txtNegativePositionFormat.Text = "The sign precedes the monetary symbol"
        Case 4: txtNegativePositionFormat.Text = "The sign follows the monetary symbol"
        Case Else: txtNegativePositionFormat.Text = ""
    End Select
    Select Case Val(Get_LocaleInfo(lngLocale, LOCALE_INEGSYMPRECEDES))
        Case 0: txtNegativePosition.Text = "Follows Negative Amount"
        Case 1: txtNegativePosition.Text = "Precedes Negative Amount"
        Case Else: txtNegativePosition.Text = ""
    End Select
    
    chkPositiveSpaceSeperation.value = Val(Get_LocaleInfo(lngLocale, LOCALE_IPOSSEPBYSPACE))

    Select Case Val(Get_LocaleInfo(lngLocale, LOCALE_IPOSSIGNPOSN))
        Case 0: txtPositivePositionFormat.Text = "Parentheses surround the amount and the monetary symbol"
        Case 1: txtPositivePositionFormat.Text = "The sign precedes the number"
        Case 2: txtPositivePositionFormat.Text = "The sign follows the number"
        Case 3: txtPositivePositionFormat.Text = "The sign precedes the monetary symbol"
        Case 4: txtPositivePositionFormat.Text = "The sign follows the monetary symbol"
        Case Else: txtPositivePositionFormat.Text = ""
    End Select
    Select Case Val(Get_LocaleInfo(lngLocale, LOCALE_IPOSSYMPRECEDES))
        Case 0: txtPositivePosition.Text = "Follows Negative Amount"
        Case 1: txtPositivePosition.Text = "Precedes Negative Amount"
        Case Else: txtPositivePosition.Text = ""
    End Select
    
    txtCurrencySymbol.Text = Get_LocaleInfo(lngLocale, LOCALE_SCURRENCY)
    If WinVersion(-1, 5000000, True) = True Then
        txtFullEnglishCurrencyName.Text = Get_LocaleInfo(lngLocale, LOCALE_SENGCURRNAME)
    End If
    txtCurrencyGroupSeperator.Text = Get_LocaleInfo(lngLocale, LOCALE_SMONTHOUSANDSEP)
    txtIntlCurrencySymbol.Text = Get_LocaleInfo(lngLocale, LOCALE_SINTLSYMBOL)
    txtCurrencyDecimalSeperator.Text = Get_LocaleInfo(lngLocale, LOCALE_SMONDECIMALSEP)
    txtCurrencyGroupSize.Text = Get_LocaleInfo(lngLocale, LOCALE_SMONGROUPING)
    
    If WinVersion(-1, 5000000, True) = True Then
        txtNativeCurrencyName.Text = Get_LocaleInfo(lngLocale, LOCALE_SNATIVECURRNAME)
    End If
End Sub

Private Sub txtCurrencyDecimalSeperator_Change()
    If Len(txtCurrencyDecimalSeperator.Text) > 4 Then
        txtCurrencyDecimalSeperator.Text = Left$(txtCurrencyDecimalSeperator.Text, 4)
    End If
End Sub

Private Sub txtCurrencyDigits_Change()
    txtCurrencyDigits.Text = CStr(Val(Rem_NonNumeric_Chr(txtCurrencyDigits.Text)))
    If Val(txtCurrencyDigits.Text) < 0 Then txtCurrencyDigits.Text = "0"
    If Val(txtCurrencyDigits.Text) > 999 Then txtCurrencyDigits.Text = "999"
End Sub

Private Sub txtCurrencyGroupSeperator_Change()
    If Len(txtCurrencyGroupSeperator.Text) > 4 Then
        txtCurrencyGroupSeperator.Text = Left$(txtCurrencyGroupSeperator.Text, 4)
    End If
End Sub

Private Sub txtCurrencyGroupSize_Change()
    If Len(txtCurrencySymbol.Text) > 4 Then
        txtCurrencySymbol.Text = Left$(txtCurrencySymbol.Text, 4)
    End If
End Sub

Private Sub txtCurrencySymbol_Change()
    If Len(txtCurrencySymbol.Text) > 6 Then
        txtCurrencySymbol.Text = Left$(txtCurrencySymbol.Text, 6)
    End If
End Sub
