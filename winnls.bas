Attribute VB_Name = "winnls"
Option Explicit


Public Declare Function EnumSystemLocales Lib "kernel32.dll" Alias "EnumSystemLocalesA" (ByVal lpLocaleEnumProc As Long, ByVal dwFlags As Long) As Boolean
Public Declare Function GetLocaleInfo Lib "kernel32.dll" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
Public Declare Function MultiByteToWideChar Lib "kernel32.dll" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As String, ByVal cbMultiByte As Long, ByVal lpWideCharStr As String, ByVal cchWideChar As Long) As Long
Public Declare Function SetLocaleInfo Lib "kernel32.dll" Alias "SetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String) As Boolean
Public Declare Function WideCharToMultiByte Lib "kernel32.dll" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As String, ByVal cchWideChar As Long, ByVal lpMultiByteStr As String, ByVal cbMultiByte As Long, ByVal lpDefaultChar As String, ByVal lpUsedDefaultChar As Boolean) As Long


Public Const CP_ACP = 0
Public Const CP_OEMCP = 1
Public Const CP_MACCP = 2
Public Const CP_THREAD_ACP = 3
Public Const CP_SYMBOL = 42
Public Const CP_UTF7 = 65000
Public Const CP_UTF8 = 65001

Public Const LCID_INSTALLED = &H1
Public Const LCID_SUPPORTED = &H2
Public Const LCID_ALTERNATE_SORTS = &H4

Public Const LOCALE_NOUSEROVERRIDE = &H80000000
Public Const LOCALE_USE_CP_ACP = &H40000000
Public Const LOCALE_RETURN_NUMBER = &H20000000
Public Const LOCALE_ILANGUAGE = &H1
Public Const LOCALE_SLANGUAGE = &H2
Public Const LOCALE_SENGLANGUAGE = &H1001
Public Const LOCALE_SABBREVLANGNAME = &H3
Public Const LOCALE_SNATIVELANGNAME = &H4
Public Const LOCALE_ICOUNTRY = &H5
Public Const LOCALE_SCOUNTRY = &H6
Public Const LOCALE_SENGCOUNTRY = &H1002
Public Const LOCALE_SABBREVCTRYNAME = &H7
Public Const LOCALE_SNATIVECTRYNAME = &H8
Public Const LOCALE_IDEFAULTLANGUAGE = &H9
Public Const LOCALE_IDEFAULTCOUNTRY = &HA
Public Const LOCALE_IDEFAULTCODEPAGE = &HB
Public Const LOCALE_IDEFAULTANSICODEPAGE = &H1004
Public Const LOCALE_IDEFAULTMACCODEPAGE = &H1011
Public Const LOCALE_SLIST = &HC
Public Const LOCALE_IMEASURE = &HD
Public Const LOCALE_SDECIMAL = &HE
Public Const LOCALE_STHOUSAND = &HF
Public Const LOCALE_SGROUPING = &H10
Public Const LOCALE_IDIGITS = &H11
Public Const LOCALE_ILZERO = &H12
Public Const LOCALE_INEGNUMBER = &H1010
Public Const LOCALE_SNATIVEDIGITS = &H13
Public Const LOCALE_SCURRENCY = &H14
Public Const LOCALE_SINTLSYMBOL = &H15
Public Const LOCALE_SMONDECIMALSEP = &H16
Public Const LOCALE_SMONTHOUSANDSEP = &H17
Public Const LOCALE_SMONGROUPING = &H18
Public Const LOCALE_ICURRDIGITS = &H19
Public Const LOCALE_IINTLCURRDIGITS = &H1A
Public Const LOCALE_ICURRENCY = &H1B
Public Const LOCALE_INEGCURR = &H1C
Public Const LOCALE_SDATE = &H1D
Public Const LOCALE_STIME = &H1E
Public Const LOCALE_SSHORTDATE = &H1F
Public Const LOCALE_SLONGDATE = &H20
Public Const LOCALE_STIMEFORMAT = &H1003
Public Const LOCALE_IDATE = &H21
Public Const LOCALE_ILDATE = &H22
Public Const LOCALE_ITIME = &H23
Public Const LOCALE_ITIMEMARKPOSN = &H1005
Public Const LOCALE_ICENTURY = &H24
Public Const LOCALE_ITLZERO = &H25
Public Const LOCALE_IDAYLZERO = &H26
Public Const LOCALE_IMONLZERO = &H27
Public Const LOCALE_S1159 = &H28
Public Const LOCALE_S2359 = &H29
Public Const LOCALE_ICALENDARTYPE = &H1009
Public Const LOCALE_IOPTIONALCALENDAR = &H100B
Public Const LOCALE_IFIRSTDAYOFWEEK = &H100C
Public Const LOCALE_IFIRSTWEEKOFYEAR = &H100D
Public Const LOCALE_SDAYNAME1 = &H2A
Public Const LOCALE_SDAYNAME2 = &H2B
Public Const LOCALE_SDAYNAME3 = &H2C
Public Const LOCALE_SDAYNAME4 = &H2D
Public Const LOCALE_SDAYNAME5 = &H2E
Public Const LOCALE_SDAYNAME6 = &H2F
Public Const LOCALE_SDAYNAME7 = &H30
Public Const LOCALE_SABBREVDAYNAME1 = &H31
Public Const LOCALE_SABBREVDAYNAME2 = &H32
Public Const LOCALE_SABBREVDAYNAME3 = &H33
Public Const LOCALE_SABBREVDAYNAME4 = &H34
Public Const LOCALE_SABBREVDAYNAME5 = &H35
Public Const LOCALE_SABBREVDAYNAME6 = &H36
Public Const LOCALE_SABBREVDAYNAME7 = &H37
Public Const LOCALE_SMONTHNAME1 = &H38
Public Const LOCALE_SMONTHNAME2 = &H39
Public Const LOCALE_SMONTHNAME3 = &H3A
Public Const LOCALE_SMONTHNAME4 = &H3B
Public Const LOCALE_SMONTHNAME5 = &H3C
Public Const LOCALE_SMONTHNAME6 = &H3D
Public Const LOCALE_SMONTHNAME7 = &H3E
Public Const LOCALE_SMONTHNAME8 = &H3F
Public Const LOCALE_SMONTHNAME9 = &H40
Public Const LOCALE_SMONTHNAME10 = &H41
Public Const LOCALE_SMONTHNAME11 = &H42
Public Const LOCALE_SMONTHNAME12 = &H43
Public Const LOCALE_SMONTHNAME13 = &H100E
Public Const LOCALE_SABBREVMONTHNAME1 = &H44
Public Const LOCALE_SABBREVMONTHNAME2 = &H45
Public Const LOCALE_SABBREVMONTHNAME3 = &H46
Public Const LOCALE_SABBREVMONTHNAME4 = &H47
Public Const LOCALE_SABBREVMONTHNAME5 = &H48
Public Const LOCALE_SABBREVMONTHNAME6 = &H49
Public Const LOCALE_SABBREVMONTHNAME7 = &H4A
Public Const LOCALE_SABBREVMONTHNAME8 = &H4B
Public Const LOCALE_SABBREVMONTHNAME9 = &H4C
Public Const LOCALE_SABBREVMONTHNAME10 = &H4D
Public Const LOCALE_SABBREVMONTHNAME11 = &H4E
Public Const LOCALE_SABBREVMONTHNAME12 = &H4F
Public Const LOCALE_SABBREVMONTHNAME13 = &H100F
Public Const LOCALE_SPOSITIVESIGN = &H50
Public Const LOCALE_SNEGATIVESIGN = &H51
Public Const LOCALE_IPOSSIGNPOSN = &H52
Public Const LOCALE_INEGSIGNPOSN = &H53
Public Const LOCALE_IPOSSYMPRECEDES = &H54
Public Const LOCALE_IPOSSEPBYSPACE = &H55
Public Const LOCALE_INEGSYMPRECEDES = &H56
Public Const LOCALE_INEGSEPBYSPACE = &H57
Public Const LOCALE_FONTSIGNATURE = &H58
Public Const LOCALE_SISO639LANGNAME = &H59
Public Const LOCALE_SISO3166CTRYNAME = &H5A
Public Const LOCALE_IDEFAULTEBCDICCODEPAGE = &H1012
Public Const LOCALE_IPAPERSIZE = &H100A
Public Const LOCALE_SENGCURRNAME = &H1007
Public Const LOCALE_SNATIVECURRNAME = &H1008
Public Const LOCALE_SYEARMONTH = &H1006
Public Const LOCALE_SSORTNAME = &H1013
Public Const LOCALE_IDIGITSUBSTITUTION = &H1014


Public Function AsciiToUnicode(ByVal strAscii As String, ByVal lngFlags As Long) As String
    Dim strUnicode As String
    
    strUnicode = String$(Len(strAscii) * 2, &H0)
    apiError = MultiByteToWideChar(CP_ACP, lngFlags, strAscii, Len(strAscii), strUnicode, Len(strUnicode)): If apiError = 0 Then Failed "MultiByteToWideChar"
    
    AsciiToUnicode = Left$(strUnicode, apiError * 2)
End Function

Public Function EnumLocalesProc(ByRef lpLocaleString As Long) As Long 'Boolean
    Dim strLocale As String
    strLocale = String$(8, &H0)
    CopyMemory ByVal strLocale, lpLocaleString, ByVal Len(strLocale)
    
    LocaleListNum = LocaleListNum + 1
    ReDim Preserve LocaleList(LocaleListNum)
    LocaleList(LocaleListNum) = Fix_NullTermStr(strLocale)
    
    EnumLocalesProc = 1 'True
End Function

Public Function Get_LocaleInfo(lngLocale As Long, LCType As Long) As String
    Dim strBuffer As String
    strBuffer = String$(256, &H0)
    
    If GetLocaleInfo(lngLocale, LCType, strBuffer, Len(strBuffer)) = 0 Then Failed "GetLocaleInfo"
    
    Get_LocaleInfo = Fix_NullTermStr(strBuffer)
End Function

Public Function UnicodeToAscii(ByVal strUnicode As String, ByVal lngFlags As Long) As String
    Dim strAscii As String
    
    strAscii = String$(Len(strUnicode), &H0)
    apiError = WideCharToMultiByte(CP_ACP, lngFlags, strUnicode, Len(strUnicode), strAscii, Len(strAscii), &H0, False): If apiError = 0 Then Failed "WideCharToMultiByte"
    
    UnicodeToAscii = Left$(strAscii, apiError)
End Function
