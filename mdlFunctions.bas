Attribute VB_Name = "mdlFunctions"
Option Explicit


Public Function CLargeInt(ByVal Lo As Long, ByVal Hi As Long) As Double
    Dim dblLo As Double
    Dim dblHi As Double
    
    If Lo < 0 Then
        dblLo = 2 ^ 32 + Lo
    Else
        dblLo = Lo
    End If
    If Hi < 0 Then
        dblHi = 2 ^ 32 + Hi
    Else
        dblHi = Hi
    End If
    
    CLargeInt = dblLo + dblHi * 2 ^ 32
End Function

Public Function Fix_Dir(ByVal strDir As String) As String
    If Len(strDir) < 2 Then Exit Function
    
    If Right$(strDir, 1) = "\" Then
        Fix_Dir = Left$(strDir, Len(strDir) - 1)
    Else
        Fix_Dir = strDir
    End If
End Function

Public Function Fix_NullTermStr(ByVal strData As String) As String
    Dim pos As Long
    pos = InStr(1, strData, Chr$(0))
    
    If pos > 0 Then
        Fix_NullTermStr = Left$(strData, pos - 1)
    Else
        Fix_NullTermStr = strData
    End If
End Function

Public Function LangIdent(ByVal lngCode As Long) As String
    Select Case lngCode
        Case &H0: LangIdent = "Language Neutral"
        Case &H400: LangIdent = "Process Default Language"
        Case &H436: LangIdent = "Afrikaans"
        Case &H41C: LangIdent = "Albanian"
        Case &H401: LangIdent = "Arabic (Saudi Arabia)"
        Case &H801: LangIdent = "Arabic (Iraq)"
        Case &HC01: LangIdent = "Arabic (Egypt)"
        Case &H1001: LangIdent = "Arabic (Libya)"
        Case &H1401: LangIdent = "Arabic (Algeria)"
        Case &H1801: LangIdent = "Arabic (Morocco)"
        Case &H1C01: LangIdent = "Arabic (Tunisia)"
        Case &H2001: LangIdent = "Arabic (Oman)"
        Case &H2401: LangIdent = "Arabic (Yemen)"
        Case &H2801: LangIdent = "Arabic (Syria)"
        Case &H2C01: LangIdent = "Arabic (Jordan)"
        Case &H3001: LangIdent = "Arabic (Lebanon)"
        Case &H3401: LangIdent = "Arabic (Kuwait)"
        Case &H3801: LangIdent = "Arabic (U.A.E.)"
        Case &H3C01: LangIdent = "Arabic (Bahrain)"
        Case &H4001: LangIdent = "Arabic (Qatar)"
        Case &H42B: LangIdent = "Armenian"
        Case &H44D: LangIdent = "Assamese"
        Case &H42C: LangIdent = "Azeri (Latin)"
        Case &H82C: LangIdent = "Azeri (Cyrillic)"
        Case &H42D: LangIdent = "Basque"
        Case &H423: LangIdent = "Belarussian"
        Case &H445: LangIdent = "Bengali"
        Case &H402: LangIdent = "Bulgarian"
        Case &H455: LangIdent = "Burmese"
        Case &H403: LangIdent = "Catalan"
        Case &H404: LangIdent = "Chinese (Taiwan)"
        Case &H804: LangIdent = "Chinese (PRC)"
        Case &HC04: LangIdent = "Chinese (Hong Kong SAR, PRC)"
        Case &H1004: LangIdent = "Chinese (Singapore)"
        Case &H1404: LangIdent = "Chinese (Macau SAR)"
        Case &H41A: LangIdent = "Croatian"
        Case &H405: LangIdent = "Czech"
        Case &H406: LangIdent = "Danish"
        Case &H413: LangIdent = "Dutch (Netherlands)"
        Case &H813: LangIdent = "Dutch (Belgium)"
        Case &H409: LangIdent = "English (United States)"
        Case &H809: LangIdent = "English (United Kingdom)"
        Case &HC09: LangIdent = "English (Australian)"
        Case &H1009: LangIdent = "English (Canadian)"
        Case &H1409: LangIdent = "English (New Zealand)"
        Case &H1809: LangIdent = "English (Ireland)"
        Case &H1C09: LangIdent = "English (South Africa)"
        Case &H2009: LangIdent = "English (Jamaica)"
        Case &H2409: LangIdent = "English (Caribbean)"
        Case &H2809: LangIdent = "English (Belize)"
        Case &H2C09: LangIdent = "English (Trinidad)"
        Case &H3009: LangIdent = "English (Zimbabwe)"
        Case &H3409: LangIdent = "English (Philippines)"
        Case &H425: LangIdent = "Estonian"
        Case &H438: LangIdent = "Faeroese"
        Case &H429: LangIdent = "Farsi"
        Case &H40B: LangIdent = "Finnish"
        Case &H40C: LangIdent = "French (Standard)"
        Case &H80C: LangIdent = "French (Belgian)"
        Case &HC0C: LangIdent = "French (Canadian)"
        Case &H100C: LangIdent = "French (Switzerland)"
        Case &H140C: LangIdent = "French (Luxembourg)"
        Case &H180C: LangIdent = "French (Monaco)"
        Case &H43C: LangIdent = "Gaelic - Scotland"
        Case &H437: LangIdent = "Georgian"
        Case &H407: LangIdent = "German (Standard)"
        Case &H807: LangIdent = "German (Switzerland)"
        Case &HC07: LangIdent = "German (Austria)"
        Case &H1007: LangIdent = "German (Luxembourg)"
        Case &H1407: LangIdent = "German (Liechtenstein)"
        Case &H408: LangIdent = "Greek"
        Case &H447: LangIdent = "Gujarati"
        Case &H40D: LangIdent = "Hebrew"
        Case &H439: LangIdent = "Hindi"
        Case &H40E: LangIdent = "Hungarian"
        Case &H40F: LangIdent = "Icelandic"
        Case &H421: LangIdent = "Indonesian"
        Case &H410: LangIdent = "Italian (Standard)"
        Case &H810: LangIdent = "Italian (Switzerland)"
        Case &H411: LangIdent = "Japanese"
        Case &H44B: LangIdent = "Kannada"
        Case &H860: LangIdent = "Kashmiri (India)"
        Case &H43F: LangIdent = "Kazakh"
        Case &H457: LangIdent = "Konkani"
        Case &H412: LangIdent = "Korean"
        Case &H812: LangIdent = "Korean (Johab)"
        Case &H426: LangIdent = "Latvian"
        Case &H427: LangIdent = "Lithuanian"
        Case &H827: LangIdent = "Lithuanian (Classic)"
        Case &H42F: LangIdent = "Macedonian"
        Case &H43E: LangIdent = "Malay (Malaysian)"
        Case &H83E: LangIdent = "Malay (Brunei Darussalam)"
        Case &H44C: LangIdent = "Malayalam"
        Case &H43A: LangIdent = "Maltese"
        Case &H458: LangIdent = "Manipuri"
        Case &H44E: LangIdent = "Marathi"
        Case &H861: LangIdent = "Nepali (India)"
        Case &H414: LangIdent = "Norwegian (Bokmal)"
        Case &H814: LangIdent = "Norwegian (Nynorsk)"
        Case &H448: LangIdent = "Oriya"
        Case &H415: LangIdent = "Polish"
        Case &H416: LangIdent = "Portuguese (Brazil)"
        Case &H816: LangIdent = "Portuguese (Standard)"
        Case &H446: LangIdent = "Punjabi"
        Case &H417: LangIdent = "Raeto-Romance"
        Case &H418: LangIdent = "Romanian"
        Case &H818: LangIdent = "Romanian - Moldova"
        Case &H419: LangIdent = "Russian"
        Case &H819: LangIdent = "Russian - Moldova"
        Case &H44F: LangIdent = "Sanskrit"
        Case &HC1A: LangIdent = "Serbian (Cyrillic)"
        Case &H81A: LangIdent = "Serbian (Latin)"
        Case &H459: LangIdent = "Sindhi"
        Case &H41B: LangIdent = "Slovak"
        Case &H424: LangIdent = "Slovenian"
        Case &H42E: LangIdent = "Sorbian"
        Case &H40A: LangIdent = "Spanish (Traditional Sort)"
        Case &H80A: LangIdent = "Spanish (Mexican)"
        Case &HC0A: LangIdent = "Spanish (Modern Sort)"
        Case &H100A: LangIdent = "Spanish (Guatemala)"
        Case &H140A: LangIdent = "Spanish (Costa Rica)"
        Case &H180A: LangIdent = "Spanish (Panama)"
        Case &H1C0A: LangIdent = "Spanish (Dominican Republic)"
        Case &H200A: LangIdent = "Spanish (Venezuela)"
        Case &H240A: LangIdent = "Spanish (Colombia)"
        Case &H280A: LangIdent = "Spanish (Peru)"
        Case &H2C0A: LangIdent = "Spanish (Argentina)"
        Case &H300A: LangIdent = "Spanish (Ecuador)"
        Case &H340A: LangIdent = "Spanish (Chile)"
        Case &H380A: LangIdent = "Spanish (Uruguay)"
        Case &H3C0A: LangIdent = "Spanish (Paraguay)"
        Case &H400A: LangIdent = "Spanish (Bolivia)"
        Case &H440A: LangIdent = "Spanish (El Salvador)"
        Case &H480A: LangIdent = "Spanish (Honduras)"
        Case &H4C0A: LangIdent = "Spanish (Nicaragua)"
        Case &H500A: LangIdent = "Spanish (Puerto Rico)"
        Case &H430: LangIdent = "Sutu"
        Case &H441: LangIdent = "Swahili (Kenya)"
        Case &H41D: LangIdent = "Swedish"
        Case &H81D: LangIdent = "Swedish (Finland)"
        Case &H449: LangIdent = "Tamil"
        Case &H444: LangIdent = "Tatar (Tatarstan)"
        Case &H44A: LangIdent = "Telugu"
        Case &H41E: LangIdent = "Thai"
        Case &H431: LangIdent = "Tsonga"
        Case &H41F: LangIdent = "Turkish"
        Case &H422: LangIdent = "Ukrainian"
        Case &H420: LangIdent = "Urdu (Pakistan)"
        Case &H820: LangIdent = "Urdu (India)"
        Case &H443: LangIdent = "Uzbek (Latin)"
        Case &H843: LangIdent = "Uzbek (Cyrillic)"
        Case &H42A: LangIdent = "Vietnamese"
        Case &H434: LangIdent = "Xhosa"
        Case &H43D: LangIdent = "Yiddish"
        Case &H435: LangIdent = "Zulu"
    End Select
End Function

Public Function Percentage(ByVal dblValue As Double, ByVal dblTotal As Double, ByVal lngRound As Long) As Double
    If dblValue <> 0 Then
        If dblTotal <> 0 Then
            Percentage = Round((dblValue / dblTotal) * 100, lngRound)
        End If
    End If
End Function

Public Function Rem_NonFat_Chr(ByVal strData As String) As String
    If strData = "" Then Exit Function
    
    strData = Replace$(strData, "*", "", 1, -1)
    strData = Replace$(strData, "?", "", 1, -1)
    strData = Replace$(strData, "/", "", 1, -1)
    strData = Replace$(strData, "\", "", 1, -1)
    strData = Replace$(strData, "|", "", 1, -1)
    strData = Replace$(strData, ".", "", 1, -1)
    strData = Replace$(strData, ",", "", 1, -1)
    strData = Replace$(strData, ";", "", 1, -1)
    strData = Replace$(strData, ":", "", 1, -1)
    strData = Replace$(strData, "+", "", 1, -1)
    strData = Replace$(strData, "=", "", 1, -1)
    strData = Replace$(strData, " ", "", 1, -1)
    strData = Replace$(strData, "[", "", 1, -1)
    strData = Replace$(strData, "]", "", 1, -1)
    strData = Replace$(strData, "(", "", 1, -1)
    strData = Replace$(strData, ")", "", 1, -1)
    strData = Replace$(strData, "&", "", 1, -1)
    strData = Replace$(strData, "^", "", 1, -1)
    strData = Replace$(strData, "<", "", 1, -1)
    strData = Replace$(strData, ">", "", 1, -1)
    strData = Replace$(strData, Chr$(34), "", 1, -1)
    
    Rem_NonFat_Chr = strData
End Function

Public Function Rem_NonStd_Chr(ByVal strData As String) As String
    If strData = "" Then Exit Function
    
    
    Dim lngIncrement As Long
    
    For lngIncrement = 0 To 32
        strData = Replace$(strData, Chr$(lngIncrement), "", 1, -1)
    Next lngIncrement
    For lngIncrement = 42 To 44
        strData = Replace$(strData, Chr$(lngIncrement), "", 1, -1)
    Next lngIncrement
    For lngIncrement = 58 To 63
        strData = Replace$(strData, Chr$(lngIncrement), "", 1, -1)
    Next lngIncrement
    For lngIncrement = 91 To 93
        strData = Replace$(strData, Chr$(lngIncrement), "", 1, -1)
    Next lngIncrement
    For lngIncrement = 128 To 255
        strData = Replace$(strData, Chr$(lngIncrement), "", 1, -1)
    Next lngIncrement
    
    
    strData = Replace$(strData, Chr$(34), "", 1, -1)
    strData = Replace$(strData, Chr$(47), "", 1, -1)
    strData = Replace$(strData, Chr$(96), "", 1, -1)
    strData = Replace$(strData, Chr$(124), "", 1, -1)
    
    Rem_NonStd_Chr = strData
End Function

Public Function Rem_NonNumeric_Chr(ByVal strData As String) As String
    Dim lngIncrement As Long
    
    For lngIncrement = 0 To 44
        strData = Replace$(strData, Chr$(lngIncrement), "", 1, -1)
    Next lngIncrement
    
    strData = Replace$(strData, Chr$(46), "", 1, -1)
    strData = Replace$(strData, Chr$(47), "", 1, -1)
    
    For lngIncrement = 58 To 255
        strData = Replace$(strData, Chr$(lngIncrement), "", 1, -1)
    Next lngIncrement
    
    Rem_NonNumeric_Chr = strData
End Function

Public Function WinVersion(ByVal Windows As Long, ByVal NT As Long, ByVal Required As Boolean) As Boolean
    If WinID = VER_PLATFORM_WIN32_WINDOWS Then
        If Windows = -1 Then
            WinVersion = False
        Else
            If Required = True Then
                If Windows <= WinVer Then WinVersion = True
            Else
                If Windows > WinVer Then WinVersion = True
            End If
        End If
    End If
    If WinID = VER_PLATFORM_WIN32_NT Then
        If NT = -1 Then
            WinVersion = False
        Else
            If Required = True Then
                If NT <= WinVer Then WinVersion = True
            Else
                If NT > WinVer Then WinVersion = True
            End If
        End If
    End If
End Function
