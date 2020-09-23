Attribute VB_Name = "mdlWinsock"
Option Explicit


Public wsDayTime_OldProc As Long
Public wsDayTime_RTT As Double
Public wsDayTime_sockaddr As sockaddr
Public wsDayTime_Socket As Long

Public wsEcho_Data As String
Public wsEcho_OldProc As Long
Public wsEcho_RTT As Double
Public wsEcho_sockaddr As sockaddr
Public wsEcho_Socket As Long

Public wsICMP_Echo_OldProc As Long
Public wsICMP_Echo_Ret As Boolean
Public wsICMP_Echo_RTT As Double
Public wsICMP_Echo_RTTs() As Long
Public wsICMP_Echo_RTT_Num As Long
Public wsICMP_Echo_sockaddr As sockaddr
Public wsICMP_Echo_Socket As Long

Public wsName_Finger_OldProc As Long
Public wsName_Finger_sockaddr As sockaddr
Public wsName_Finger_Socket As Long

Public wsNicname_Whois_OldProc As Long
Public wsNicname_Whois_sockaddr As sockaddr
Public wsNicname_Whois_Socket As Long

Public wsQOTD_OldProc As Long
Public wsQOTD_RTT As Double
Public wsQOTD_sockaddr As sockaddr
Public wsQOTD_Socket As Long

Public wsTime_OldProc As Long
Public wsTime_RTT As Double
Public wsTime_SetTime As Boolean
Public wsTime_sockaddr As sockaddr
Public wsTime_Socket As Long


Public Function wsDayTime_Proc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Select Case uMsg
        Case WM_PROJECT_WS
            Select Case LO_WORD(lParam)
                Case FD_READ
                    wsDayTime_RTT = PerformanceCounter - wsDayTime_RTT
                    
                    Dim strBuffer As String
                    strBuffer = String$(65536, Chr$(0))
                    
                    Select Case frmDayTime.cboMethod.ListIndex
                        Case 0 'UDP
                            apiError = recvfrom(wsDayTime_Socket, strBuffer, Len(strBuffer), 0, wsDayTime_sockaddr, Len(wsDayTime_sockaddr)): If apiError = SOCKET_ERROR Then WinsockError "recvfrom"
                            If apiError > 0 Then strBuffer = Left$(strBuffer, apiError)
                            
                            If shutdown(wsDayTime_Socket, SD_BOTH) = SOCKET_ERROR Then WinsockError "shutdown"
                        Case 1 'TCP
                            apiError = recv(wsDayTime_Socket, strBuffer, Len(strBuffer), 0): If apiError = SOCKET_ERROR Then WinsockError "recv"
                            If apiError > 0 Then strBuffer = Left$(strBuffer, apiError)
                            
                            If shutdown(wsDayTime_Socket, SD_BOTH) = SOCKET_ERROR Then WinsockError "shutdown"
                    End Select
                    
                    With frmDayTime
                        .txtReturned.Text = strBuffer
                        .txtRoundTripTime.Text = CStr(Round((wsDayTime_RTT / dblCounterFrequency) * 1000, 0))
                        .cmdStop.Enabled = False
                        .cmdGetData.Enabled = True
                    End With
                    
                Case FD_CLOSE
                    Close_Socket wsDayTime_Socket
                    
                    frmDayTime.cmdStop.Enabled = False
                    frmDayTime.cmdGetData.Enabled = True
                    wsDayTime_RTT = 0
            End Select
            
            'wsError = HI_WORD(lParam)
        Case Else
            wsDayTime_Proc = DefWindowProc(frmDayTime.txtDayTime.hwnd, uMsg, wParam, lParam)
    End Select
End Function

Public Function wsEcho_Proc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Select Case uMsg
        Case WM_PROJECT_WS
            Select Case LO_WORD(lParam)
                Case FD_READ
                    wsEcho_RTT = PerformanceCounter - wsEcho_RTT
                    
                    Dim strBuffer As String
                    strBuffer = String$(65536, Chr$(0))
                    
                    Select Case frmEcho.cboMethod.ListIndex
                        Case 0 'UDP
                            apiError = recvfrom(wsEcho_Socket, strBuffer, Len(strBuffer), 0, wsEcho_sockaddr, Len(wsEcho_sockaddr)): If apiError = SOCKET_ERROR Then WinsockError "recvfrom"
                            If apiError > 0 Then strBuffer = Left$(strBuffer, apiError)
                            
                            If shutdown(wsEcho_Socket, SD_BOTH) = SOCKET_ERROR Then WinsockError "shutdown"
                        Case 1 'TCP
                            apiError = recv(wsEcho_Socket, strBuffer, Len(strBuffer), 0): If apiError = SOCKET_ERROR Then WinsockError "recv"
                            If apiError > 0 Then strBuffer = Left$(strBuffer, apiError)
                            
                            If shutdown(wsEcho_Socket, SD_BOTH) = SOCKET_ERROR Then WinsockError "shutdown"
                    End Select
                    
                    With frmEcho
                        If strBuffer = wsEcho_Data Then .chkReturnOK.value = 1
                        .txtRoundTripTime.Text = CStr(Round((wsEcho_RTT / dblCounterFrequency) * 1000, 0))
                        .cmdStop.Enabled = False
                        .cmdSendData.Enabled = True
                    End With
                    
                Case FD_CLOSE
                    Close_Socket wsEcho_Socket
                    
                    frmEcho.cmdStop.Enabled = False
                    frmEcho.cmdSendData.Enabled = True
                    wsEcho_RTT = 0
                    wsEcho_Data = ""
            End Select
            
            'wsError = HI_WORD(lParam)
        Case Else
            wsEcho_Proc = DefWindowProc(frmEcho.txtEcho.hwnd, uMsg, wParam, lParam)
    End Select
End Function

Public Function wsICMP_Echo_Proc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Select Case uMsg
        Case WM_PROJECT_WS
            Select Case LO_WORD(lParam)
                Case FD_READ
                    wsICMP_Echo_RTT = PerformanceCounter - wsICMP_Echo_RTT
                    
                    Dim strBuffer As String
                    strBuffer = String$(65536, Chr$(0))
                    
                    apiError = recvfrom(wsICMP_Echo_Socket, strBuffer, Len(strBuffer), 0, wsICMP_Echo_sockaddr, Len(wsICMP_Echo_sockaddr)): If apiError = SOCKET_ERROR Then WinsockError "recvfrom"
                    If apiError > 0 Then strBuffer = Left$(strBuffer, apiError)
                    
                    If Len(strBuffer) >= 20 Then
                        Dim IPHDR As IPHDR
                        With IPHDR
                            .VIHL = strtoul_(Right$("00" & ltoa_(Asc(Mid$(strBuffer, 1, 1)), 16), 2), 16)
                            .TOS = strtoul_(Right$("00" & ltoa_(Asc(Mid$(strBuffer, 2, 1)), 16), 2), 16)
                            .TotLen = typecast_int32_int16(strtoul_(Right$("00" & ltoa_(Asc(Mid$(strBuffer, 4, 1)), 16), 2) & _
                                              Right$("00" & ltoa_(Asc(Mid$(strBuffer, 3, 1)), 16), 2), 16))
                            .ID = typecast_int32_int16(strtoul_(Right$("00" & ltoa_(Asc(Mid$(strBuffer, 6, 1)), 16), 2) & _
                                          Right$("00" & ltoa_(Asc(Mid$(strBuffer, 5, 1)), 16), 2), 16))
                            .FlagOff = typecast_int32_int16(strtoul_(Right$("00" & ltoa_(Asc(Mid$(strBuffer, 8, 1)), 16), 2) & _
                                          Right$("00" & ltoa_(Asc(Mid$(strBuffer, 7, 1)), 16), 2), 16))
                            .TTL = strtoul_(Right$("00" & ltoa_(Asc(Mid$(strBuffer, 9, 1)), 16), 2), 16)
                            .protocol = strtoul_(Right$("00" & ltoa_(Asc(Mid$(strBuffer, 10, 1)), 16), 2), 16)
                            .checksum = typecast_int32_int16(strtoul_(Right$("00" & ltoa_(Asc(Mid$(strBuffer, 12, 1)), 16), 2) & _
                                                Right$("00" & ltoa_(Asc(Mid$(strBuffer, 11, 1)), 16), 2), 16))
                            .iaSrc = strtoul_(Right$("00" & ltoa_(Asc(Mid$(strBuffer, 16, 1)), 16), 2) & _
                                              Right$("00" & ltoa_(Asc(Mid$(strBuffer, 15, 1)), 16), 2) & _
                                              Right$("00" & ltoa_(Asc(Mid$(strBuffer, 14, 1)), 16), 2) & _
                                              Right$("00" & ltoa_(Asc(Mid$(strBuffer, 13, 1)), 16), 2), 16)
                            .iaDst = strtoul_(Right$("00" & ltoa_(Asc(Mid$(strBuffer, 20, 1)), 16), 2) & _
                                              Right$("00" & ltoa_(Asc(Mid$(strBuffer, 19, 1)), 16), 2) & _
                                              Right$("00" & ltoa_(Asc(Mid$(strBuffer, 18, 1)), 16), 2) & _
                                              Right$("00" & ltoa_(Asc(Mid$(strBuffer, 17, 1)), 16), 2), 16)
                        End With
                        
                        Dim lngIPOffset As Long
                        lngIPOffset = (strtoul_(Mid$(ltoa_(IPHDR.VIHL, 2), 4, 4), 2) * 32) / 8
                        
                        Dim ICMPHDR As ICMPHDR
                        If Len(strBuffer) >= (lngIPOffset + 8) Then
                            With ICMPHDR
                                .Type = strtoul_(Right$("00" & ltoa_(Asc(Mid$(strBuffer, lngIPOffset + 1, 1)), 16), 2), 16)
                                .Code = strtoul_(Right$("00" & ltoa_(Asc(Mid$(strBuffer, lngIPOffset + 2, 1)), 16), 2), 16)
                                .checksum = typecast_int32_int16(strtoul_(Right$("00" & ltoa_(Asc(Mid$(strBuffer, lngIPOffset + 4, 1)), 16), 2) & _
                                                    Right$("00" & ltoa_(Asc(Mid$(strBuffer, lngIPOffset + 3, 1)), 16), 2), 16))
                                .ID = typecast_int32_int16(strtoul_(Right$("00" & ltoa_(Asc(Mid$(strBuffer, lngIPOffset + 6, 1)), 16), 2) & _
                                              Right$("00" & ltoa_(Asc(Mid$(strBuffer, lngIPOffset + 5, 1)), 16), 2), 16))
                                .Seq = typecast_int32_int16(strtoul_(Right$("00" & ltoa_(Asc(Mid$(strBuffer, lngIPOffset + 8, 1)), 16), 2) & _
                                               Right$("00" & ltoa_(Asc(Mid$(strBuffer, lngIPOffset + 7, 1)), 16), 2), 16))
                            End With
                        End If
                        
                        If ICMPHDR.ID = GetCurrentProcessId Then
                            wsICMP_Echo_RTT_Num = wsICMP_Echo_RTT_Num + 1
                            ReDim Preserve wsICMP_Echo_RTTs(wsICMP_Echo_RTT_Num)
                            
                            Dim lngChecksum As Long
                            lngChecksum = IPHDR.checksum
                            IPHDR.checksum = 0
                            
                            If checksum(IPHDR, Len(IPHDR)) = lngChecksum Then
                                lngChecksum = ICMPHDR.checksum
                                ICMPHDR.checksum = 0
                                
                                If checksum(ICMPHDR, Len(ICMPHDR)) = lngChecksum Then
                                    wsICMP_Echo_RTTs(wsICMP_Echo_RTT_Num) = Round((wsICMP_Echo_RTT / dblCounterFrequency) * 1000, 0)
                                    If wsICMP_Echo_RTTs(wsICMP_Echo_RTT_Num) < 0 Then wsICMP_Echo_RTTs(wsICMP_Echo_RTT_Num) = 0
                                    frmICMP_Echo.lstRoundTripTime.AddItem Left$(ICMPHDR.Seq & Space$(7), 7) & CStr(wsICMP_Echo_RTTs(wsICMP_Echo_RTT_Num))
                                Else
                                    wsICMP_Echo_RTTs(wsICMP_Echo_RTT_Num) = -1
                                    frmICMP_Echo.lstRoundTripTime.AddItem Left$(ICMPHDR.Seq & Space$(7), 7) & "Failed"
                                End If
                            Else
                                wsICMP_Echo_RTTs(wsICMP_Echo_RTT_Num) = -1
                            End If
                        End If
                    End If
                    
                    wsICMP_Echo_Ret = True
                Case FD_CLOSE
                    Close_Socket wsICMP_Echo_Socket
                    
                    wsICMP_Echo_RTT = 0
                    
                    frmICMP_Echo.cmdStop.Enabled = False
                    frmICMP_Echo.cmdSend.Enabled = True
            End Select
            
            'wsError = HI_WORD(lParam)
        Case Else
            wsICMP_Echo_Proc = DefWindowProc(frmICMP_Echo.txtICMP_Echo.hwnd, uMsg, wParam, lParam)
    End Select
End Function

Public Function wsName_Finger_Proc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Select Case uMsg
        Case WM_PROJECT_WS
            Select Case LO_WORD(lParam)
                Case FD_READ
                    Dim strBuffer As String
                    strBuffer = String$(65536, Chr$(0))
                    
                    apiError = recv(wsName_Finger_Socket, strBuffer, Len(strBuffer), 0): If apiError = SOCKET_ERROR Then WinsockError "recv"
                    If apiError > 0 Then strBuffer = Left$(strBuffer, apiError)
                    
                    frmName_Finger.txtReturned.Text = frmName_Finger.txtReturned.Text & Replace$(strBuffer, Chr$(10), Chr$(13) & Chr$(10), 1, -1)
                    
                Case FD_CLOSE
                    Close_Socket wsName_Finger_Socket
                    
                    frmName_Finger.cmdStop.Enabled = False
                    frmName_Finger.cmdSendData.Enabled = True
            End Select
            
            'wsError = HI_WORD(lParam)
        Case Else
            wsName_Finger_Proc = DefWindowProc(frmName_Finger.txtName_Finger.hwnd, uMsg, wParam, lParam)
    End Select
End Function

Public Function wsNicname_Whois_Proc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Select Case uMsg
        Case WM_PROJECT_WS
            Select Case LO_WORD(lParam)
                Case FD_READ
                    Dim strBuffer As String
                    strBuffer = String$(65536, Chr$(0))
                    
                    apiError = recv(wsNicname_Whois_Socket, strBuffer, Len(strBuffer), 0): If apiError = SOCKET_ERROR Then WinsockError "recv"
                    If apiError > 0 Then strBuffer = Left$(strBuffer, apiError)
                    
                    frmNicname_Whois.txtReturned.Text = frmNicname_Whois.txtReturned.Text & Replace$(strBuffer, Chr$(10), Chr$(13) & Chr$(10), 1, -1)
                    
                Case FD_CLOSE
                    Close_Socket wsNicname_Whois_Socket
                    
                    frmNicname_Whois.cmdStop.Enabled = False
                    frmNicname_Whois.cmdSendData.Enabled = True
            End Select
            
            'wsError = HI_WORD(lParam)
        Case Else
            wsNicname_Whois_Proc = DefWindowProc(frmNicname_Whois.txtNicname_Whois.hwnd, uMsg, wParam, lParam)
    End Select
End Function

Public Function wsQOTD_Proc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Select Case uMsg
        Case WM_PROJECT_WS
            Select Case LO_WORD(lParam)
                Case FD_READ
                    wsQOTD_RTT = PerformanceCounter - wsQOTD_RTT
                    
                    Dim strBuffer As String
                    strBuffer = String$(65536, Chr$(0))
                    
                    Select Case frmQOTD.cboMethod.ListIndex
                        Case 0 'UDP
                            apiError = recvfrom(wsQOTD_Socket, strBuffer, Len(strBuffer), 0, wsQOTD_sockaddr, Len(wsQOTD_sockaddr)): If apiError = SOCKET_ERROR Then WinsockError "recvfrom"
                            If apiError > 0 Then strBuffer = Left$(strBuffer, apiError)
                            
                            If shutdown(wsQOTD_Socket, SD_BOTH) = SOCKET_ERROR Then WinsockError "shutdown"
                        Case 1 'TCP
                            apiError = recv(wsQOTD_Socket, strBuffer, Len(strBuffer), 0): If apiError = SOCKET_ERROR Then WinsockError "recv"
                            If apiError > 0 Then strBuffer = Left$(strBuffer, apiError)
                            
                            If shutdown(wsQOTD_Socket, SD_BOTH) = SOCKET_ERROR Then WinsockError "shutdown"
                    End Select
                    
                    With frmQOTD
                        .txtReturned.Text = strBuffer
                        .txtRoundTripTime.Text = CStr(Round((wsQOTD_RTT / dblCounterFrequency) * 1000, 0))
                        .cmdStop.Enabled = False
                        .cmdGetData.Enabled = True
                    End With
                    
                Case FD_CLOSE
                    Close_Socket wsQOTD_Socket
                    
                    frmQOTD.cmdStop.Enabled = False
                    frmQOTD.cmdGetData.Enabled = True
                    wsQOTD_RTT = 0
            End Select
            
            'wsError = HI_WORD(lParam)
        Case Else
            wsQOTD_Proc = DefWindowProc(frmQOTD.txtQOTD.hwnd, uMsg, wParam, lParam)
    End Select
End Function

Public Function wsTime_Proc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Select Case uMsg
        Case WM_PROJECT_WS
            Select Case LO_WORD(lParam)
                Case FD_READ
                    wsTime_RTT = PerformanceCounter - wsTime_RTT
                    
                    Dim strBuffer As String
                    strBuffer = String$(65535, Chr$(0))
                    
                    Select Case frmTime.cboMethod.ListIndex
                        Case 0 'UDP
                            apiError = recvfrom(wsTime_Socket, strBuffer, Len(strBuffer), 0, wsTime_sockaddr, Len(wsTime_sockaddr)): If apiError = SOCKET_ERROR Then WinsockError "recvfrom"
                            If apiError > 0 Then strBuffer = Left$(strBuffer, apiError)
                            
                            If shutdown(wsTime_Socket, SD_BOTH) = SOCKET_ERROR Then WinsockError "shutdown"
                        Case 1 'TCP
                            apiError = recv(wsTime_Socket, strBuffer, Len(strBuffer), 0): If apiError = SOCKET_ERROR Then WinsockError "recv"
                            If apiError > 0 Then strBuffer = Left$(strBuffer, apiError)
                            
                            If shutdown(wsTime_Socket, SD_BOTH) = SOCKET_ERROR Then WinsockError "shutdown"
                    End Select
                    
                    
                    If Len(strBuffer) = 4 Then
                        Dim varTime As Variant
                        Dim dblTime As Double
                        
                        dblTime = strtoul_(Right$("00" & ltoa_(Asc(Mid$(strBuffer, 1, 1)), 16), 2) & _
                                           Right$("00" & ltoa_(Asc(Mid$(strBuffer, 2, 1)), 16), 2) & _
                                           Right$("00" & ltoa_(Asc(Mid$(strBuffer, 3, 1)), 16), 2) & _
                                           Right$("00" & ltoa_(Asc(Mid$(strBuffer, 4, 1)), 16), 2), 16)
                        
                        If dblTime < 0 Then dblTime = (2 ^ 32) + dblTime
                        
                        varTime = DateAdd("s", dblTime - 2208988800#, "1/1/1970")
                        If wsTime_SetTime = True Then
                            Dim SYSTEMTIME As SYSTEMTIME
                            With SYSTEMTIME
                                .wYear = Year(varTime)
                                .wMonth = Month(varTime)
                                .wDay = Day(varTime)
                                .wHour = Hour(varTime)
                                .wMinute = Minute(varTime)
                                .wSecond = Second(varTime)
                                .wMilliseconds = Round(((wsTime_RTT / dblCounterFrequency) * 1000) / 2, 0)
                            End With
                            
                            If SetSystemTime(SYSTEMTIME) = False Then Failed "SetSystemTime"
                            
                            If WinVersion(0, 5000000, False) = True Then
                                SendMessage HWND_TOPMOST, WM_TIMECHANGE, 0, 0
                            End If
                        End If
                        
                        
                        frmTime.txtUnFormatted.Text = CStr(dblTime)
                        frmTime.txtReturnedGMT = CStr(varTime)
                        
                        
                        Dim TIME_ZONE_INFORMATION As TIME_ZONE_INFORMATION
                        Dim Bias As Long
                        
                        apiError = GetTimeZoneInformation(TIME_ZONE_INFORMATION)
                        Select Case apiError
                            Case TIME_ZONE_ID_INVALID: Failed "GetTimeZoneInformation"
                            Case TIME_ZONE_ID_UNKNOWN: Failed "GetTimeZoneInformation"
                        End Select
                        
                        If TIME_ZONE_INFORMATION.Bias < 0 Then
                            Bias = Abs(TIME_ZONE_INFORMATION.Bias)
                        Else
                            Bias = TIME_ZONE_INFORMATION.Bias - (TIME_ZONE_INFORMATION.Bias * 2)
                        End If
                        
                        
                        varTime = DateAdd("n", Bias, varTime)
                        If frmTime.chkDaylightSavings.value = 1 Then varTime = DateAdd("n", Abs(TIME_ZONE_INFORMATION.DaylightBias), varTime)
                    End If
                    
                    With frmTime
                        .txtReturnedLocal.Text = CStr(varTime)
                        .txtRoundTripTime.Text = CStr(Round((wsTime_RTT / dblCounterFrequency) * 1000, 0))
                        .cmdSetTime.Enabled = True
                        .cmdGetData.Enabled = True
                        .cmdStop.Enabled = False
                    End With
                    
                Case FD_CLOSE
                    Close_Socket wsTime_Socket
                    
                    frmTime.cmdStop.Enabled = False
                    frmTime.cmdGetData.Enabled = True
                    wsTime_RTT = 0
                    wsTime_SetTime = False
            End Select
        
        Case Else
            wsTime_Proc = DefWindowProc(frmTime.txtTime.hwnd, uMsg, wParam, lParam)
    End Select
End Function
