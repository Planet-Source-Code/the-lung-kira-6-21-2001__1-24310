Attribute VB_Name = "winsock2"
Option Explicit


Public Declare Function closesocket Lib "ws2_32.dll" (ByVal s As Long) As Long
Public Declare Function connect Lib "ws2_32.dll" (ByVal s As Long, ByRef addr As sockaddr, ByVal namelen As Long) As Long
Public Declare Function gethostbyaddr Lib "ws2_32.dll" (ByRef addr As Long, ByRef addrlen As Long, ByVal addrType As Long) As Long
Public Declare Function gethostbyname Lib "ws2_32.dll" (ByVal host_name As String) As Long
Public Declare Function getprotobynumber Lib "ws2_32.dll" (ByVal number As Long) As Long
Public Declare Function getservbyport Lib "ws2_32.dll" (ByVal port As Long, ByRef proto As Any) As Long
Public Declare Function htonl Lib "ws2_32.dll" (ByVal hostlong As Long) As Long
Public Declare Function htons Lib "ws2_32.dll" (ByVal hostshort As Long) As Integer
Public Declare Function inet_addr Lib "ws2_32.dll" (ByVal cp As String) As Long
Public Declare Function inet_ntoa Lib "ws2_32.dll" (ByVal inn As Long) As Long
Public Declare Function ntohl Lib "ws2_32.dll" (ByVal netlong As Long) As Long
Public Declare Function ntohs Lib "ws2_32.dll" (ByVal netshort As Long) As Integer
Public Declare Function recv Lib "ws2_32.dll" (ByVal s As Long, ByVal buf As Any, ByVal buflen As Long, ByVal flags As Long) As Long
Public Declare Function recvfrom Lib "ws2_32.dll" (ByVal s As Long, ByVal buf As Any, ByVal buflen As Long, ByVal flags As Long, ByRef fromaddr As sockaddr, ByRef fromlen As Long) As Long
Public Declare Function send Lib "ws2_32.dll" (ByVal s As Long, ByRef buf As Any, ByVal buflen As Long, ByVal flags As Long) As Long
Public Declare Function sendto Lib "ws2_32.dll" (ByVal s As Long, ByRef buf As Any, ByVal buflen As Long, ByVal flags As Long, toaddr As sockaddr, ByVal tolen As Long) As Long
'Public Declare Function setsockopt Lib "ws2_32.dll" (ByVal s As Long, ByVal level As Long, ByVal optname As Long, ByRef optval As Any, ByVal optlen As Long) As Long
Public Declare Function shutdown Lib "ws2_32.dll" (ByVal s As Long, ByVal how As Long) As Long
Public Declare Function socket Lib "ws2_32.dll" (ByVal af As Long, ByVal s_type As Long, ByVal protocol As Long) As Long
Public Declare Function WSAAsyncSelect Lib "ws2_32.dll" (ByVal s As Long, ByVal hwnd As Long, ByVal wMsg As Long, ByVal lEvent As Long) As Long
Public Declare Function WSACleanup Lib "ws2_32.dll" () As Long
Public Declare Function WSAGetLastError Lib "ws2_32.dll" () As Long
Public Declare Function WSAStartup Lib "ws2_32.dll" (ByVal wVersionRequested As Long, ByRef lpWSAData As WSADATA) As Long


Public Const INVALID_SOCKET = &HFFFF
Public Const SOCKET_ERROR = -1

Public Const AF_UNSPEC = 0
Public Const AF_UNIX = 1
Public Const AF_INET = 2
Public Const AF_IMPLINK = 3
Public Const AF_PUP = 4
Public Const AF_CHAOS = 5
Public Const AF_NS = 6
Public Const AF_IPX = AF_NS
Public Const AF_ISO = 7
Public Const AF_OSI = AF_ISO
Public Const AF_ECMA = 8
Public Const AF_DATAKIT = 9
Public Const AF_CCITT = 10
Public Const AF_SNA = 11
Public Const AF_DECnet = 12
Public Const AF_DLI = 13
Public Const AF_LAT = 14
Public Const AF_HYLINK = 15
Public Const AF_APPLETALK = 16
Public Const AF_NETBIOS = 17
Public Const AF_VOICEVIEW = 18
Public Const AF_FIREFOX = 19
Public Const AF_UNKNOWN1 = 20
Public Const AF_BAN = 21
Public Const AF_ATM = 22
Public Const AF_INET6 = 23
Public Const AF_CLUSTER = 24
Public Const AF_12844 = 25
Public Const AF_IRDA = 26
Public Const AF_NETDES = 28

Public Const FD_READ = &H1
Public Const FD_WRITE = &H2
Public Const FD_OOB = &H4
Public Const FD_ACCEPT = &H8
Public Const FD_CONNECT = &H10
Public Const FD_CLOSE = &H20

Public Const ICMP_ECHO_REPLY = 0
Public Const ICMP_DESTINATION_UNREACHABLE = 3
Public Const ICMP_SOURCE_QUENCH = 4
Public Const ICMP_REDIRECT = 5
Public Const ICMP_ECHO = 8
Public Const ICMP_ROUTER_ADVERTISEMENT = 9
Public Const ICMP_ROUTER_SELECTION = 10
Public Const ICMP_TIME_EXCEEDED = 11
Public Const ICMP_PARAMETER_PROBLEM = 12
Public Const ICMP_TIMESTAMP = 13
Public Const ICMP_TIMESTAMP_REPLY = 14
Public Const ICMP_INFORMATION_REQUEST = 15
Public Const ICMP_INFORMATION_REPLY = 16
Public Const ICMP_ADDRESS_MASK_REQUEST = 17
Public Const ICMP_ADDRESS_MASK_REPLY = 18
Public Const ICMP_ADDRESS_TRACEROUTE = 30
Public Const ICMP_DATAGRAM_CONVERSION_ERROR = 31

Public Const INADDR_ANY = &H0
Public Const INADDR_LOOPBACK = &H7F000001
Public Const INADDR_BROADCAST = &HFFFFFFFF
Public Const INADDR_NONE = &HFFFFFFFF

Public Const IPPROTO_IP = 0
Public Const IPPROTO_ICMP = 1
Public Const IPPROTO_IGMP = 2
Public Const IPPROTO_GGP = 3
Public Const IPPROTO_TCP = 6
Public Const IPPROTO_PUP = 12
Public Const IPPROTO_UDP = 17
Public Const IPPROTO_IDP = 22
Public Const IPPROTO_ND = 77
Public Const IPPROTO_RAW = 255
Public Const IPPROTO_MAX = 256

Public Const SD_RECEIVE = &H0
Public Const SD_SEND = &H1
Public Const SD_BOTH = &H2

Public Const SOCK_STREAM = 1
Public Const SOCK_DGRAM = 2
Public Const SOCK_RAW = 3
Public Const SOCK_RDM = 4
Public Const SOCK_SEQPACKET = 5

Public Const WSA_DESCRIPTION_LEN = 256
Public Const WSA_SYS_STATUS_LEN = 128


Public Type hostent
    h_name As Long
    h_aliases As Long
    h_addrtype As Integer
    h_length As Integer
    h_addr_list As Long
End Type

Public Type ICMPHDR
    Type As Byte
    Code As Byte
    checksum As Integer
    ID As Integer
    Seq As Integer
    Data As String
End Type

Public Type IPHDR
    VIHL As Byte
    TOS  As Byte
    TotLen As Integer
    ID As Integer
    FlagOff As Integer
    TTL As Byte
    protocol As Byte
    checksum As Integer
    iaSrc As Long
    iaDst As Long
    options As String
End Type

Public Type ICMP_Packet
    IPHDR As IPHDR
    ICMPHDR As ICMPHDR
End Type

Public Type protoent
    p_name As Long
    p_aliases As Long
    p_proto As Integer
End Type

Public Type servent
    s_name As Long
    s_aliases As Long
    s_port As Integer
    s_proto As Long
End Type

Public Type sockaddr
    sin_family As Integer
    sin_port As Integer
    sin_addr As Long
    sin_zero As String * 8
End Type

Public Type WSADATA
    wVersion As Integer
    wHighVersion As Integer
    szDescription As String * 257   'WSADESCRIPTION_LEN + 1
    szSystemStatus As String * 129  'WSASYS_STATUS_LEN + 1
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpVendorInfo As Long
End Type


Public Sub Close_Socket(ByRef lngSocket As Long)
    If WS2 = True Then
        If lngSocket <> 0 Then
            If closesocket(lngSocket) = SOCKET_ERROR Then WinsockError "closesocket"
            lngSocket = 0
        End If
    End If
End Sub

Public Function GetHostByIP(ByVal strIP As String) As String
    If WS2 = True Then
        Dim hostent As hostent
        Dim lngIP As Long
        Dim strHost As String * 255
        
        lngIP = inet_addr(strIP)
        
        apiError = gethostbyaddr(lngIP, Len(lngIP), AF_INET)
        If apiError = &H0 Then
            Failed "gethostbyaddr"
            Exit Function
        End If
        
        CopyMemory hostent, ByVal apiError, Len(hostent)
        CopyMemory ByVal strHost, ByVal hostent.h_name, 255
        
        GetHostByIP = Fix_NullTermStr(strHost)
    End If
End Function

Public Function GetIPByHost(ByVal strHost As String, ByRef aryIP() As String) As Long
    If WS2 = True Then
        Dim hostent As hostent
        Dim lngHostIp As Long
        Dim strIP As String
        Dim lngIP As Long
        Dim lngValue As Long
        
        apiError = gethostbyname(strHost)
        If apiError = &H0 Then
            Failed "gethostbyname"
            Exit Function
        End If
        
        CopyMemory hostent, ByVal apiError, Len(hostent)
        CopyMemory lngHostIp, ByVal hostent.h_addr_list, hostent.h_length
        
        Do
            CopyMemory lngValue, ByVal lngHostIp, hostent.h_length
            
            lngValue = inet_ntoa(lngValue)
            strIP = String$(15, &H0)
            CopyMemory ByVal strIP, ByVal lngValue, 15
            
            
            lngIP = lngIP + 1
            ReDim Preserve aryIP(lngIP)
            aryIP(lngIP) = Fix_NullTermStr(strIP)
            
            
            hostent.h_addr_list = hostent.h_addr_list + LenB(hostent.h_addr_list)
            CopyMemory lngHostIp, ByVal hostent.h_addr_list, 4
        Loop While lngHostIp <> 0
        
        GetIPByHost = lngIP
    End If
End Function

Public Function HostIPToInAddr(ByVal HostIP As String) As Long
    If WS2 = True Then
        Dim lngReturn As Long
        lngReturn = inet_addr(HostIP)
        
        If lngReturn = INADDR_NONE Then
            Dim lngIP As Long
            Dim aryIP() As String
            lngIP = GetIPByHost(HostIP, aryIP())
            
            If lngIP > 0 Then
                lngReturn = inet_addr(aryIP(1) & Chr$(0))
                If lngReturn = INADDR_NONE Then
                    Failed "inet_addr"
                Else
                    HostIPToInAddr = lngReturn
                End If
            End If
        Else
            HostIPToInAddr = lngReturn
        End If
    End If
End Function

Public Sub WinsockError(ByVal apiFunction As String)
    If Err.LastDllError > 0 Then
        Errors Err.LastDllError, apiFunction
    Else
        If Function_Exist("ws2_32.dll", "WSAGetLastError") = True Then
            If WSAGetLastError > 0 Then
                Errors WSAGetLastError, apiFunction
            Else
                Failed apiFunction
            End If
        End If
    End If
End Sub

Public Sub Winsock_Startup()
    If Function_Exist("ws2_32.dll", "WSAStartup") = True Then
        'Winsock 2.2
        Dim WSADATA As WSADATA
        With WSADATA
            .szDescription = String$(256, &H0)
            .szSystemStatus = String$(128, &H0)
            .wHighVersion = 2
            .wVersion = 2
        End With
        
        apiError = WSAStartup(&H202, WSADATA)
        If apiError <> 0 Then
            Errors apiError, "WSAStartup"
            WS2 = False
        Else
            WS2 = True
        End If
    Else
        WS2 = False
    End If
End Sub

Public Sub Winsock_Shutdown()
    If WS2 = True Then
        apiError = WSACleanup: If apiError <> 0 Then Errors apiError, "WSACleanup"
    End If
End Sub
