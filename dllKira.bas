Attribute VB_Name = "dllKira"
Option Explicit


Public Declare Sub cpuid_ Lib "kira_ext.dll" (ByVal inpEAX As Long, ByRef outEAX As Long, ByRef outEBX As Long, ByRef outECX As Long, ByRef outEDX As Long)
Public Declare Sub rdtsc Lib "kira_ext.dll" Alias "rdtsc_" (ByRef tsc As LARGE_INTEGER)

Public Declare Function GET_X_LPARAM_ Lib "kira_ext.dll" (ByVal lParam As Long) As Long
Public Declare Function GET_Y_LPARAM_ Lib "kira_ext.dll" (ByVal lParam As Long) As Long
Public Declare Function HI_BYTE Lib "kira_ext.dll" (ByVal wValue As Integer) As Byte
Public Declare Function HI_WORD Lib "kira_ext.dll" (ByVal dwValue As Long) As Integer
Public Declare Function LO_BYTE Lib "kira_ext.dll" (ByVal wValue As Integer) As Byte
Public Declare Function LO_WORD Lib "kira_ext.dll" (ByVal dwValue As Long) As Integer
Public Declare Function MAKE_LONG Lib "kira_ext.dll" (ByVal wLow As Integer, ByVal wHigh As Integer) As Long
Public Declare Function MAKE_LPARAM Lib "kira_ext.dll" (ByVal wLow As Integer, ByVal wHigh As Integer) As Long
Public Declare Function MAKE_WPARAM Lib "kira_ext.dll" (ByVal wLow As Integer, ByVal wHigh As Integer) As Long
Public Declare Function MAKE_WORD Lib "kira_ext.dll" (ByVal bLow As Byte, ByVal bHigh As Byte) As Integer

Public Declare Function ltoa Lib "kira_ext.dll" Alias "ltoa_" (ByVal value As Long, ByVal radix As Long, ByVal buffer As String) As String
Public Declare Function strtol_ Lib "kira_ext.dll" (ByVal ptr As String, ByVal radix As Long) As Long
Public Declare Function strtoul_ Lib "kira_ext.dll" (ByVal ptr As String, ByVal radix As Long) As Long
Public Declare Function ultoa Lib "kira_ext.dll" Alias "ultoa_" (ByVal value As Long, ByVal radix As Long, ByVal buffer As String) As String

Public Declare Function checksum Lib "kira_ext.dll" (ByRef buffer As Any, ByVal size As Integer) As Integer
Public Declare Function typecast_int32_int16 Lib "kira_ext.dll" (ByVal int32 As Long) As Integer

Public Declare Function MouseHook_Install Lib "kira_ext.dll" (ByVal hwnd As Long) As Long
Public Declare Sub MouseHook_Remove Lib "kira_ext.dll" ()
Public Declare Function ShellHook_Install Lib "kira_ext.dll" (ByVal hwnd As Long) As Long
Public Declare Sub ShellHook_Remove Lib "kira_ext.dll" ()

'1.1.3
Public Declare Function adler32 Lib "kira_ext.dll" (ByVal adler As Long, ByVal buf As String, ByVal buf_len As Long) As Long
Public Declare Function crc32 Lib "kira_ext.dll" (ByVal crc As Long, ByVal buf As String, ByVal buf_len As Long) As Long


Public Function rdtsc_() As Double
    Dim tsc As LARGE_INTEGER
    rdtsc tsc
    
    rdtsc_ = CLargeInt(tsc.LowPart, tsc.HighPart)
End Function

Public Function ltoa_(ByVal value As Long, ByVal radix As Long) As String
    Dim buffer As String
    buffer = Space$(64)
    
    ltoa_ = UCase$(Trim$(Fix_NullTermStr(ltoa(value, radix, buffer))))
End Function

Public Function ultoa_(ByVal value As Long, ByVal radix As Long) As String
    Dim buffer As String
    buffer = Space$(64)
    
    ultoa_ = UCase$(Trim$(Fix_NullTermStr(ultoa(value, radix, buffer))))
End Function

Public Function MaxCPUIDLevel() As Long
    Dim outEAX As Long
    Dim outEBX As Long
    Dim outECX As Long
    Dim outEDX As Long
    
    cpuid_ 0, outEAX, outEBX, outECX, outEDX
    
    MaxCPUIDLevel = outEAX
End Function

Public Function MaxExtCPUIDLevel() As Long
    Dim outEAX As Long
    Dim outEBX As Long
    Dim outECX As Long
    Dim outEDX As Long
    
    cpuid_ strtoul_("80000000", 16), outEAX, outEBX, outECX, outEDX
    
    MaxExtCPUIDLevel = outEAX
End Function

Public Sub MouseHookInstall()
    If MouseHook = 0 Then
        MouseHook_OldProc = SetWindowLong(frmMain.txtMouseHook.hwnd, GWL_WNDPROC, AddressOf MouseHook_Proc): If MouseHook_OldProc = 0 Then Failed "GetWindowLong"
        MouseHook_Install frmMain.txtMouseHook.hwnd
    End If
    
    MouseHook = MouseHook + 1
End Sub

'Public Sub ShellHookInstall()
'    If ShellHook = 0 Then
'        ShellHook_OldProc = SetWindowLong(frmMain.txtShellHook.hwnd, GWL_WNDPROC, AddressOf ShellHook_Proc): If ShellHook_OldProc = 0 Then Failed "GetWindowLong"
'        ShellHook_Install frmMain.txtShellHook.hwnd
'    End If
'
'    ShellHook = ShellHook + 1
'End Sub

Public Sub MouseHookRemove()
    If MouseHook = 1 Then
        MouseHook_Remove
        If SetWindowLong(frmMain.txtMouseHook.hwnd, GWL_WNDPROC, MouseHook_OldProc) = 0 Then Failed "SetWindowLong"
    End If
    
    MouseHook = MouseHook - 1
End Sub

'Public Sub ShellHookRemove()
'    If ShellHook = 1 Then
'        ShellHook_Remove
'        If SetWindowLong(frmMain.txtShellHook.hwnd, GWL_WNDPROC, ShellHook_OldProc) = 0 Then Failed "SetWindowLong"
'    End If
'
'    ShellHook = ShellHook - 1
'End Sub


Public Function MouseHook_Proc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Select Case uMsg
        Dim POINTAPI As POINTAPI
        
        Case WM_LBUTTONUP
            With MouseMonitor
                .TotalLClicks = .TotalLClicks + 1
                frmMouseMonitor.txtLeft.Text = CStr(.TotalLClicks)
                frmMouseMonitor.txtTotalClicks.Text = CStr(.TotalLClicks + .TotalMClicks + .TotalRClicks + .TotalX1Clicks + .TotalX2Clicks)
            End With
            
        Case WM_MBUTTONUP
            With MouseMonitor
                .TotalMClicks = .TotalMClicks + 1
                frmMouseMonitor.txtMiddle.Text = CStr(.TotalMClicks)
                frmMouseMonitor.txtTotalClicks.Text = CStr(.TotalLClicks + .TotalMClicks + .TotalRClicks + .TotalX1Clicks + .TotalX2Clicks)
            End With
            
        Case WM_RBUTTONUP
            With MouseMonitor
                .TotalRClicks = .TotalRClicks + 1
                frmMouseMonitor.txtRight.Text = CStr(.TotalRClicks)
                frmMouseMonitor.txtTotalClicks.Text = CStr(.TotalLClicks + .TotalMClicks + .TotalRClicks + .TotalX1Clicks + .TotalX2Clicks)
            End With
        
        Case WM_XBUTTONUP
            Select Case HI_WORD(wParam)
                Case XBUTTON1
                    With MouseMonitor
                        .TotalX1Clicks = .TotalX1Clicks + 1
                        frmMouseMonitor.txtX1.Text = CStr(.TotalX1Clicks)
                        frmMouseMonitor.txtTotalClicks.Text = CStr(.TotalLClicks + .TotalMClicks + .TotalRClicks + .TotalX1Clicks + .TotalX2Clicks)
                    End With
                Case XBUTTON2
                    With MouseMonitor
                        .TotalX2Clicks = .TotalX2Clicks + 1
                        frmMouseMonitor.txtX2.Text = CStr(.TotalX2Clicks)
                        frmMouseMonitor.txtTotalClicks.Text = CStr(.TotalLClicks + .TotalMClicks + .TotalRClicks + .TotalX1Clicks + .TotalX2Clicks)
                    End With
            End Select
            
        Case WM_MOUSEMOVE
            Dim ScreenEdgeX As Long
            Dim ScreenEdgeY As Long
            ScreenEdgeX = Screen.Width \ Screen.TwipsPerPixelX
            ScreenEdgeY = Screen.Height \ Screen.TwipsPerPixelY
            
            POINTAPI.X = LO_WORD(lParam)
            POINTAPI.Y = HI_WORD(lParam)
            
            If frmMain.mnuMouseMonitorOO.Checked = True Then
                With MouseMonitor
                    .TotalXMovement = CDbl(Abs(POINTAPI.X - .LastCoordinate.X) + .TotalXMovement)
                    .LastCoordinate.X = POINTAPI.X
                    .TotalYMovement = CDbl(Abs(POINTAPI.Y - .LastCoordinate.Y) + .TotalYMovement)
                    .LastCoordinate.Y = POINTAPI.Y
                    
                    frmMouseMonitor.txtX.Text = CStr(.TotalXMovement)
                    frmMouseMonitor.txtY.Text = CStr(.TotalYMovement)
                    frmMouseMonitor.txtTotalMovement.Text = CStr(.TotalXMovement + .TotalYMovement)
                End With
            End If
            
            If frmMain.mnuMouseWarpOO.Checked = True Then
                If POINTAPI.X = ScreenEdgeX - 1 Then  'If at right edge reset to left
                    If SetCursorPos(1, POINTAPI.Y) = False Then Failed "SetCursorPos"
                    MouseMonitor.TotalWarp = MouseMonitor.TotalWarp + 1
                Else
                    If POINTAPI.X = 0 Then 'If at left edge reset to right
                        If SetCursorPos(ScreenEdgeX - 2, POINTAPI.Y) = False Then Failed "SetCursorPos"
                        MouseMonitor.TotalWarp = MouseMonitor.TotalWarp + 1
                    End If
                End If
                
                If POINTAPI.Y = ScreenEdgeY - 1 Then 'If at bottom edge reset to top
                    If SetCursorPos(POINTAPI.X, 1) = False Then Failed "SetCursorPos"
                    MouseMonitor.TotalWarp = MouseMonitor.TotalWarp + 1
                Else
                    If POINTAPI.Y = 0 Then 'If at top edge reset to bottom
                        If SetCursorPos(POINTAPI.X, ScreenEdgeY - 2) = False Then Failed "SetCursorPos"
                        MouseMonitor.TotalWarp = MouseMonitor.TotalWarp + 1
                    End If
                End If
                
                frmMouseWarp.txtWarps.Text = CStr(MouseMonitor.TotalWarp)
            End If
        Case WM_MOUSEWHEEL
            With MouseMonitor
                .TotalWheelMovement = .TotalWheelMovement + 1
                frmMouseMonitor.txtWheel.Text = CStr(.TotalWheelMovement)
            End With
            
        Case Else
            MouseHook_Proc = DefWindowProc(frmMain.txtMouseHook.hwnd, uMsg, wParam, lParam)
    End Select
End Function

'Public Function ShellHook_Proc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'    Select Case uMsg
'        Case Else
'            ShellHook_Proc = DefWindowProc(frmMain.txtShellHook.hwnd, uMsg, wParam, lParam)
'    End Select
'End Function
