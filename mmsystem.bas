Attribute VB_Name = "mmsystem"
Option Explicit


Public Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long


Public Sub mciError(ByVal lngError, ByVal strFunction As String, Optional ByRef errDescription As String, Optional ByVal NoMsgBox As Boolean)
    Dim lenError As Long
    
    errDescription = String$(2048, &H0)
    apiError = mciGetErrorString(lngError, errDescription, Len(errDescription))
    
    If errDescription = "" Then
        errDescription = "No description available."
    End If
    
    If NoMsgBox = False Then
        If errMsg = True Then
            MessageBoxEx &H0, strFunction & Chr$(13) & Chr$(10) & Chr$(13) & Chr$(10) & errDescription, "Error", MB_OK Or MB_ICONWARNING Or MB_SETFOREGROUND, 0
        End If
    End If
End Sub
