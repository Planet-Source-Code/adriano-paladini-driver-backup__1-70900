Attribute VB_Name = "md_reboot"
Private Type LUID
   LowPart As Long
   HighPart As Long
End Type

Private Type TOKEN_PRIVILEGES
    PrivilegeCount As Long
    LuidUDT As LUID
   Attributes As Long
End Type

Const TOKEN_ADJUST_PRIVILEGES = &H20
Const TOKEN_QUERY = &H8
Const SE_PRIVILEGE_ENABLED = &H2

Const EWX_SHUTDOWN = 1
Const EWX_REBOOT = 2
Const EWX_FORCE = 4

Private Declare Function GetVersion Lib "kernel32" () As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function OpenProcessToken Lib "advapi32" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function LookupPrivilegeValue Lib "advapi32" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LUID) As Long
Private Declare Function AdjustTokenPrivileges Lib "advapi32" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As Any, ReturnLength As Any) As Long
Private Declare Function ExitWindowsEx Lib "user32" (ByVal dwOptions As Long, ByVal dwReserved As Long) As Long

' Shut down windows, and optional reboot it
' if the 2nd argument is True, no WM_QUERYENDSESSION and WM_ENDSESSION
' messages are sent to active applications

Sub ShutDownWindows(ByVal Reboot As Boolean, Optional ByVal Force As Boolean)
    Dim hToken As Long
    Dim tp As TOKEN_PRIVILEGES
    Dim flags As Long
    
    ' Windows NT/2000 require a special treatment
    ' to ensure that the calling process has the
    ' privileges to shut down the system
    
    ' under NT the high-order bit (that is, the sign bit)
    ' of the value retured by GetVersion is cleared
    If GetVersion() >= 0 Then
        ' Open this process for adjusting its privileges
        OpenProcessToken GetCurrentProcess(), (TOKEN_ADJUST_PRIVILEGES Or _
            TOKEN_QUERY), hToken
        
        ' Get the LUID for shutdown privilege.
        ' retrieves the locally unique identifier (LUID) used
        ' to locally represent the specified privilege name
        ' (first argument = "" means the local system)
        LookupPrivilegeValue "", "SeShutdownPrivilege", tp.LuidUDT
        
        ' complete the TOKEN_PRIVILEGES structure with the # of
        ' privileges and the desired attribute
        tp.PrivilegeCount = 1
        tp.Attributes = SE_PRIVILEGE_ENABLED
        
        ' enables or disables privileges in the specified access token
        ' last 3 arguments are zero because we aren't interested
        ' in previous privilege attributes.
        AdjustTokenPrivileges hToken, False, tp, 0, ByVal 0&, ByVal 0&
    End If
    
    ' prepare shutdown flags
    flags = EWX_SHUTDOWN
    If Reboot Then flags = flags Or EWX_REBOOT
    If Force Then flags = flags Or EWX_FORCE
    
    ' finally, you can shut down Windows
    ExitWindowsEx flags, &HFFFF
    
End Sub


