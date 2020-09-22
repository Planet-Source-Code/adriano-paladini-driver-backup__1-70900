Attribute VB_Name = "md_shell"
Option Explicit

'# use to shell process #
Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
'# use to shell process #

Public Sub ShellAndWait(ByVal strProg As String, ByVal lStyle As VbAppWinStyle)
Dim ProcessId As Long
Dim ProcessHandle As Long
Const Access As Long = &H100000
ProcessId = Shell(strProg, lStyle)
Do
    ProcessHandle = OpenProcess(Access, False, ProcessId)
    If ProcessHandle <> 0 Then
        CloseHandle ProcessHandle
    End If
    DoEvents
Loop Until ProcessHandle = 0
End Sub
