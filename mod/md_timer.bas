Attribute VB_Name = "md_timer"
Option Explicit
Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
Private lngTimerID As Long

Public Sub TimerProc(ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long)
DisableTimer
End Sub

Public Sub EnableTimer(ByVal lngInterval As Long)
lngTimerID = SetTimer(0, 0, lngInterval, AddressOf TimerProc)
End Sub

Public Sub DisableTimer()
KillTimer 0, lngTimerID
End Sub
