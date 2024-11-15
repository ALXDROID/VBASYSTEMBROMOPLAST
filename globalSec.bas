Option Compare Database
Option Explicit
Public StopSleep As Boolean
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal Milliseconds As LongPtr)

Public Sub SleepSec(NumSec As Long, Optional AllowEvents As Boolean = True)
    Dim X As Integer
    For X = 1 To NumSec
        Sleep 1000
        If AllowEvents Then DoEvents
        If StopSleep Then Exit Sub
    Next
End Sub
'Sub routine para generar 1 sec(1000)en milisegundos,
'StopSleep es una variable global para comenzar o detener el contador SleepSec