Option Compare Database
Option Explicit

Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HT_CAPTION = &H2

' 32/64 Bit Windows API calls for VBA 7 (A2010 or later)


Public Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As LongPtr, _
    ByVal wMsg As Long, ByVal wParam As LongPtr, lParam As Any) As LongPtr


Public Declare PtrSafe Function ReleaseCapture Lib "user32.dll" () As Long

'Use the following code in a mouse down event:
'DragFormWindow Me

Public Function DragFormWindow(frm As Form)
    
    With frm
        ReleaseCapture
        SendMessage .hWnd, WM_NCLBUTTONDOWN, HT_CAPTION, 0
    End With
    
End Function