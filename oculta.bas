Option Compare Database
Option Explicit


Const SW_HIDE = 0
Const SW_SHOWNORMAL = 1
Const SW_SHOWMINIMIZED = 2
Const SW_SHOWMAXIMIZED = 3

Private Declare PtrSafe Function apiShowWindow Lib "user32" Alias "ShowWindow" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long

Function fSetAccessWindow(nCmdShow As Long)
    Dim loX As Long
    Dim loForm As Form
   
    loX = apiShowWindow(hWndAccessApp, nCmdShow)
End Function