Option Compare Database
Option Explicit

Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hWnd As LongPtr) As LongPtr

Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hWnd As LongPtr, ByVal hDC As LongPtr) As Long

Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hDC As LongPtr, ByVal nIndex As Long) As Long

Private Const LOGPIXELSX = 88
Private Const LOGPIXELSY = 90
Private Const TWIPSPERINCH = 1440

Dim hDC As LongPtr
Dim PixelsPerInchX As Long, PixelsPerInchY As Long, lngPixelsPerInch As Long

Public Function TwipsPerPixelX(PixelWidthX As Long) As Integer

    hDC = GetDC(0)
    PixelsPerInchX = GetDeviceCaps(hDC, LOGPIXELSX)
    TwipsPerPixelX = (PixelWidthX / PixelsPerInchX) * TWIPSPERINCH
    ReleaseDC 0, hDC

End Function
Public Function TwipsPerPixelY(PixelHeightY As Long) As Integer

    hDC = GetDC(0)
    PixelsPerInchY = GetDeviceCaps(hDC, LOGPIXELSY)
    TwipsPerPixelY = (PixelHeightY / PixelsPerInchY) * TWIPSPERINCH
    ReleaseDC 0, hDC
End Function
Public Function PixelsPerTwipsX(TwipsWidthX As Long) As Integer

    hDC = GetDC(0)
    lngPixelsPerInch = GetDeviceCaps(hDC, LOGPIXELSX)
    PixelsPerTwipsX = (TwipsWidthX / TWIPSPERINCH) * lngPixelsPerInch
    ReleaseDC 0, hDC

End Function
Public Function PixelsPerTwipsY(TwipsHeightY As Long) As Integer

    hDC = GetDC(0)
    lngPixelsPerInch = GetDeviceCaps(hDC, LOGPIXELSY)
    PixelsPerTwipsY = (TwipsHeightY / TWIPSPERINCH) * lngPixelsPerInch
    ReleaseDC 0, hDC

End Function