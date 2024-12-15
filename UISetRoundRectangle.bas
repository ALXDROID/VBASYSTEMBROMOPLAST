
Option Compare Database
Option Explicit

Private Declare PtrSafe Function CreateRoundRectRgn Lib "gdi32" ( _
        ByVal nLeftRect As Long, ByVal nTopRect As Long, ByVal nRightRect As Long, _
        ByVal nBottomRect As Long, ByVal nWidthEllipse As Long, ByVal nHeightEllipse As Long) As LongPtr

Private Declare PtrSafe Function SetWindowRgn Lib "user32" ( _
        ByVal hWnd As LongPtr, ByVal hRgn As LongPtr, ByVal bRedraw As Boolean) As Long

Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
        ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" ( _
        ByVal hWnd As LongPtr, ByVal nIndex As Long) As Long

Private Const GWL_STYLE As Long = -16
Private Const WS_BORDER As Long = &H800000
Private Const WS_DLGFRAME As Long = &H400000
Private Const WS_CAPTION As Long = &HC00000

Dim hRgn As LongPtr

Public Function UISetRoundRect( _
       ByVal UIForm As Form, _
       ByVal CornersInPixels As Integer, _
       Optional ByVal TopCornersOnly As Boolean = True) As Boolean
       
    Dim intRight As Integer
    Dim intHeight As Integer

    ' Se obtiene el tamaño del formulario en píxeles
    With UIForm
        intRight = PixelsPerTwipsX(.WindowWidth)
        intHeight = PixelsPerTwipsY(.WindowHeight)

        ' Si solo se redondean las esquinas superiores
        If TopCornersOnly Then
            intHeight = intHeight + CornersInPixels
        Else
            intHeight = intHeight + 1
'            intRight = intRight + 1
        End If

        ' Eliminar los bordes del formulario
        Dim hWnd As LongPtr
        hWnd = .hWnd

        ' Obtiene el estilo actual de la ventana
        Dim currentStyle As Long
        currentStyle = GetWindowLong(hWnd, GWL_STYLE)
        
        ' Elimina los bordes de la ventana y la barra de título
'        SetWindowLong hWnd, GWL_STYLE, currentStyle And Not (WS_BORDER Or WS_DLGFRAME Or WS_CAPTION)
SetWindowLong hWnd, GWL_STYLE, currentStyle   'And Not (WS_BORDER Or WS_DLGFRAME Or WS_CAPTION)

        ' Crea el borde redondeado
        hRgn = CreateRoundRectRgn(0, 0, intRight, intHeight, CornersInPixels, CornersInPixels)

        ' Aplica el borde redondeado
        SetWindowRgn hWnd, hRgn, True
    End With
End Function





'Option Compare Database
'Option Explicit
'
'Private Declare PtrSafe Function CreateRoundRectRgn Lib "gdi32" ( _
'        ByVal nLeftRect As Long, ByVal nTopRect As Long, ByVal nRightRect As Long, _
'        ByVal nBottomRect As Long, ByVal nWidthEllipse As Long, ByVal nHeightEllipse As Long) As LongPtr
'
'Private Declare PtrSafe Function SetWindowRgn Lib "user32" ( _
'        ByVal hWnd As LongPtr, ByVal hRgn As LongPtr, ByVal bRedraw As Boolean) As Long
'
'Dim hRgn As LongPtr
'
'Public Function UISetRoundRect( _
'       ByVal UIForm As Form, _
'       ByVal CornersInPixels As Integer, _
'       Optional ByVal TopCornersOnly As Boolean = True) As Boolean
'
''CR - changed CornersInPixels from Byte to Integer to fix overflow error
'    Dim intRight As Integer
'    Dim intHeight As Integer
'    With UIForm
'
'        intRight = PixelsPerTwipsX(.WindowWidth)
'        intHeight = PixelsPerTwipsY(.WindowHeight)
'
'        If TopCornersOnly Then
'            intHeight = intHeight + CornersInPixels
'        Else
'            intHeight = intHeight + 5 '+ 1
'            intRight = intRight + 5
'        End If
'
'        hRgn = CreateRoundRectRgn(0, 0, intRight, intHeight, CornersInPixels, CornersInPixels)
'
'        SetWindowRgn .hWnd, hRgn, False
'    End With
'End Function
'