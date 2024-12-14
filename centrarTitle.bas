Option Compare Database
Option Explicit
Private Declare PtrSafe Function FindWindowA Lib "user32" ( _
    ByVal lpClassName As String, _
    ByVal lpWindowName As String) As LongPtr

Private Declare PtrSafe Function SetWindowTextA Lib "user32" ( _
    ByVal hWnd As LongPtr, _
    ByVal lpString As String) As Long

Private Declare PtrSafe Function GetSysColor Lib "user32" ( _
    ByVal nIndex As Long) As Long

Private Declare PtrSafe Function SetSysColors Lib "user32" ( _
    ByVal nChanges As Long, _
    lpElements As LongPtr, _
    lpRgbValues As LongPtr) As Long

Private Const COLOR_ACTIVECAPTION As Long = 2 ' �ndice del color de la barra de t�tulo activa

Public Sub CambiarFondoBarraTitulo(nuevoColor As Long)
    Dim hWnd As LongPtr
    Dim M As Form
   Set M = Form_MyMsgBox
    ' Obtener el handle de la ventana actual
    hWnd = FindWindowA(vbNullString, M.Caption) ', Me.Caption
    If hWnd = 0 Then Exit Sub

    ' Cambiar el color de la barra de t�tulo
    SetSysColors 1, VarPtr(COLOR_ACTIVECAPTION), VarPtr(nuevoColor)
End Sub
' Declaraciones de API para Windows
'Private Declare PtrSafe Function FindWindowA Lib "user32" ( _
'    ByVal lpClassName As String, _
'    ByVal lpWindowName As String) As LongPtr
'
'Private Declare PtrSafe Function GetWindowRect Lib "user32" ( _
'    ByVal hWnd As LongPtr, _
'    lpRect As RECT) As Long
'
'Private Declare PtrSafe Function GetDC Lib "user32" ( _
'    ByVal hWnd As LongPtr) As LongPtr
'
'Private Declare PtrSafe Function ReleaseDC Lib "user32" ( _
'    ByVal hWnd As LongPtr, _
'    ByVal hDC As LongPtr) As Long
'
'Private Declare PtrSafe Function DrawText Lib "user32" Alias "DrawTextA" ( _
'    ByVal hDC As LongPtr, _
'    ByVal lpStr As String, _
'    ByVal nCount As Long, _
'    lpRect As RECT, _
'    ByVal wFormat As Long) As Long
'
'' Constantes para alineaci�n de texto
'Private Const DT_CENTER As Long = &H1
'Private Const DT_VCENTER As Long = &H4
'Private Const DT_SINGLELINE As Long = &H20
'
'' Definici�n de tipo RECT para almacenar las coordenadas del rect�ngulo
'Private Type RECT
'    Left As Long
'    Top As Long
'    Right As Long
'    Bottom As Long
'End Type
'' Procedimiento para centrar el t�tulo
'Public Sub CentrarCaption()
'    Dim hWnd As LongPtr
'    Dim hDC As LongPtr
'    Dim r As RECT
'
'    ' Obt�n el handle de la ventana actual
'    hWnd = FindWindowA(vbNullString, Form_MyMsgBox.Caption)
'    If hWnd = 0 Then Exit Sub
'
'    ' Obt�n el �rea del formulario
'    GetWindowRect hWnd, r
'
'    ' Obt�n el contexto de dispositivo (DC) para dibujar
'    hDC = GetDC(hWnd)
'    If hDC = 0 Then Exit Sub
'
'    ' Ajusta el rect�ngulo al �rea de la barra de t�tulo
'    r.Top = 0
'    r.Bottom = 30 ' Altura aproximada de la barra de t�tulo
'
'    ' Dibuja el texto centrado usando el valor actual de Me.Caption
'    DrawText hDC, Form_MyMsgBox.Caption, Len(Form_MyMsgBox.Caption), r, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE
'
'    ' Libera el contexto de dispositivo
'    ReleaseDC hWnd, hDC
'
'End Sub