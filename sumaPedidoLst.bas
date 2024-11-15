Option Compare Database
Option Explicit



Public Sub totalPedido()
On Error Resume Next
Dim i As Integer
Dim suma As Double
Dim colNumero As Integer

    
If Form_frm_Pedidos.ListaPedido.ListCount >= 1 Then i = 0

    For i = i To Form_frm_Pedidos.ListaPedido.ListCount - 1
        If Form_frm_Pedidos.ListaPedido.Column(2, i) = 0 Then
        'suma = Nz(suma, 0) + Form_frm_Pedidos.ListaPedido.column(2, i)
        Else
        suma = Nz(suma, 0) + Form_frm_Pedidos.ListaPedido.Column(2, i)
        'End If
        End If
        Next i

Form_frm_Pedidos.txtSubTotal = suma
Form_frm_Pedidos.txt_totallst.Value = suma - Form_frm_Pedidos.txtTotDesc

    'totalPedido = sum
'End If
End Sub

Function BorrarDetalleDescuento(txtBox As TextBox, palabra As String)
    Dim lineas As Variant
    Dim nuevaTexto As String
    Dim linea As Variant
    Dim descuento As Long
    ' Verifica si el TextBox est� vac�o
    If IsNull(txtBox.Value) Or txtBox.Value = "" Then
        'MsgBox "El campo est� vac�o.", vbExclamation, "Aviso"
        Exit Function
    End If
If palabra = "total" Then

    lineas = Split(txtBox.Value, vbCrLf)
    
    ' Recorre cada l�nea
    For Each linea In lineas
        ' Si la l�nea no contiene la palabra espec�fica de forma exacta, la conserva
        If InStr(1, linea, palabra, vbTextCompare) = 0 Then
            nuevaTexto = nuevaTexto & linea & vbCrLf
        Else
        'Dim i As Integer
        'Dim suma As Double
        'suma = 0
        
        ' Recorre cada fila del ListBox
        'For i = 0 To Form_frm_Pedidos.ListaPedido.ListCount - 1
        ' Suma el valor de la columna 2 (�ndice 1) en cada fila
        'suma = suma + CDbl(Form_frm_Pedidos.ListaPedido.Column(2, i))
        'Next i
          'If Form_frm_Pedidos.txtSumaDedescuentosXProductos = "" Or Form_frm_Pedidos.txtSumaDedescuentosXProductos = 0 Or IsNull(Form_frm_Pedidos.txtSumaDedescuentosXProductos) Then
            descuento = ((ExtraerNumerosHastaPorcentaje(txtBox) / 100) * Form_frm_Pedidos.txtSubTotal) '(Form_frm_Pedidos.txt_totallst + Form_frm_Pedidos.txtTotDesc))
          'Else
            'descuento = ((ExtraerNumerosHastaPorcentaje(txtBox) / 100) * (Form_frm_Pedidos.txt_totallst + Form_frm_Pedidos.txtTotDesc)) '((Form_frm_Pedidos.txtSubTotal - Form_frm_Pedidos.txtSumaDedescuentosXProductos)))
            
            'Else
        'End If
             Form_frm_Pedidos.txtTotDesc = Form_frm_Pedidos.txtTotDesc - descuento
            ' Si la l�nea contiene exactamente la palabra, la elimina por completo
            linea = Replace(linea, palabra & vbCrLf, "")
        End If
    Next linea
    
    ' Quita el �ltimo salto de l�nea adicional
    If Right(nuevaTexto, 2) = vbCrLf Then
        nuevaTexto = Left(nuevaTexto, Len(nuevaTexto) - 2)
    End If
    
    ' Asigna el texto modificado al TextBox
    txtBox.Value = nuevaTexto

Else
    ' Divide el contenido del TextBox en l�neas, usando salto de l�nea como delimitador
    lineas = Split(txtBox.Value, vbCrLf)
    
    ' Recorre cada l�nea
    For Each linea In lineas
        ' Si la l�nea no contiene la palabra espec�fica de forma exacta, la conserva
        If InStr(1, linea, palabra, vbTextCompare) = 0 Then
            nuevaTexto = nuevaTexto & linea & vbCrLf
        Else
        
            descuento = (ExtraerNumerosHastaPorcentaje(txtBox) / 100) * Form_frm_Pedidos.ListaPedido.Column(2)
            Form_frm_Pedidos.txtTotDesc = Form_frm_Pedidos.txtTotDesc - descuento
            'If Form_frm_Pedidos.txtSumaDedescuentosXProductos > 0 Or Form_frm_Pedidos.txtSumaDedescuentosXProductos <> "" Or Not IsNull(Form_frm_Pedidos.txtSumaDedescuentosXProductos) Then
            ' Si la l�nea contiene exactamente la palabra, la elimina por completo
            Form_frm_Pedidos.txtSumaDedescuentosXProductos = Form_frm_Pedidos.txtSumaDedescuentosXProductos - descuento  'Form_frm_Pedidos.txtSumaDedescuentosXProductos = Form_frm_Pedidos.txtTotDesc
            'End If
            linea = Replace(linea, palabra & vbCrLf, "")
        End If
    Next linea
    
    ' Quita el �ltimo salto de l�nea adicional
    If Right(nuevaTexto, 2) = vbCrLf Then
        nuevaTexto = Left(nuevaTexto, Len(nuevaTexto) - 2)
    End If
    
    ' Asigna el texto modificado al TextBox
    txtBox.Value = nuevaTexto
End If
End Function

Function ExtraerNumerosHastaPorcentaje(txtBox As TextBox) As Integer
    Dim texto As String
    Dim numero As String
    Dim i As Integer
    
    ' Inicializa la variable
    numero = ""
    
    ' Verifica si el TextBox tiene alg�n valor
'    If IsNull(txtBox.Value) Or txtBox.Value = "" Then
'        'MsgBox "El campo est� vac�o.", vbExclamation, "Aviso"
'        Exit Function
'    End If
    
    ' Asigna el contenido del TextBox a la variable texto
    texto = txtBox.Value
    
    ' Recorre cada car�cter de la cadena hasta encontrar "%"
    For i = 1 To Len(texto)
        Dim caracter As String
        caracter = Mid(texto, i, 1)
        
        ' Verifica si el car�cter es "%" y sale del bucle si lo encuentra
        If caracter = "%" Then
            Exit For
        End If
        
        ' Agrega los caracteres num�ricos a la variable n�mero
        If IsNumeric(caracter) Then
            numero = numero & caracter
        End If
    Next i
    
    ' Convierte el n�mero a Integer y lo retorna
    If numero <> "" Then
        ExtraerNumerosHastaPorcentaje = CInt(numero)
    Else
        'MsgBox "No se encontraron n�meros antes del s�mbolo %.", vbExclamation, "Aviso"
        'Exit Function
    End If
End Function

Public Sub totalUltimoPedido()
On Error Resume Next
Dim ii As Integer
Dim sum As Double
Dim colNum As Integer

    
If Form_frm_ModuloClientes.txt_Descripcion.ListCount >= 1 Then ii = 0

    For ii = ii To Form_frm_ModuloClientes.txt_Descripcion.ListCount - 1
        If Form_frm_ModuloClientes.txt_Descripcion.Column(2, ii) = 0 Then
        'suma = Nz(suma, 0) + Form_frm_Pedidos.ListaPedido.column(2, i)
        Else
        sum = Nz(sum, 0) + Form_frm_ModuloClientes.txt_Descripcion.Column(2, ii)
        'End If
        End If
        Next ii


Form_frm_ModuloClientes.txt_Total = sum - Form_frm_ModuloClientes.txtdesc

    'totalPedido = sum
'End If
End Sub