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

Function BorrarDetalleDescuento(TxtBox As TextBox, palabra As String)
    Dim lineas As Variant
    Dim nuevaTexto As String
    Dim linea As Variant
    Dim descuento As Long
    ' Verifica si el TextBox está vacío
    If IsNull(TxtBox.Value) Or TxtBox.Value = "" Then
        'MsgBox "El campo está vacío.", vbExclamation, "Aviso"
        Exit Function
    End If
If palabra = "total" Then

    lineas = Split(TxtBox.Value, vbCrLf)
    
    ' Recorre cada línea
    For Each linea In lineas
        ' Si la línea no contiene la palabra específica de forma exacta, la conserva
        If InStr(1, linea, palabra, vbTextCompare) = 0 Then
            nuevaTexto = nuevaTexto & linea & vbCrLf
        Else
        'Dim i As Integer
        'Dim suma As Double
        'suma = 0
        
        ' Recorre cada fila del ListBox
        'For i = 0 To Form_frm_Pedidos.ListaPedido.ListCount - 1
        ' Suma el valor de la columna 2 (índice 1) en cada fila
        'suma = suma + CDbl(Form_frm_Pedidos.ListaPedido.Column(2, i))
        'Next i
          'If Form_frm_Pedidos.txtSumaDedescuentosXProductos = "" Or Form_frm_Pedidos.txtSumaDedescuentosXProductos = 0 Or IsNull(Form_frm_Pedidos.txtSumaDedescuentosXProductos) Then
            descuento = ((ExtraerNumerosHastaPorcentaje(TxtBox) / 100) * Form_frm_Pedidos.txtSubTotal) '(Form_frm_Pedidos.txt_totallst + Form_frm_Pedidos.txtTotDesc))
          'Else
            'descuento = ((ExtraerNumerosHastaPorcentaje(txtBox) / 100) * (Form_frm_Pedidos.txt_totallst + Form_frm_Pedidos.txtTotDesc)) '((Form_frm_Pedidos.txtSubTotal - Form_frm_Pedidos.txtSumaDedescuentosXProductos)))
            
            'Else
        'End If
             Form_frm_Pedidos.txtTotDesc = Form_frm_Pedidos.txtTotDesc - descuento
            ' Si la línea contiene exactamente la palabra, la elimina por completo
            linea = Replace(linea, palabra & vbCrLf, "")
        End If
    Next linea
    
    ' Quita el último salto de línea adicional
    If Right(nuevaTexto, 2) = vbCrLf Then
        nuevaTexto = Left(nuevaTexto, Len(nuevaTexto) - 2)
    End If
    
    ' Asigna el texto modificado al TextBox
    TxtBox.Value = nuevaTexto

Else
    ' Divide el contenido del TextBox en líneas, usando salto de línea como delimitador
    lineas = Split(TxtBox.Value, vbCrLf)
    
    ' Recorre cada línea
    For Each linea In lineas
        ' Si la línea no contiene la palabra específica de forma exacta, la conserva
        If InStr(1, linea, palabra, vbTextCompare) = 0 Then
            nuevaTexto = nuevaTexto & linea & vbCrLf
        Else
        
            descuento = (ExtraerNumerosHastaPorcentaje(TxtBox) / 100) * Form_frm_Pedidos.ListaPedido.Column(2)
            Form_frm_Pedidos.txtTotDesc = Form_frm_Pedidos.txtTotDesc - descuento
            'If Form_frm_Pedidos.txtSumaDedescuentosXProductos > 0 Or Form_frm_Pedidos.txtSumaDedescuentosXProductos <> "" Or Not IsNull(Form_frm_Pedidos.txtSumaDedescuentosXProductos) Then
            ' Si la línea contiene exactamente la palabra, la elimina por completo
            Form_frm_Pedidos.txtSumaDedescuentosXProductos = Form_frm_Pedidos.txtSumaDedescuentosXProductos - descuento  'Form_frm_Pedidos.txtSumaDedescuentosXProductos = Form_frm_Pedidos.txtTotDesc
            'End If
            linea = Replace(linea, palabra & vbCrLf, "")
        End If
    Next linea
    
    ' Quita el último salto de línea adicional
    If Right(nuevaTexto, 2) = vbCrLf Then
        nuevaTexto = Left(nuevaTexto, Len(nuevaTexto) - 2)
    End If
    
    ' Asigna el texto modificado al TextBox
    TxtBox.Value = nuevaTexto
End If
End Function

Function ExtraerNumerosHastaPorcentaje(TxtBox As TextBox) As Integer
    Dim Texto As String
    Dim numero As String
    Dim i As Integer
    
    ' Inicializa la variable
    numero = ""
    
    ' Verifica si el TextBox tiene algún valor
'    If IsNull(txtBox.Value) Or txtBox.Value = "" Then
'        'MsgBox "El campo está vacío.", vbExclamation, "Aviso"
'        Exit Function
'    End If
    
    ' Asigna el contenido del TextBox a la variable texto
    Texto = TxtBox.Value
    
    ' Recorre cada carácter de la cadena hasta encontrar "%"
    For i = 1 To Len(Texto)
        Dim caracter As String
        caracter = Mid(Texto, i, 1)
        
        ' Verifica si el carácter es "%" y sale del bucle si lo encuentra
        If caracter = "%" Then
            Exit For
        End If
        
        ' Agrega los caracteres numéricos a la variable número
        If IsNumeric(caracter) Then
            numero = numero & caracter
        End If
    Next i
    
    ' Convierte el número a Integer y lo retorna
    If numero <> "" Then
        ExtraerNumerosHastaPorcentaje = CInt(numero)
    Else
        'MsgBox "No se encontraron números antes del símbolo %.", vbExclamation, "Aviso"
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

Sub restastock()

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim datos As String
    Dim filas() As String
    Dim columnas() As String
    Dim i As Integer
    Dim producto As String
    Dim valorCambio As Long
    Dim sql As String

    ' Obtener el contenido del campo Long Text (RowSource)
    datos = Form_frm_Pedidos.ListaPedido.RowSource ' Reemplaza con el control o fuente de texto

    ' Dividir las filas (asume salto de línea como delimitador)
    filas = Split(datos, vbCrLf)

    ' Abrir la tabla Productos
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT nombre, stockActual FROM Productos")

    ' Recorrer cada producto de la tabla
    Do While Not rs.EOF
        producto = rs!Nombre ' Nombre del producto actual
        valorCambio = 0 ' Inicializar valor de cambio

        ' Recorrer cada fila del RowSource
        For i = LBound(filas) To UBound(filas)
            ' Dividir columnas de la fila (delimitador ';')
            columnas = Split(filas(i), ";")

            ' Verificar si el producto coincide con la primera columna
            If UBound(columnas) >= 1 And columnas(1) = producto Then
                ' Sumar o restar el valor de la columna 2 (índice 1)
                valorCambio = valorCambio + CLng(columnas(0))
            End If
        Next i

        ' Actualizar el valor en la tabla Productos
        If valorCambio <> 0 And columnas(1) = producto Then
            sql = "UPDATE Productos SET stockActual = stockActual - " & valorCambio & " WHERE nombre = '" & producto & "';"
            DoCmd.RunSQL sql
        End If

        ' Mover al siguiente registro
        rs.MoveNext
    Loop

    ' Cerrar el Recordset
    rs.Close
    Set rs = Nothing
    Set db = Nothing

    MsgBox "Stock actualizado correctamente.", vbInformation


    
End Sub