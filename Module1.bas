Option Compare Database
Option Explicit
Public cargado As Boolean
 
'Sub cuentaProducto(nomBtn As Control)
'Dim conexion As New ADODB.Connection
'Set conexion = CurrentProject.Connection
''Dim sqlCuenta, precio As String
'Dim instruccion As String
'Dim txtArticulo As Control
''Dim i As Integer
'
'  'If boton.Name = nomBtn Then
'
'     instruccion = "SELECT nombre ,precioUnidad FROM Productos WHERE nombre = '" & nomBtn.Caption & "';"
'     Dim memo As New ADODB.Recordset
'     memo.Open instruccion, conexion
'
'  'End If
'
'        If nomBtn.Name = "btn_combo1" Then
'           Set txtArticulo = Form_frm_Pedidos.txt_cantidad
'
'          If txtArticulo.Value = 0 Or txtArticulo.Value = "" Then
'              txtArticulo.Value = 1
'              Form_frm_Pedidos.textoprecio.Value = memo!PrecioUnidad
'          Else
'              txtArticulo.Value = txtArticulo.Value + 1
'          End If
'          If txtArticulo.Value > 1 Then
'              'For i = 2 To txtArticulo.Value Step 1
'                  Form_frm_Pedidos.textoprecio.Value = Form_frm_Pedidos.textoprecio.Value + memo!PrecioUnidad
'             ' Next i
'        End If
'        End If
'        If nomBtn.Name = "btn_combo2" Then
'           Set txtArticulo = Form_frm_Pedidos.txt_cantidad1
'
'          If txtArticulo.Value = 0 Or txtArticulo.Value = "" Then
'              txtArticulo.Value = 1
'              Form_frm_Pedidos.textoprecio.Value = memo!PrecioUnidad
'          Else
'              txtArticulo.Value = txtArticulo.Value + 1
'          End If
'          If txtArticulo.Value > 1 Then
'             ' For i = 2 To txtArticulo.Value Step 1
'                 Form_frm_Pedidos.textoprecio.Value = Form_frm_Pedidos.textoprecio.Value + memo!PrecioUnidad
'              'Next i
'        End If
'        End If
'        If nomBtn.Name = "btn_combo3" Then
'        Set txtArticulo = Form_frm_Pedidos.txt_cantidad2
'
'          If txtArticulo.Value = 0 Or txtArticulo.Value = "" Then
'              txtArticulo.Value = 1
'              Form_frm_Pedidos.textoprecio.Value = memo!PrecioUnidad
'          Else
'              txtArticulo.Value = txtArticulo.Value + 1
'          End If
'          If txtArticulo.Value > 1 Then
'             ' For i = 2 To txtArticulo.Value Step 1
'                  Form_frm_Pedidos.textoprecio.Value = Form_frm_Pedidos.textoprecio.Value + memo!PrecioUnidad
'              'Next i
'        End If
'        End If
'        If nomBtn.Name = "btn_combo4" Then
'        Set txtArticulo = Form_frm_Pedidos.txt_cantidad3
'
'          If txtArticulo.Value = 0 Or txtArticulo.Value = "" Then
'              txtArticulo.Value = 1
'              Form_frm_Pedidos.textoprecio.Value = memo!PrecioUnidad
'          Else
'              txtArticulo.Value = txtArticulo.Value + 1
'          End If
'          If txtArticulo.Value > 1 Then
'              'For i = 2 To txtArticulo.Value Step 1
'                  Form_frm_Pedidos.textoprecio.Value = Form_frm_Pedidos.textoprecio.Value + memo!PrecioUnidad
'             ' Next i
'        End If
'        End If
'
'
'
'Form_frm_Pedidos.textoarticulo.SetFocus
'Form_frm_Pedidos.textoarticulo.Value = memo!Nombre
'
'memo.Close
'Set memo = Nothing
'conexion.Close
'Set conexion = Nothing
'
'End Sub
Sub cuentaProducto(nomArt As String)
Dim conexion As New ADODB.Connection
Set conexion = CurrentProject.Connection
'Dim sqlCuenta, precio As String
Dim instruccion As String
Dim txtArticulo As Control
Dim Cantidad As Long
Dim cantidadBD As Long
  'If boton.Name = nomBtn Then

     instruccion = "SELECT nombre ,stockActual, precioUnidad FROM Productos WHERE nombre = '" & nomArt & "';"
     Dim memo As New ADODB.Recordset
     memo.Open instruccion, conexion
Form_frm_Pedidos.txtBDStock.Value = memo!stockActual
cantidadBD = Form_frm_Pedidos.txtBDStock.Value
  Set txtArticulo = Form_frm_Pedidos.txtCantidadProducto
'  If Form_frm_Pedidos.txtCantidadProducto = "" Or IsNull(Form_frm_Pedidos.txtCantidadProducto) Then
'       Form_frm_Pedidos.txtCantidadProducto=
If txtArticulo.Value = "" Or IsNull(txtArticulo.Value) Or txtArticulo.Value = 0 Then
    'Form_frm_Pedidos.txtCantidadProducto.Value = 1
    txtArticulo.Value = 1
    End If
'cantidad = txtArticulo.Value
        If txtArticulo.Value = 0 Or IsNull(txtArticulo.Value) And ((cantidadBD - txtArticulo.Value) >= 0) Then
           
           Form_frm_Pedidos.txtCantidadProducto.Value = 1
             
            'txtArticulo.Value = 1
             
                 'If txtArticulo.Value = 0 Or txtArticulo.Value = "" Then
              'txtArticulo.Value = 1
              Form_frm_Pedidos.textoprecio.Value = memo!PrecioUnidad
        ElseIf (cantidadBD - txtArticulo.Value) < 0 Then
             'MsgBox "NO HAY STOCK PARA ESTE PRODUCTO", vbCritical, "BROMOPLAST"
         Form_frm_Pedidos.txtCantTblPrStock.Value = "N"
         Exit Sub
        End If
          If txtArticulo.Value = 1 And (cantidadBD - txtArticulo.Value) >= 0 Then
             
            'txtArticulo.Value = Form_frm_Pedidos.txtCantidadProducto.Value
            Form_frm_Pedidos.textoprecio.Value = memo!PrecioUnidad
          ElseIf (cantidadBD - txtArticulo.Value) < 0 Then
             'MsgBox "NO HAY STOCK PARA ESTE PRODUCTO", vbCritical, "BROMOPLAST"
             Form_frm_Pedidos.txtCantTblPrStock.Value = "N"
             Exit Sub
          End If
      
          
          If txtArticulo.Value > 1 And (cantidadBD - txtArticulo.Value) >= 0 Then
              'For i = 2 To txtArticulo.Value Step 1
             
            'txtArticulo.Value = Form_frm_Pedidos.txtCantidadProducto.Value
                  Form_frm_Pedidos.textoprecio.Value = Form_frm_Pedidos.txtCantidadProducto.Value * memo!PrecioUnidad
             ' Next i
          ElseIf (cantidadBD - txtArticulo.Value) < 0 Then
             'MsgBox "NO HAY STOCK PARA ESTE PRODUCTO", vbCritical, "BROMOPLAST"
             Form_frm_Pedidos.txtCantTblPrStock.Value = "N"
             Exit Sub
             
        End If
        


Form_frm_Pedidos.textoarticulo.SetFocus
Form_frm_Pedidos.textoarticulo.Value = memo!Nombre

memo.Close
Set memo = Nothing
conexion.Close
Set conexion = Nothing

End Sub

Sub RowsSelected()

 Dim ctlList As Control, varItem As Variant
 
 ' Return Control object variable pointing to list box.
 'Set ctlList = Form! Employees!EmployeeList
 Set ctlList = Form_frm_Pedidos.ListaPedido
 ' Enumerate through selected items.
 For Each varItem In ctlList.ItemsSelected
 ' Print value of bound column.
Debug.Print ctlList.ItemData(varItem)
MsgBox ctlList.ItemData(varItem)
 Next varItem
End Sub

Sub transferInfo()
'Debug.Print Form_frm_Pedidos.ListaPedido.RowSource
'Form_frm_Pedidos.detPedido.Value = Form_frm_Pedidos.ListaPedido.RowSource
Dim cont, cont2 As Integer, dataSTR, char As String

dataSTR = Form_frm_Pedidos.ListaPedido.RowSource
'Debug.Print dataSTR

 For cont = 1 To Len(dataSTR) Step 1
    Debug.Print Mid(dataSTR, cont, 1)
    If Mid(dataSTR, cont, 1) = ";" Then
    cont2 = cont2 + 1
    End If
    If Mid(dataSTR, cont, 1) = ";" And cont2 Mod 2 = 0 Then
    Mid(dataSTR, cont, 1) = "x"
    End If
    If Mid(dataSTR, cont, 1) = "x" Then
    dataSTR = Replace(dataSTR, "x", "<br>")
    
    End If
    
 Next cont
Form_frm_Pedidos.ListaPedido.RowSource = dataSTR
Form_frm_Pedidos.detPedido.Value = Form_frm_Pedidos.ListaPedido.RowSource
Form_frm_Pedidos.ListaPedido.RowSource = ""



End Sub
Sub EnviarDatosAMySQL()
    Dim coneccion As ADODB.Connection
    Dim sqlInsert As String
    Dim rs As ADODB.Recordset
    Dim maxID As Integer
    Dim Nombre As String
    Dim Descripcion As String
    Dim listIndex, i As Integer
    Dim totalCantidadListBox As Integer
    Dim cantidadActual As Integer
    
    ' Inicializar el total de la cantidad en el ListBox
    totalCantidadListBox = 0
    
    ' Configurar la conexión
    Set coneccion = New ADODB.Connection
    coneccion.ConnectionString = "Driver={MySQL ODBC 8.4 Unicode Driver};Server=localhost;Database=test3;User=root;Option=3;"
    coneccion.Open
    
    ' Obtener el valor máximo actual de IDAuto en la tabla MySQL
    Set rs = New ADODB.Recordset
    rs.Open "SELECT MAX(IDAuto) AS maxID FROM orden", coneccion, adOpenStatic, adLockReadOnly
    If Not rs.EOF Then
        If IsNull(rs!maxID) Then
            maxID = 0
        Else
            maxID = rs!maxID
        End If
    Else
        maxID = 0
    End If
    rs.Close
    
    ' Calcular el total de la cantidad en el ListBox
    For listIndex = 0 To Form_frm_Clientes!frm_Pedidos.Form!ListaPedido.ListCount - 1       ' Form_Form1.List2.ListCount - 1
        totalCantidadListBox = totalCantidadListBox + Form_frm_Clientes!frm_Pedidos.Form!ListaPedido.Column(0, listIndex)
    Next listIndex
    
    ' Inicializar la cantidad actual
    cantidadActual = 1
    
    ' Iterar a través de todos los elementos del ListBox (sin importar si están seleccionados)
    For listIndex = 0 To Form_frm_Clientes!frm_Pedidos.Form!ListaPedido.ListCount - 1
        ' Extraer los datos del ListBox
        Nombre = Form_frm_Clientes.txtNombre.Value '& " " & Form_frm_ClientesPedido.txt_apellido_cliente.Value
        Descripcion = Form_frm_Clientes!frm_Pedidos.Form!ListaPedido.Column(1, listIndex)
        
        ' Insertar una fila por cada cantidad de la fila del ListBox
        For i = 1 To Form_frm_Clientes!frm_Pedidos.Form!ListaPedido.Column(0, listIndex)
            maxID = maxID + 1 ' Incrementar el valor de IDAuto
            sqlInsert = "INSERT INTO orden (IDAuto, cantidad, nomCliente, nomProd) VALUES (" & maxID & ", '" & cantidadActual & " de " & totalCantidadListBox & "', '" & Nombre & "', '" & Descripcion & "');"
            coneccion.Execute sqlInsert
            cantidadActual = cantidadActual + 1
        Next i
    Next listIndex
    
    ' Cerrar la conexión
    coneccion.Close
    Set coneccion = Nothing
    
    MsgBox "Datos enviados a MySQL correctamente."
End Sub


'Sub CargaProduct()
'
'Dim con As New ADODB.Connection
'Set con = CurrentProject.Connection
'Dim inst As String
'Dim mem As New ADODB.Recordset
'inst = "SELECT IDProd, nombre FROM Productos WHERE IDProd IN(1, 2, 3, 4)"
'mem.Open inst, con, adOpenStatic, adLockReadOnly
'' Asignar los valores a los botones
'    Do While Not mem.EOF
'        Select Case mem("IDProd").Value
'            Case 1
'                Form_frm_Pedidos.btn_combo1.Caption = mem("Nombre").Value
'            Case 2
'                Form_frm_Pedidos.btn_combo2.Caption = mem("Nombre").Value
'            Case 3
'                Form_frm_Pedidos.btn_combo3.Caption = mem("Nombre").Value
'            Case 4
'                Form_frm_Pedidos.btn_combo4.Caption = mem("Nombre").Value
'        End Select
'        mem.MoveNext
'    Loop
'
'    ' Cerrar el Recordset y la conexión
'    mem.Close
'    con.Close
'
'    ' Limpiar objetos
'    Set mem = Nothing
'    Set con = Nothing
'
'End Sub


Sub PedidoR()
Dim bd As DAO.Database
Dim rs As DAO.Recordset
Dim sql As String
Dim clienteID, idPedido As Integer
Report_Pedido.Cliente.Value = Form_frm_Clientes!frm_Pedidos.Form.txt_IDCli
'Report_Pedido.Lista.RowSource = Form_frm_Clientes!frm_Pedidos.Form.ListaPedido.RowSource
'Report_Pedido.Lista.Requery
'Report_Pedido.txtIdPedido.Value = Form_frm_Clientes!frm_Pedidos.Form.txt_IdPedido
    clienteID = Report_Pedido.Cliente.Value
    sql = "SELECT * FROM Clientes WHERE Id_Cliente = " & clienteID
    Set bd = CurrentDb
    Set rs = bd.OpenRecordset(sql, dbOpenDynaset)
    Report_Pedido.txtDetalleDesc.Value = Form_frm_Clientes!frm_Pedidos.Form.txtDetalleDesc.Value
    Report_Pedido.Nombre.Value = rs!Nombre
    Report_Pedido.Direccion.Value = rs!Direccion
    Report_Pedido.Telefono.Value = rs!Telefono
    Report_Pedido.e_mail.Value = rs!e_mail
    Report_Pedido.rut.Value = rs!rut
    Report_Pedido.ciudad.Value = rs!ciudad

rs.Close
Set rs = Nothing
Set bd = Nothing
    idPedido = Form_frm_Clientes!frm_Pedidos.Form.txt_IdPedido.Value
    'idPedido = Report_Pedido.txtIdPedido.Value
    sql = "SELECT * FROM Pedido WHERE ID_Pedido = " & idPedido
    Set bd = CurrentDb
    Set rs = bd.OpenRecordset(sql, dbOpenDynaset)
    Report_Pedido.fecha.Value = Format(rs!fecha, "dd/mm/yyyy")
     Report_Pedido.Factura.Value = rs!Factura
     Report_Pedido.txtdesc = rs!Descuentos
    Report_Pedido.Lista.RowSource = Form_frm_Clientes!frm_Pedidos.Form!ListaPedido.RowSource 'rs!Descripcion
    Report_Pedido.Lista.Requery
    
On Error Resume Next
Dim i As Integer
Dim suma As Double
Dim colNumero As Integer

    
If Report_Pedido.Lista.ListCount >= 1 Then i = 0

    For i = i To Report_Pedido.Lista.ListCount - 1
        If Report_Pedido.Lista.Column(2, i) = 0 Then
        'suma = Nz(suma, 0) + Form_frm_Pedidos.ListaPedido.column(2, i)
        Else
        suma = Nz(suma, 0) + Report_Pedido.Lista.Column(2, i)
        'End If
        End If
        Next i

'Report_Pedido.Total = suma
Report_Pedido.Total = suma - Report_Pedido.txtdesc
    
  rs.Close
Set rs = Nothing
Set bd = Nothing
 
End Sub

Sub VerifProve()
Dim bd As DAO.Database
    Dim rst As DAO.Recordset
    Dim strSQLP As String
    Dim Nom, rs, Tel As String
    
    'ProveID = Form_frm_Productos.cmbProvee.Value
Nom = Form_frm_Proveedores.txtNombre.Value
rs = Form_frm_Proveedores.txtRS.Value

Tel = Form_frm_Proveedores.txtTel.Value


'If IsNull(ProveID) Then
    strSQLP = "SELECT Id_proveedor FROM Proveedores WHERE Nombre = '" & Nom & "' AND RazonSocial = '" & rs & "' AND Telefono = " & Tel & " "
'     ' Establecer la base de datos actual
    Set bd = CurrentDb
    ' Abrir el Recordset con la consulta
    Set rst = bd.OpenRecordset(strSQLP, dbOpenDynaset)
'    ' Verificar si se encontró algún registro
'    'If Form_frm_Productos.txtNomProd = "" Then
If rst.EOF Then
     Form_frm_Proveedores.txtIdProv.Value = Null
     
Else
 'MsgBox "Proveedor ya existe"
     Form_frm_Proveedores.txtIdProv.Value = rst!Id_proveedor
''    Else
''        MsgBox "No se encontró el cliente con el ID proporcionado.", vbExclamation
    End If
    
    rst.Close
    Set rst = Nothing
    Set bd = Nothing
    
End Sub
Function cerra()
DoCmd.Close acForm, "frm_ControlCaja", acSaveNo
End Function