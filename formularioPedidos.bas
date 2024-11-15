Option Compare Database
Option Explicit

Sub ultimoPedido()

    
   Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim strSQL, strSQL1 As String
    Dim clienteID, pedidoID As Integer
    
    ' Obtener el ID del cliente desde un TextBox o de alguna manera
'    clienteID = Form_BuscarClientes.ListaClientes.Column(0)
    clienteID = Form_frm_ModuloClientes.ListaClientes.Column(0)
    
    ' Construir la consulta SQL
    strSQL = "SELECT MAX(ID_Pedido)as maxID FROM Pedido WHERE Cliente = " & clienteID
    
    ' Establecer la base de datos actual
    Set db = CurrentDb
    
    ' Abrir el Recordset con la consulta
    Set rs = db.OpenRecordset(strSQL, dbOpenDynaset)
    If Not IsNull(rs!maxID) Then
    pedidoID = rs!maxID
    
    rs.Close
    Set rs = Nothing
    Set db = Nothing

    strSQL1 = "SELECT  Cliente, Fecha,  Factura, Descripcion, Descuentos FROM Pedido WHERE ID_Pedido = " & pedidoID  'Format(Fecha, 'dd/mm/yyyy') AS FechaFormateada,
     ' Establecer la base de datos actual
    Set db = CurrentDb
    ' Abrir el Recordset con la consulta
    Set rs = db.OpenRecordset(strSQL1, dbOpenDynaset)
        'Form_frm_ModuloClientes.cmbFechaFact.ColumnCount = 1 '2
 'se oculta al establecer su ancho a 0.
    'Form_frm_ModuloClientes.cmbFechaFact.ColumnWidths = "3cm"
    ' Verificar si se encontró algún registro
    If Not rs.EOF Then
            'Form_frm_ModuloClientes.cmbFechaFact.AddItem rs!Fecha & ";" & rs!FechaFormateada
            ' Colocar los valores en los TextBoxes del formulario
            Form_frm_ModuloClientes.txtIDPed.Value = rs!Cliente
            'Form_frm_ModuloClientes.txt_Fecha.Value = rs!Fecha
            Form_frm_ModuloClientes.cmbFechaFact.Value = rs!Fecha
            'Form_frm_ModuloClientes.cmbFechaFact.RowSource = rs!Fecha
'            Form_frm_ModuloClientes.cmbFechaFact.Value = rs!FechaFormateada
           'Form_frm_ModuloClientes.cmbFechaFact.Column(1) = Form_frm_ModuloClientes.cmbFechaFact.Value
            Form_frm_ModuloClientes.txt_Factura.Value = rs!Factura
            Form_frm_ModuloClientes.txt_Descripcion.RowSource = rs!Descripcion
            Form_frm_ModuloClientes.txtdesc = rs!Descuentos
            Call totalUltimoPedido
    Else
        MsgBox "No se encontró el cliente con el ID proporcionado.", vbExclamation
    End If
    'Form_frm_ModuloClientes.cmbFechaFact.Value = Form_frm_ModuloClientes.cmbFechaFact.ItemData(0)
    Form_frm_ModuloClientes.cmbFechaFact.Value = Format(rs!Fecha, "dd/mm/yyyy hh:mm:ss")  'Form_frm_ModuloClientes.cmbFechaFact.Value 'Format(rs!Fecha, Now, "mm/dd/yyyy")
    ' Cerrar el Recordset
       rs.Close
    Set rs = Nothing
    Set db = Nothing

Else

    MsgBox "CLIENTE NO TIENE PEDIDOS", vbInformation, "BROMOPLAST SYSTEM"
     rs.Close
    Set rs = Nothing
    Set db = Nothing
End If

End Sub

Sub IDProveedor()

      Dim bd As DAO.Database
    Dim rst As DAO.Recordset
    Dim strSQL, strSQL2 As String
    Dim ProveID As String
    
'    ProveID = Form_frm_Productos.cmbProvee.Value
ProveID = Form_frm_Inventario.cmbProvee.Value
    ' Construir la consulta SQL
    
    
    strSQL = "SELECT Proveedores.Id_proveedor FROM Proveedores WHERE Proveedores.RazonSocial = '" & ProveID & " '"
    
    ' Establecer la base de datos actual
    Set bd = CurrentDb
    
    ' Abrir el Recordset con la consulta
    Set rst = bd.OpenRecordset(strSQL)
    'If Form_frm_Productos.txtNomProd = "" Then
If Not rst.EOF Then
    
    Form_frm_Inventario.Text28txtIdPro.Value = rst!Id_proveedor
   
    
End If

rst.Close
Set rst = Nothing
Set bd = Nothing
End Sub
Sub idCat()
Dim bd As DAO.Database
    Dim rst As DAO.Recordset
    Dim strSQL2 As String
    Dim CatID As String
    
    'ProveID = Form_frm_Productos.cmbProvee.Value
CatID = Form_frm_Inventario.cmbCat.Value
'If IsNull(ProveID) Then
    strSQL2 = "SELECT IDCategoria FROM Categoria WHERE nombreCategoria = '" & CatID & "'"
'     ' Establecer la base de datos actual
    Set bd = CurrentDb
    ' Abrir el Recordset con la consulta
    Set rst = bd.OpenRecordset(strSQL2, dbOpenDynaset)
'    ' Verificar si se encontró algún registro
'    'If Form_frm_Productos.txtNomProd = "" Then
If Not rst.EOF Then
     Form_frm_Inventario.txtIdcat.Value = rst!IDCategoria
''    Else
''        MsgBox "No se encontró el cliente con el ID proporcionado.", vbExclamation
    End If
    
    rst.Close
    Set rst = Nothing
    Set bd = Nothing
    
End Sub

Sub idCiudad()
    Dim base As DAO.Database
    Dim rse As DAO.Recordset
    Dim strSQL3 As String
    Dim CiuID As String
    
    'ProveID = Form_frm_Productos.cmbProvee.Value
CiuID = Form_frm_ModuloClientes.cmbBuscarCiu.Column(1)
'If IsNull(ProveID) Then
    strSQL3 = "SELECT ID_Ciudad FROM Ciudad WHERE Nombre = '" & CiuID & "'"
'     ' Establecer la base de datos actual
    Set base = CurrentDb
    ' Abrir el Recordset con la consulta
    Set rse = base.OpenRecordset(strSQL3, dbOpenDynaset)
'    ' Verificar si se encontró algún registro
'    'If Form_frm_Productos.txtNomProd = "" Then
If Not rse.EOF Then
     Form_frm_ModuloClientes.txtIDP.Value = rse!ID_Ciudad
''    Else
''        MsgBox "No se encontró el cliente con el ID proporcionado.", vbExclamation
    End If
    
    rse.Close
    Set rse = Nothing
    Set base = Nothing
    
End Sub

Sub idCiudadClientes()
    Dim bas As DAO.Database
    Dim rst As DAO.Recordset
    Dim strSQL4 As String
    Dim CiuID As String
    
    'ProveID = Form_frm_Productos.cmbProvee.Value
CiuID = Form_frm_Clientes.cmbCiudad.Column(1)
'If IsNull(ProveID) Then
    strSQL4 = "SELECT ID_Ciudad FROM Ciudad WHERE Nombre = '" & CiuID & "'"
'     ' Establecer la base de datos actual
    Set bas = CurrentDb
    ' Abrir el Recordset con la consulta
    Set rst = bas.OpenRecordset(strSQL4, dbOpenDynaset)
'    ' Verificar si se encontró algún registro
'    'If Form_frm_Productos.txtNomProd = "" Then
If Not rst.EOF Then
     Form_frm_Clientes.txtIDCiu.Value = rst!ID_Ciudad
''    Else
''        MsgBox "No se encontró el cliente con el ID proporcionado.", vbExclamation
    End If
    
    rst.Close
    Set rst = Nothing
    Set bas = Nothing
    
End Sub

Function existeCat(Nombre As String) As Boolean
Dim b As DAO.Database
Dim r As DAO.Recordset
Dim sql As String
Dim ex As Boolean
If Not IsNull(Form_subFormCategoria) Then
    sql = "SELECT nombreCategoria FROM Categoria WHERE nombreCategoria = '" & Nombre & "'"
    Set b = CurrentDb
    Set r = b.OpenRecordset(sql, dbOpenDynaset)
    
If Not r.EOF Then
     ex = True
    Else
     ex = False
     End If
    
    r.Close
    Set r = Nothing
    Set b = Nothing
 existeCat = ex
 Else
 End If
End Function

Sub PedidoClientId()
Dim baseDat As DAO.Database
    Dim recSet As DAO.Recordset
    Dim strPC As String
    Dim CliID As Integer
    
    'ProveID = Form_frm_Productos.cmbProvee.Value
CliID = Form_frm_Pedido.Cliente
'If IsNull(ProveID) Then
    strPC = "SELECT ID_Pedido FROM Pedido WHERE Cliente = " & CliID & ";"
    
    '   Consulta = "SELECT Pedido.ID_Pedido " & _
'           "FROM Pedido INNER JOIN Clientes ON Pedido.ID_Pedido = Clientes.Id_Cliente " & _
'           "WHERE Clientes.Id_Cliente = '" & Me.Cliente.Value
'     ' Establecer la base de datos actual
    Set baseDat = CurrentDb
    ' Abrir el Recordset con la consulta
    Set recSet = baseDat.OpenRecordset(strPC, dbOpenDynaset)
'    ' Verificar si se encontró algún registro
'    'If Form_frm_Productos.txtNomProd = "" Then
If Not recSet.EOF Then
'Form_frm_Pedido.btnMenu.SetFocus
'Form_frm_Pedido.ID_Pedido.SetFocus
     Form_frm_Pedido.ID_Pedido.Value = recSet!ID_Pedido
''    Else
''        MsgBox "No se encontró el cliente con el ID proporcionado.", vbExclamation
    End If
    
    recSet.Close
    Set recSet = Nothing
    Set baseDat = Nothing

End Sub