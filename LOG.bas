'Option Compare Database
'Option Explicit
'Public usuarioId As Integer
'
'Sub AgregarUsuario(nombreUsuario As String, tipo As Integer, contraseña As String, estado As Boolean)
'    Dim conn  As New ADODB.Connection
'    Dim sql As String
''    Dim contraseñaHash As String
'
'    ' Hashea la contraseña
''    contraseñaHash = GetSHA256Hash(contraseña)
'
'    ' Conecta a la base de datos y ejecuta la consulta de inserción
'    Set conn = CurrentProject.Connection
'    sql = "INSERT INTO tbl_Usuario(nom_usuario,tipo_usuario, pass_usuario,estado_usuario ) VALUES ('" & nombreUsuario & "',' & tipo & ', '" & contraseña & "', ' & estado & ')"
'    conn.Execute sql
'    Set conn = Nothing
'End Sub
'    Sub agrega()
'       Call AgregarUsuario("poi", 1, "123", True)
'
'    End Sub
''
''    Function GetSHA256Hash(input As String) As String
''        Dim sha256 As Object
''        Dim byteData() As Byte
''        Dim byteHash() As Byte
''        Dim i As Integer
''        Dim hash As String
''
''        ' Crear un objeto SHA-256
''        Set sha256 = CreateObject("System.Security.Cryptography.SHA256CryptoServiceProvider")
''
''        ' Convertir la cadena de entrada en bytes
''        byteData = StrConv(input, vbFromUnicode)
''
''        ' Computar el hash
''        byteHash = sha256.ComputeHash_2(byteData)
''
''        ' Convertir el hash en una cadena hexadecimal
''        For i = LBound(byteHash) To UBound(byteHash)
''            hash = hash & LCase(Right("0" & Hex(byteHash(i)), 2))
''        Next i
''
''        GetSHA256Hash = hash
''        Set sha256 = Nothing
''    End Sub
''
''
'Function ValidarLogin(nombreUsuario As String, contraseña As String) As Boolean
'    Dim conn As DAO.Database
'    Dim rs As DAO.Recordset
'    Dim sql As String
''    Dim contraseñaHash As String
'
'    ' Hashea la contraseña ingresada
''    contraseñaHash = GetSHA256Hash(contraseña)
'
'    ' Conectar a la base de datos y buscar el usuario
'    Set conn = CurrentDb()
'    sql = "SELECT * FROM tbl_Usuario WHERE nom_usuario = '" & nombreUsuario & "' AND pass_usuario = '" & contraseña & "'"
'    Set rs = conn.OpenRecordset(sql)
'
'    ' Verificar si el usuario existe
'    If Not rs.EOF Then
'        ValidarLogin = True
'
'
'    Else
'        ValidarLogin = False
'    End If
'usuarioId = rs!ID_usuario
'    rs.Close
'    Set rs = Nothing
'    Set conn = Nothing
'End Function
'Function telefonoCliente(tel As String, nom As String, apellido As String, ByVal ope As String)
'
'Dim db As New ADODB.Connection
'    Dim rs As New ADODB.Recordset
'    Dim sql As String
'    Dim strSQL As String
'    Dim sqlNombres As String
'    Dim newclient As Boolean
'    Dim newclient2 As Boolean
'    Dim numero1 As String
'    Dim strSQLNumero1 As String
'
'    If ope = "A" Then
'        Set db = CurrentProject.Connection
'        numero1 = "SELECT telefono,COUNT(*) As Cantidad FROM tbl_Cliente WHERE telefono = '" & tel & "' GROUP BY telefono HAVING COUNT(*) > 1 ;"
'        rs.Open numero1, db
'
'        If Not rs.EOF Then
'            MsgBox "Existen múltiples registros con el telefono '" & tel & "'.", vbInformation
'            Form_frm_ClientesPedido!frm_vistaPedidos.Form!listaPedidos.RowSource = ""
'            strSQLNumero1 = "SELECT * FROM tbl_Cliente WHERE telefono = '" & tel & "'"
'            rs.Close
'            rs.Open strSQLNumero1, db
'            Form_frm_ClientesPedido!frm_vistaPedidos.Form!listaPedidos.RowSource = strSQLNumero1
'            Form_frm_ClientesPedido!frm_vistaPedidos.Form!listaPedidos.ColumnHeads = True
'            Form_frm_ClientesPedido!frm_vistaPedidos.Form!listaPedidos.ColumnWidths = "0;5cm;5cm;0;5cm;7cm"
'            Form_frm_ClientesPedido!frm_vistaPedidos.Form!listaPedidos.Locked = False
'            newclient = MsgBox("Pulsa SI para seleccionar o No para crear nuevo Cliente.", vbYesNo, "La Mecha Company")
'
'            If newclient = vbYes Then
'                Form_frm_ClientesPedido!frm_vistaPedidos.Form!listaPedidos.SetFocus
'                rs.Close
'                Set rs = Nothing
'                Exit Function
'            Else
'                Form_frm_ClientesPedido.txt_ID_cliente.Value = ""
'                Form_frm_ClientesPedido.txt_nom_cliente.Value = ""
'                Form_frm_ClientesPedido.txt_apellido_cliente.Value = ""
'                Form_frm_ClientesPedido.txt_Direccion.Value = ""
'                Form_frm_ClientesPedido.txt_comentario.Value = ""
'                Exit Function
'            End If
'        End If
'    End If
'
'    If ope = "B" Then
'        Set db = CurrentProject.Connection
'        sqlNombres = "SELECT nom_cliente, apellido_cliente, COUNT(*) As Cantidad FROM tbl_Cliente WHERE apellido_cliente = '" & apellido & "' AND nom_cliente = '" & nom & "'  GROUP BY nom_cliente, apellido_cliente HAVING COUNT(*) > 1 ;"
'        rs.Open sqlNombres, db
'
'        If Not rs.EOF Then
'            MsgBox "Existen múltiples registros con el nombre '" & nom & "' y apellido '" & apellido & "'.", vbInformation
'            Form_frm_ClientesPedido!frm_vistaPedidos.Form!listaPedidos.RowSource = ""
'            strSQL = "SELECT * FROM tbl_Cliente WHERE nom_cliente = '" & nom & "' AND apellido_cliente = '" & apellido & "'"
'            rs.Close
'            rs.Open strSQL, db
'            Form_frm_ClientesPedido!frm_vistaPedidos.Form!listaPedidos.ColumnHeads = True
'            Form_frm_ClientesPedido!frm_vistaPedidos.Form!listaPedidos.RowSource = strSQL
'            Form_frm_ClientesPedido!frm_vistaPedidos.Form!listaPedidos.ColumnWidths = "0;5cm;5cm;0;5cm;7cm"
'            Form_frm_ClientesPedido!frm_vistaPedidos.Form!listaPedidos.Locked = False
'            newclient2 = MsgBox("Pulsa SI para seleccionar o No para crear nuevo Cliente.", vbYesNo, "La Mecha Company")
'
'            If newclient2 = vbYes Then
'                Form_frm_ClientesPedido!frm_vistaPedidos.Form!listaPedidos.SetFocus
'                rs.Close
'                Set rs = Nothing
'                Exit Function
'            Else
'                Form_frm_ClientesPedido.txt_ID_cliente.Value = ""
'                Form_frm_ClientesPedido.txt_Telefono.Value = ""
'                Form_frm_ClientesPedido.txt_Direccion.Value = ""
'                Form_frm_ClientesPedido.txt_comentario.Value = ""
'                Form_frm_ClientesPedido.txt_Telefono.SetFocus
'                Exit Function
'            End If
'        End If
'    End If
'
'' Buscar por teléfono
'    If ope = "A" Then
'        numero1 = "SELECT * FROM tbl_Cliente WHERE telefono = '" & tel & "'"
'        rs.Close
'        rs.Open numero1, db
'
'        If Not rs.EOF Then
'            ' Actualizar los controles con los datos encontrados
'            Form_frm_ClientesPedido.txt_ID_cliente.Value = rs!Id_Cliente
'            Form_frm_ClientesPedido.txt_Telefono.Value = rs!Telefono
'            Form_frm_ClientesPedido.txt_nom_cliente.Value = rs!nom_cliente
'            Form_frm_ClientesPedido.txt_apellido_cliente.Value = rs!apellido_cliente
'            Form_frm_ClientesPedido.txt_Direccion.Value = rs!Direccion
'            Form_frm_ClientesPedido.txt_comentario.Value = rs!comentario
'
'            rs.Close
'            Set rs = Nothing
'            db.Close
'            Set db = Nothing
'            Exit Function
'        End If
'    End If
'
'    ' Buscar por nombre y apellido
'    If ope = "B" Then
'        sqlNombres = "SELECT * FROM tbl_Cliente WHERE apellido_cliente = '" & apellido & "' AND nom_cliente = '" & nom & "'"
'        rs.Close
'        rs.Open sqlNombres, db
'
'        If Not rs.EOF Then
'            ' Actualizar los controles con los datos encontrados
'            Form_frm_ClientesPedido.txt_ID_cliente.Value = rs!Id_Cliente
'            Form_frm_ClientesPedido.txt_Telefono.Value = rs!Telefono
'            Form_frm_ClientesPedido.txt_nom_cliente.Value = rs!nom_cliente
'            Form_frm_ClientesPedido.txt_apellido_cliente.Value = rs!apellido_cliente
'            Form_frm_ClientesPedido.txt_Direccion.Value = rs!Direccion
'            Form_frm_ClientesPedido.txt_comentario.Value = rs!comentario
'
'            rs.Close
'            Set rs = Nothing
'            db.Close
'            Set db = Nothing
'            Exit Function
'        End If
'    End If
'    rs.Close
'    Set rs = Nothing
'    db.Close
'    Set db = Nothing
'
'End Function
'
'