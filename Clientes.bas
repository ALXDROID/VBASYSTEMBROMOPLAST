Option Compare Database
Option Explicit

Sub Detalle()
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim strSQL As String
    Dim fecha As Date
    Dim PedidoId As Integer

    ' Obtener y formatear la fecha
    fecha = Form_frm_ModuloClientes.cmbFechaFact.Value
    
    ' Obtener el ID del pedido
    PedidoId = Form_frm_ModuloClientes.txtIDPed.Value
    
    ' Construir la consulta SQL con fechas rodeadas por #
    strSQL = "SELECT * FROM Pedido WHERE Fecha = #" & Format(fecha, "mm/dd/yyyy") & "# AND ID_Pedido = " & PedidoId
    
    ' Abrir la base de datos y el conjunto de registros
    Set db = CurrentDb
    Set rs = db.OpenRecordset(strSQL, dbOpenDynaset)
      
    ' Verificar si el conjunto de registros est� vac�o
    If Not rs.EOF Then
        ' Mover el puntero al primer registro
        rs.MoveFirst

        ' Acceder a los datos
        MsgBox rs!Descripcion

        ' Asignar valores a los controles del formulario
        Form_frm_ModuloClientes.txt_Factura.Value = rs!Factura
        Form_frm_ModuloClientes.txt_Descripcion.Value = rs!Descripcion
    Else
        ' Mostrar un mensaje si no se encuentran registros
        MsgBox "No se encontr� ning�n registro con los criterios especificados."
    End If

    ' Cerrar el conjunto de registros y la base de datos
    rs.Close
    Set rs = Nothing
    Set db = Nothing
End Sub