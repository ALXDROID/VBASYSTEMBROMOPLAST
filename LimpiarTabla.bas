Option Compare Database
Option Explicit

'Sub BorrarYRestablecerAutoIncrement()
'    Dim db As DAO.Database
'    Dim tabla As String
'
'    ' Nombre de la tabla a limpiarSaldoCaja,ID_Saldo(SaldoCaja),DetallePedido(DetPedidoId),MovimientosCaja(IDMov)
'    '
'    tabla = "Pedido"
'
'    ' Obtener referencia a la base de datos actual
'    Set db = CurrentDb
'
'    ' Eliminar todos los registros de la tabla
'    db.Execute "DELETE FROM " & tabla, dbFailOnError
'
'    ' Restablecer el campo autonumérico
'    db.Execute "ALTER TABLE " & tabla & " ALTER COLUMN ID_Pedido COUNTER (1, 1);", dbFailOnError
'
'    MsgBox "Registros eliminados y autonumérico restablecido.", vbInformation, "Proceso Completado"
'End Sub

Sub BorrarYRestablecerAutoIncrement()
    Dim db As DAO.Database
    Dim tabla As String
    Dim campoAutonumerico As String
    Dim tablasYCampos As Variant
    Dim i As Integer
    
    ' Listado de tablas y sus campos autonuméricos
    tablasYCampos = Array( _
        Array("SaldoCaja", "ID_Saldo"), _
        Array("DetallePedido", "DetPedidoId"), _
        Array("MovimientosCaja", "IDMov"), _
        Array("Pedido", "ID_Pedido") _
    )
    
    ' Obtener referencia a la base de datos actual
    Set db = CurrentDb
    
    ' Iterar por todas las tablas y restablecer sus autonuméricos
    For i = LBound(tablasYCampos) To UBound(tablasYCampos)
        tabla = tablasYCampos(i)(0)
        campoAutonumerico = tablasYCampos(i)(1)
        
        On Error Resume Next ' Manejo de errores para evitar que se detenga el bucle si hay un error
        ' Eliminar todos los registros de la tabla
        db.Execute "DELETE FROM " & tabla, dbFailOnError
        
        ' Restablecer el campo autonumérico
        db.Execute "ALTER TABLE " & tabla & " ALTER COLUMN " & campoAutonumerico & " COUNTER (1, 1);", dbFailOnError
        On Error GoTo 0
    Next i
    
    MsgBox "Registros eliminados y autonuméricos restablecidos para todas las tablas.", vbInformation, "Proceso Completado"
End Sub