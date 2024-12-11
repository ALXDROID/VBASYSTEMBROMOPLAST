Option Compare Database
Option Explicit

Sub BorrarYRestablecerAutoIncrement()
    Dim db As DAO.Database
    Dim tabla As String
    
    ' Nombre de la tabla a limpiar
    tabla = "SaldoCaja"
    
    ' Obtener referencia a la base de datos actual
    Set db = CurrentDb
    
    ' Eliminar todos los registros de la tabla
    db.Execute "DELETE FROM " & tabla, dbFailOnError
    
    ' Restablecer el campo autonumérico
    db.Execute "ALTER TABLE " & tabla & " ALTER COLUMN ID_Saldo COUNTER (1, 1);", dbFailOnError
    
    MsgBox "Registros eliminados y autonumérico restablecido.", vbInformation, "Proceso Completado"
End Sub