Option Compare Database
Option Explicit

Sub ExportarTablasYConsultas()
    Dim db As DAO.Database
    Dim td As DAO.TableDef
    Dim qd As DAO.QueryDef
    Dim fs As Object
    Dim archivoSQL As Object
    Dim rutaDestino As String
    Dim nombreArchivo As String
    Dim sqlScript As String
    
    ' Ruta destino (cambia a la ubicación deseada)
    rutaDestino = "C:\Users\Public\Documents\VBASYSTEMBROMOPLAST\" ' Asegúrate de que la carpeta exista.
    
    ' Crear objeto FileSystemObject
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    ' Crear carpeta si no existe
    If Not fs.FolderExists(rutaDestino) Then
        fs.CreateFolder rutaDestino
    End If
    
    ' Abre la base de datos actual
    Set db = CurrentDb
    
    ' Exportar tablas
    For Each td In db.TableDefs
        If Left(td.Name, 4) <> "MSys" Then ' Ignorar tablas del sistema
            nombreArchivo = rutaDestino & td.Name & "_Table.sql"
            Set archivoSQL = fs.CreateTextFile(nombreArchivo, True)
            sqlScript = GenerarScriptTabla(td)
            archivoSQL.WriteLine sqlScript
            archivoSQL.Close
        End If
    Next td
    
    ' Exportar consultas
    For Each qd In db.QueryDefs
    If Left(qd.Name, 1) <> "~" Then ' Ignorar consultas internas
        ' Limpia el nombre del archivo
        nombreArchivo = rutaDestino & LimpiarNombreArchivo(qd.Name) & "_Query.sql"
        Set archivoSQL = fs.CreateTextFile(nombreArchivo, True)
        sqlScript = qd.sql
        archivoSQL.WriteLine sqlScript
        archivoSQL.Close
    End If
Next qd

   
    
    ' Liberar memoria
    Set archivoSQL = Nothing
    Set fs = Nothing
    Set db = Nothing
    
    MsgBox "Exportación completada en " & rutaDestino, vbInformation
End Sub
Function LimpiarNombreArchivo(Nombre As String) As String
    Dim caracteresProhibidos As String
    Dim i As Integer
    Dim caracter As String
    
    ' Lista de caracteres prohibidos en nombres de archivo
    caracteresProhibidos = "\/:*?""<>|"
    
    ' Reemplazar caracteres prohibidos por guión bajo
    For i = 1 To Len(caracteresProhibidos)
        caracter = Mid(caracteresProhibidos, i, 1)
        Nombre = Replace(Nombre, caracter, "_")
    Next i
    
    LimpiarNombreArchivo = Nombre
End Function


Function GenerarScriptTabla(td As DAO.TableDef) As String
    Dim campo As DAO.Field
    Dim sqlScript As String
    Dim primaryKeys As String
    
    ' Inicia el script
    sqlScript = "CREATE TABLE " & td.Name & " (" & vbCrLf
    
    ' Agrega los campos
    For Each campo In td.Fields
        sqlScript = sqlScript & "    " & campo.Name & " " & ObtenerTipoSQL(campo) & "," & vbCrLf
    Next campo
    
    ' Elimina la última coma
    sqlScript = Left(sqlScript, Len(sqlScript) - 3) & vbCrLf & ");"
    
    GenerarScriptTabla = sqlScript
End Function

Function ObtenerTipoSQL(campo As DAO.Field) As String
    Select Case campo.Type
        Case dbText
            ObtenerTipoSQL = "VARCHAR(" & campo.Size & ")"
        Case dbLong
            ObtenerTipoSQL = "INTEGER"
        Case dbDouble
            ObtenerTipoSQL = "DOUBLE"
        Case dbDate
            ObtenerTipoSQL = "DATETIME"
        Case dbBoolean
            ObtenerTipoSQL = "BOOLEAN"
        Case Else
            ObtenerTipoSQL = "TEXT"
    End Select
End Function