Option Compare Database
Option Explicit

Sub ExportVBA()
    Dim obj As AccessObject
    Dim exportPath As String
    exportPath = "C:\Users\Public\Documents\VBASYSTEMBROMOPLAST\" ' Cambia esto por la ubicaci�n de tu repositorio

    ' Crear carpeta si no existe
    If Dir(exportPath, vbDirectory) = "" Then
        MkDir exportPath
    End If

    ' Exportar m�dulos est�ndar
    For Each obj In CurrentProject.AllModules
        Application.SaveAsText acModule, obj.Name, exportPath & obj.Name & ".bas"
    Next

    ' Exportar formularios con c�digo VBA
    For Each obj In CurrentProject.AllForms
        On Error Resume Next
        DoCmd.OpenForm obj.Name, acDesign
        If Err.Number = 0 Then
            Application.SaveAsText acForm, obj.Name, exportPath & obj.Name & ".frm"
            DoCmd.Close acForm, obj.Name, acSaveNo
        End If
        On Error GoTo 0
    Next

    ' Exportar reportes con c�digo VBA
    For Each obj In CurrentProject.AllReports
        On Error Resume Next
        DoCmd.OpenReport obj.Name, acDesign
        If Err.Number = 0 Then
            Application.SaveAsText acReport, obj.Name, exportPath & obj.Name & ".rep"
            DoCmd.Close acReport, obj.Name, acSaveNo
        End If
        On Error GoTo 0
    Next

    MsgBox "Exportaci�n de c�digo VBA completada."
End Sub