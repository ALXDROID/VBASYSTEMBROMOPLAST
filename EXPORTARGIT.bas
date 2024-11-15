Option Compare Database
Option Explicit

Sub ExportVBA()
    Dim obj As AccessObject
    Dim exportPath As String
    exportPath = "C:\Users\Public\Documents\VBASYSTEMBROMOPLAST\" ' Cambia esto por la ubicación de tu repositorio

    ' Crear carpeta si no existe
    If Dir(exportPath, vbDirectory) = "" Then
        MkDir exportPath
    End If

    ' Exportar módulos estándar
    For Each obj In CurrentProject.AllModules
        Application.SaveAsText acModule, obj.Name, exportPath & obj.Name & ".bas"
    Next

    ' Exportar formularios con código VBA
    For Each obj In CurrentProject.AllForms
        On Error Resume Next
        DoCmd.OpenForm obj.Name, acDesign
        If Err.Number = 0 Then
            Application.SaveAsText acForm, obj.Name, exportPath & obj.Name & ".frm"
            DoCmd.Close acForm, obj.Name, acSaveNo
        End If
        On Error GoTo 0
    Next

    ' Exportar reportes con código VBA
    For Each obj In CurrentProject.AllReports
        On Error Resume Next
        DoCmd.OpenReport obj.Name, acDesign
        If Err.Number = 0 Then
            Application.SaveAsText acReport, obj.Name, exportPath & obj.Name & ".rep"
            DoCmd.Close acReport, obj.Name, acSaveNo
        End If
        On Error GoTo 0
    Next

    MsgBox "Exportación de código VBA completada."
End Sub