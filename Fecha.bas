Option Compare Database
Option Explicit

'Function TraducirFechaEnEspanol(fecha As Date) As String
'    Dim diaSemana As String
'    Dim mes As String
'
'    ' Traducir los días de la semana
'    Select Case Format(fecha, "dddd")
'        Case "Sunday": diaSemana = "domingo"
'        Case "Monday": diaSemana = "lunes"
'        Case "Tuesday": diaSemana = "martes"
'        Case "Wednesday": diaSemana = "miércoles"
'        Case "Thursday": diaSemana = "jueves"
'        Case "Friday": diaSemana = "viernes"
'        Case "Saturday": diaSemana = "sábado"
'    End Select
'
'    ' Traducir los meses
'    Select Case Format(fecha, "mmmm")
'        Case "January": mes = "enero"
'        Case "February": mes = "febrero"
'        Case "March": mes = "marzo"
'        Case "April": mes = "abril"
'        Case "May": mes = "mayo"
'        Case "June": mes = "junio"
'        Case "July": mes = "julio"
'        Case "August": mes = "agosto"
'        Case "September": mes = "septiembre"
'        Case "October": mes = "octubre"
'        Case "November": mes = "noviembre"
'        Case "December": mes = "diciembre"
'    End Select
'
'    ' Devolver la fecha formateada en español
'    TraducirFechaEnEspanol = diaSemana & ", " & mes & " " & Format(fecha, "yyyy")
'End Function