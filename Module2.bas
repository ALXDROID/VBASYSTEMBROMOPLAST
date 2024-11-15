'Option Compare Database
'Option Explicit
'
'
'
'Private rs As ADODB.Recordset
'Private cnn As ADODB.Connection
'Private strSQL As String
'
'''''''''''''''''''''''''''''''''''''''''''''
''''''Initilisation and kill code''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''
'Private Sub Class_Initialize()
'Set cnn = CurrentProject.Connection
'Set rs = New ADODB.Recordset
'End Sub
'
'Private Sub Class_Terminate()
'Call killRecordset
'End Sub
'
'Public Sub loadData(id As Long)
''This sub procedure loads the data based upon the
''passed in ID value -1 means new record
'
'If id > 0 Then
'    strSQL = "Select * From tbl_Cliente Where [ID_cliente]=" & Id_Cliente
'    Call loadRecordset(strSQL)
'    Call setFields
'Else
'    m_ID = -1
'    strSQL = "Select * From tbl_Cliente "
'    Call loadRecordset(strSQL)
'End If
'End Sub
'
'''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''Recordset code''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''
'
'Private Sub loadRecordset(strSQL As String)
'With rs
'        .Open strSQL, cnn, adOpenKeyset, adLockOptimistic
'        .MoveLast
'        .MoveFirst
'    End With
'End Sub
'
'Private Sub killRecordset()
'If Not rs Is Nothing Then rs.Close
'Set rs = Nothing
'Set cnn = Nothing
'End Sub
'
'Public Sub setFields()
'With rs
'    m_ID_cliente = .Fields("ID_cliente")
'    m_nom_cliente = .Fields("nom_cliente")
'    m_apellido_cliente = .Fields("apellido_cliente")
'    m_comentario = .Fields("comentario")
'    m_telefono = .Fields("telefono")
'    m_direccion = .Fields("direccion")
'End With
'End Sub
'
'''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''Record Operations'''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''
'
'Public Function SaveRecord() As Boolean
'
'
'
'With rs
'    If m_ID_cliente > 0 Then
'
'        .Fields("ID_cliente") = m_ID_cliente
'        .Fields("nom_cliente") = m_nom_cliente
'        .Fields("apellido_cliente") = m_apellido_cliente
'        .Fields("comentario") = m_comentario
'        .Fields("telefono") = m_telefono
'        .Fields("direccion") = m_direccion
'
'        .update
'    Else
'        .AddNew
'        .Fields("nom_cliente") = m_nom_cliente
'        .Fields("apellido_cliente") = m_apellido_cliente
'        .Fields("comentario") = m_comentario
'        .Fields("telefono") = m_telefono
'        .Fields("direccion") = m_direccion
'        .update
'    End If
'
'End With
'
'End Function
'
'Public Function UndoRecord() As Boolean
'With rs
'    m_ID = .Fields("ID")
'    m_FilmName = .Fields("FilmName")
'    m_YearOfRelease = .Fields("YearOfRelease")
'    m_RottenTomato = .Fields("RottenTomatoes")
'    m_Director = .Fields("DirectorID")
'End With
'End Function
'
'Public Function DeleteRecord() As Boolean
'Dim lngID As Long
'
'lngID = m_ID
''We are going to use a simple query to delete the record
'Call killRecordset
'
'strSQL = "DELETE FROM Films WHERE [ID]=" & lngID
'CurrentDb.Execute strSQL, dbFailOnError
'
'End Function
'