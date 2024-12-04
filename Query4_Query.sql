SELECT Clientes.Id_Cliente, Clientes.Nombre AS Nombre, Ciudad.Nombre AS Ciudad, Clientes.Direccion, Clientes.Telefono, Clientes.e_mail, Clientes.rut, Clientes.activo, Clientes.comentario
FROM Ciudad INNER JOIN Clientes ON Ciudad.ID_Ciudad = Clientes.ciudad
GROUP BY Clientes.Id_Cliente, Clientes.Nombre, Ciudad.Nombre, Clientes.Direccion, Clientes.Telefono, Clientes.e_mail, Clientes.rut, Clientes.activo, Clientes.comentario
HAVING (((Clientes.Nombre) Like Forms!frm_ModuloClientes!txtBusqueda.text & "*") And ((Ciudad.Nombre) Like Forms!frm_ModuloClientes!cmbBuscarCiu.text & "*") And ((Clientes.activo)=True))
ORDER BY Clientes.Nombre;

