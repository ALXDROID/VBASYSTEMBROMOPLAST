SELECT Pedido.ID_Pedido, Pedido.Descripcion, Pedido.Total, Clientes.Nombre, Clientes.Direccion, Clientes.Telefono, Ciudad.Nombre
FROM (Ciudad INNER JOIN Clientes ON Ciudad.ID_Ciudad = Clientes.ciudad) INNER JOIN Pedido ON Clientes.Id_Cliente = Pedido.Cliente;

