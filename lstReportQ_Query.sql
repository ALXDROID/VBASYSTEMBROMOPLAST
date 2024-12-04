SELECT DetallePedido.PedidoId, DetallePedido.Cantidad, Productos.nombre, Productos.precioUnidad, [Cantidad]*[precioUnidad] AS extPrecio, *
FROM DetallePedido, Productos
WHERE DetallePedido.PedidoId = ID_Pedido;

