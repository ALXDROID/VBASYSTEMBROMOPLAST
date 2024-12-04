SELECT DetallePedido.*, Productos.IDProd, Productos.nombre, Productos.precioUnidad, [Cantidad]*[precioUnidad] AS ExtPrecio
FROM Productos INNER JOIN DetallePedido ON Productos.IDProd = DetallePedido.ProductoId;

