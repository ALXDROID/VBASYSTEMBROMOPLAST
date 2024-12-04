SELECT Categoria.nombreCategoria, Productos.nombre, Productos.stockActual, Productos.precioUnidad, Proveedores.RazonSocial, *
FROM Proveedores INNER JOIN (Categoria INNER JOIN Productos ON Categoria.IDCategoria = Productos.categoria) ON Proveedores.Id_proveedor = Productos.proveedor
WHERE Productos.disponible = true;

