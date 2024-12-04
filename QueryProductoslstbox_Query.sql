SELECT Productos.IDProd, Productos.nombre AS PRODUCTO, Proveedores.RazonSocial AS PROVEEDOR, Categoria.nombreCategoria AS CATEGORIA, Productos.stockActual AS STOCK, Productos.precioUnidad AS PRECIO, Productos.stockMinimo, Productos.descripcion AS DESCRIPCION, Productos.imagen
FROM Proveedores INNER JOIN (Categoria INNER JOIN Productos ON Categoria.IDCategoria = Productos.categoria) ON Proveedores.Id_proveedor = Productos.proveedor
WHERE Productos.disponible = True AND Categoria.Avalaible = True AND  Proveedores.activo = True;

