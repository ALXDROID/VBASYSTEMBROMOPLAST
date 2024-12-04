SELECT Proveedores.Id_proveedor, Proveedores.Nombre, Proveedores.RazonSocial, Proveedores.Direccion, Proveedores.Telefono, Proveedores.e_mail, *
FROM Proveedores
WHERE Proveedores.activo = True;

