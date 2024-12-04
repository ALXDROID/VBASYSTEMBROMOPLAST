SELECT Sum(IIf(Tipo = 'ingreso', Monto, 0)) AS TotalIngresos, Sum(IIf(Tipo = 'egreso', Monto, 0)) AS TotalEgresos, Sum(IIf(Tipo = 'ingreso', Monto, 0)) - Sum(IIf(Tipo = 'egreso', Monto, 0)) AS Saldo
FROM MovimientosCaja;

