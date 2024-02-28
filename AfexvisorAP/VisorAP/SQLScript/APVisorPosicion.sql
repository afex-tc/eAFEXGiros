DECLARE	@CodigoMoneda Char(3)
DECLARE	@Fecha DateTime

SET @CodigoMoneda = «Moneda»
SET @Fecha = «Fecha»

DECLARE @PosicionAnt Numeric(12, 2)
DECLARE @TCCierreAnt Numeric(8, 2)
DECLARE @PosicionNacionalAnt Numeric(14)
DECLARE @Posicion Numeric(12, 2)
DECLARE @PosicionNacional Numeric(14)
DECLARE @TCCierre Numeric(8, 2)

DECLARE @Compra Numeric(12, 2)
DECLARE @TCCompra Numeric(8, 2)
DECLARE @CompraNacional Numeric(14)

DECLARE @Venta Numeric(12, 2)
DECLARE @TCVenta Numeric(8, 2)
DECLARE @VentaNacional Numeric(14)


-- Inicio obtener el saldo anterior de la Posicion
-- Para un mejor rendimiento del Visor, cerrar en AFEXchange hasta el día anterior
	DECLARE @FechaBalance DateTime
	SET @FechaBalance = (
				SELECT CONVERT(CHAR, MAX(fecha_balance), 112) 
				FROM Balance
				 WHERE fecha_balance<@Fecha
			)
	SET @PosicionAnt = -1 * ISNULL(
						(
			
						SELECT 	ISNULL(SUM(ba.saldo_extranjera), 0) AS Saldo
						FROM		Balance BA
						JOIN		plan_cuenta pc ON ba.numero_cuenta = pc.numero_cuenta
						WHERE 	pc.uso_cuenta=9 
						AND 		pc.codigo_moneda=@CodigoMoneda
						AND 		pc.tipo_moneda=2
						AND 		ba.fecha_balance = @FechaBalance
						)		
					, 0)
	
	DECLARE @CompraAnt Numeric(12, 2)
	SET @CompraAnt = ISNULL(
				(
				SELECT 	SUM(monto_extranjera) AS Monto
				FROM 		detalle_solicitud dsp
				JOIN		Solicitud SP ON sp.codigo_solicitud = dsp.codigo_solicitud
				WHERE   	fecha_solicitud > @FechaBalance AND fecha_solicitud < @Fecha
						AND	dsp.tipo_operacion IN (1)
						AND	sp.estado_solicitud <> 0
						AND 	codigo_moneda = @CodigoMoneda
				), 0)
	
	DECLARE @VentaAnt Numeric(12, 2)
	SET @VentaAnt = ISNULL(
				(
				SELECT 	SUM(monto_extranjera) AS Monto
				FROM 		detalle_solicitud dsp
				JOIN		Solicitud SP ON sp.codigo_solicitud = dsp.codigo_solicitud
				WHERE   	fecha_solicitud > @FechaBalance AND fecha_solicitud < @Fecha
						AND	dsp.tipo_operacion IN (2)
						AND	sp.estado_solicitud <> 0
						AND 	codigo_moneda = @CodigoMoneda
				), 0)
	
	SET @PosicionAnt = @PosicionAnt + @CompraAnt - @VentaAnt
-- Fin obtener el saldo anterior de la Posicion


SET @TCCierreAnt = 0
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[APPrecio_Referencia]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
	SET @TCCierreAnt =	IsNull(
					(SELECT TOP 1 CASE WHEN @PosicionAnt >= 0 
							THEN pr.pm_compra
							ELSE pr.pm_venta
						END AS PrecioCierre
					FROM	APPrecio_Referencia PR
					WHERE fecha = (	
								SELECT CONVERT(CHAR, MAX(fecha), 112) 
								FROM APPrecio_Referencia
								 WHERE fecha < @Fecha
									AND codigo = @CodigoMoneda
							)
					AND	codigo = @CodigoMoneda
					ORDER BY Correlativo DESC
					)
				, 0)
END

SET @PosicionNacionalAnt = ROUND( (@PosicionAnt * @TCCierreAnt), 2)

SET @Compra = ISNULL(
			(
			SELECT 	SUM(monto_extranjera) AS Monto
			FROM 		detalle_solicitud dsp
			JOIN		Solicitud SP ON sp.codigo_solicitud = dsp.codigo_solicitud
			WHERE   	fecha_solicitud = @Fecha
					AND	dsp.tipo_operacion IN (1)
					AND	sp.estado_solicitud <> 0
					AND 	codigo_moneda = @CodigoMoneda
			), 0)

SET @CompraNacional = ISNULL(
			(
			SELECT 	SUM(monto_nacional) AS Monto
			FROM 		detalle_solicitud dsp
			JOIN		Solicitud SP ON sp.codigo_solicitud = dsp.codigo_solicitud
			WHERE   	fecha_solicitud = @Fecha
					AND	dsp.tipo_operacion IN (1)
					AND	sp.estado_solicitud <> 0
					AND 	codigo_moneda = @CodigoMoneda
			), 0)

SET @Venta =	ISNULL(
			(
			SELECT 	SUM(monto_extranjera) AS Monto
			FROM 		detalle_solicitud dsp
			JOIN		Solicitud SP ON sp.codigo_solicitud = dsp.codigo_solicitud
			WHERE   	fecha_solicitud = @Fecha
					AND	dsp.tipo_operacion IN (2)
					AND	sp.estado_solicitud <> 0
					AND 	codigo_moneda = @CodigoMoneda
			), 0)
SET @VentaNacional =	ISNULL(
			(
			SELECT 	SUM(monto_nacional) AS Monto
			FROM 		detalle_solicitud dsp
			JOIN		Solicitud SP ON sp.codigo_solicitud = dsp.codigo_solicitud
			WHERE   	fecha_solicitud = @Fecha
					AND	dsp.tipo_operacion IN (2)
					AND	sp.estado_solicitud <> 0
					AND 	codigo_moneda = @CodigoMoneda			), 0)

SET @Posicion = @PosicionAnt + @Compra - @Venta
SET @PosicionNacional = @PosicionNacionalAnt + @CompraNacional - @VentaNacional

SET @TCCierre = CASE WHEN @Posicion >= 0 
			THEN 
				CASE WHEN @PosicionAnt >= 0
					THEN
						ROUND((ABS(@PosicionNacionalAnt) + @CompraNacional) / (ABS(@PosicionAnt) + @Compra), 2)
					ELSE
						ROUND( @CompraNacional / @Compra,2)
				END
			ELSE
				CASE WHEN @PosicionAnt < 0
					THEN
						ROUND((ABS(@PosicionNacionalAnt) + @VentaNacional) / (ABS(@PosicionAnt) + @Venta), 2)
					ELSE
						ROUND( @VentaNacional / @Venta, 2)
				END
				
		END


DECLARE @TC Numeric(8, 4)
SET @TCCompra = 	CASE WHEN @CompraNacional <> 0 AND @Compra <> 0
				THEN ROUND(@CompraNacional / @Compra, 2)
				ELSE 0
			END
SET @TCVenta = 	CASE WHEN @VentaNacional <> 0 AND @Venta <> 0
				THEN ROUND(@VentaNacional / @Venta, 2)
				ELSE 0
			END
SET @TC = 	(
				CASE WHEN @Posicion >= 0
					THEN	(@TCVenta - @TCCierre)
					ELSE	(@TCCierre - @TCCompra)
				END 
			)

DECLARE @PosicionTotal Numeric(12, 4)
SET @PosicionTotal =	CASE WHEN @PosicionAnt >= 0
				THEN	@Posicion + @Compra
				ELSE	ABS(@Posicion ) + @Venta
			END




SELECT 1 AS Tipo, 'Posicion Inicial' AS Nombre, @PosicionAnt AS MontoExtranjera, ROUND(@TCCierreAnt, 2) AS TipoCambio

UNION

SELECT 2 AS Tipo, 'Compras' AS Nombre, @Compra AS MontoExtranjera, CASE WHEN @CompraNacional <> 0 AND @Compra <> 0
										THEN ROUND( @CompraNacional / @Compra,2) 
										ELSE 0
									END AS TipoCambio
UNION

SELECT 3 AS Tipo, 'Ventas' AS Nombre, @Venta AS MontoExtranjera, CASE WHEN @VentaNacional <> 0 AND @Venta <> 0
									THEN ROUND(@VentaNacional / @Venta, 2) 
									ELSE 0
								END AS TipoCambio

UNION

SELECT 4 AS Tipo, 'Posicion Final' AS Nombre, @Posicion AS MontoExtranjera, ROUND(@TCCierre, 2) AS TipoCambio


ORDER BY Tipo
