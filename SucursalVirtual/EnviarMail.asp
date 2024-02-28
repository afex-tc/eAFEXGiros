<%@ Language=VBScript %>
<!--#include virtual="/Compartido/Rutinas.asp" -->
<%
Dim sSQL
Dim rsHistoria
Dim sMensaje

Dim sDescripcion
Dim sAsunto
Dim sMoneda

dim dbConexion
dim rs
dim sMailEjecutivo
dim sEjecutivo
dim sBancoOrigen
dim sCuentaOrigen
dim sEstadoTRF
dim sSolicitud
dim sSolicitudOriginal
dim sLineaSolicitud
dim sTiempo
dim cVoucher
dim sGlosa

dim sCodigoBanco

sMailEjecutivo = "domingo.a.avila@gmail.com"
sLineaSolicitud = 1

'On Error Resume Next

sDescripcion = "					Desde la Sucursal Virtual el Cliente: " & Session("NombreOperador") & " CÓDIGO: " & Session("CodigoCliente") & vbCrlf 
Select Case Request("Accion")
	Case 1	' comprar dolares
		sAsunto = "Compra de US$ SV"
		sDescripcion = sDescripcion & "desea vender US$ con el siguiente detalle: " & _
							vbCrlf & vbCrlf 

		sDescripcion = sDescripcion & vbCrlf & "MONTO  : US$ " & trim(request.Form("txtMontoC"))
		sDescripcion = sDescripcion & vbCrlf & "CAMBIO : US$ " & trim(request.Form("txtTCC"))
		sDescripcion = sDescripcion & vbCrlf & "TOTAL  :   $ " & trim(request.Form("txtTotalC"))
		sDescripcion = sDescripcion & vbCrlf & "BANCO	: " & trim(request.Form("txtBancoC"))
		sDescripcion = sDescripcion & vbCrlf & "CTA.CTE.: " & trim(request.Form("txtCtaCteC"))
		
	Case 2	' vender dolares
		sAsunto = "Venta de US$ SV"
		sDescripcion = sDescripcion & "desea comprar US$ con el siguiente detalle: " & _
							vbCrlf & vbCrlf 

		sDescripcion = sDescripcion & vbCrlf & "MONTO  : US$ " & trim(request.Form("txtMontoV"))
		sDescripcion = sDescripcion & vbCrlf & "CAMBIO : US$ " & trim(request.Form("txtTCV"))
		sDescripcion = sDescripcion & vbCrlf & "TOTAL  :   $ " & trim(request.Form("txtTotalV"))
		sDescripcion = sDescripcion & vbCrlf & "BANCO	: " & trim(request.Form("txtBancoV"))
		sDescripcion = sDescripcion & vbCrlf & "CTA.CTE.: " & trim(request.Form("txtCtaCteV"))

	Case 3	' enviar giro		
		Select Case request.Form("cbxPaisB")
			Case "CL"
				sMoneda = "PESOS CHILENOS"
			Case Else
				sMoneda = "DOLARES"
		End Select
		
		sAsunto = "Envío de Giro SV"
		sDescripcion = sDescripcion & "solicita el envío del siguiente Giro: " & _
							vbCrlf & vbCrlf 

		sDescripcion = sDescripcion & vbCrlf & "BENEFICIARIO "
		sDescripcion = sDescripcion & vbCrlf & "NOMBRE : " & trim(request.Form("txtNombreB")) & " " & trim(request.Form("txtApellidoB"))
		sDescripcion = sDescripcion & vbCrlf & "PAIS   : " & trim(request.Form("txtPaisB"))
		sDescripcion = sDescripcion & vbCrlf & "CIUDAD : " & trim(request.Form("txtCiudadB"))
		sDescripcion = sDescripcion & vbCrlf & "DIRECCIÓN : " & trim(request.Form("txtDireccionB"))
		sDescripcion = sDescripcion & vbCrlf & "TELÉFONO  : " & trim(request.Form("txtFonoB"))
		sDescripcion = sDescripcion & vbCrlf & "MENSAJE : " & trim(request.Form("txtMensajeB"))
		sDescripcion = sDescripcion & vbCrlf & "MONEDA : " & trim(sMoneda)
		sDescripcion = sDescripcion & vbCrlf & "MONTO  : " & trim(request.Form("txtMonto"))
		sDescripcion = sDescripcion & vbCrlf & "TARIFA : " & trim(request.Form("txtTarifaSugerida"))
		sDescripcion = sDescripcion & vbCrlf & "TOTAL  : " & trim(request.Form("txtTotal"))
		sDescripcion = sDescripcion & vbCrlf & "TIPO CAMBIO : " & trim(request.Form("txtTipoCambio"))
		sDescripcion = sDescripcion & vbCrlf & "MONTO PESOS : " & trim(request.Form("txtMontoPesos"))
		sDescripcion = sDescripcion & vbCrlf & "PAGADOR : " & trim(request.Form("txtPagador"))
	
	Case 5	' pagar tarjeta de credito
		sAsunto = "Pagar Tarjeta de Crédito en US$ SV"
		sDescripcion = sDescripcion & "desea pagar su tarjeta de crédito en US$ con el siguiente detalle: " & _
							vbCrlf & vbCrlf 

		sDescripcion = sDescripcion & vbCrlf & "MONTO  : US$ " & trim(request.Form("txtMontoTC"))
		sDescripcion = sDescripcion & vbCrlf & "Nº DE CTA. DE TARJETA : " & trim(request.Form("txtNTarjeta"))
		sDescripcion = sDescripcion & vbCrlf & "BANCO EMISOR DE LA TARJETA :     " & trim(request.Form("txtBanco"))
		sDescripcion = sDescripcion & vbCrlf & "TIPO CAMBIO : " & trim(request.Form("txtTCT"))
		sDescripcion = sDescripcion & vbCrlf & "PESOS A DEPOSITAR : " & trim(request.Form("txtTotalT"))
	
	Case 6	' enviar transfer
		' Jonathan Miranda G. 19-06-2007
		' saca los parametros de la tabla configuracion
		sSQL = "select e.valor as ejecutivo, m.valor as mail, b.valor as bancoorigen, c.valor as cuentaorigen, " & _
					" es.valor as estado, t.valor as tiempo, mo.valor as monedaequivalente " & _
				"from configuracion e " & _
				" inner join configuracion m on m.item = e.item and m.campo = 'mail' " & _
				" inner join configuracion b on b.item = e.item and b.campo = 'bancoorigen' " & _
				" inner join configuracion c on c.item = e.item and c.campo = 'cuentaorigen' " & _
				" inner join configuracion es on es.item = e.item and es.campo = 'estado' " & _
				" inner join configuracion t on t.item = e.item and t.campo = 'tiempo' " & _
				" inner join configuracion mo on mo.item = e.item and t.campo = 'monedaequivalente' " & _
				"where e.item = 'transfervirtual' and " & _
					" e.campo = 'ejecutivo' "
		set rs = ejecutarsqlcliente(session("afxCnxAFEXchange"), sSQL)
		if err.number <> 0 then
			response.Write "Ocurrió un error al buscar al ejecutivo y su mail. " & err.Description
			response.End
		end if
		if rs.eof then
			sBancoOrigen = "6"
			sCuentaOrigen = "6550628510"
			sEjecutivo = "AANTOLIN"
			sEstadoTRF = "9"
			sTiempo = "48"
			sMonedaEquivalente = "USD"
		else
			sBancoOrigen = rs("bancoorigen")
			sCuentaOrigen = rs("cuentaorigen")
			sMailEjecutivo = rs("mail")
			sEjecutivo = rs("ejecutivo")
			sEstadoTRF = rs("estado")
			sTiempo = rs("tiempo")
			sMonedaEquivalente = rs("monedaequivalente")
		end if
		set rs = nothing
	
		' graba la transferencia y una solicitud
		'**** Solicitud
		set dbConexion = server.CreateObject("ADODB.Connection")
		dbConexion.open session("afxCnxAFEXchange")
		if err.number <> 0 then
			response.Write "Ocurrió un error al abrir la conexión. " & err.Description
			response.End
		end if	
		
		' verifica si el cliente ya tiene una solicitud		
		sSQL = "select s.codigo_solicitud, " & _
				"isnull((select count(*) from detalle_solicitud ds where ds.codigo_solicitud=s.codigo_solicitud),0) + 1 as linea " & _
				"from   solicitud s " & _
					" inner join cliente c on c.codigo_cliente = s.codigo_cliente and " & _
										 " c.codigo_corporativa = " & Session("CodigoCliente") & _
				"where  s.fecha_solicitud = " & evaluarstr(date) & _
					 " and s.estado_solicitud = 1 "	' ejecutivo
		set rs = dbConexion.execute(sSQL)
		if err.number <> 0 then			
			response.Write "Ocurrió un error al buscar una solicitud existente. 1" & err.Description
			response.End
		end if
		if not rs.eof then
			sSolicitud = rs("codigo_solicitud")
			sSolicitudOriginal = rs("codigo_solicitud")
			sLineaSolicitud = rs("linea")
		end if
		set rs = nothing	
		
		' saca el numero de voucher del primero
	'	sSQL = "select isnull(correlativo_actual, 0) as voucher " & _
	'			"from   correlativo_documento " & _
	'			"where  tipo_documento = 3 and " & _
	'				" tipo_operacion = 5 AND " & _
	'				" codigo_caja = '0000'"
	'	set rs = dbConexion.execute(sSQL)
	'	if err.number <> 0 then
	'		dbConexion.rollbacktrans
	'		response.Write "Ocurrió un error al buscar el correlativo de voucher. 1" & err.Description
	'		response.End
	'	end if
	'	if rs.eof then
	'		dbConexion.rollbacktrans
	'		response.Write "No hay correlativo para el voucher. 1" & err.Description
	'		response.End
	'	end if
	'	cVoucher = rs("voucher")
	'	set rs = nothing
		' actualiza el correlativo
	'	sSQL = "update correlativo_documento set correlativo_actual = correlativo_actual + 1 " & _
	'			"where  tipo_documento = 3 and " & _
	'				"tipo_operacion = 5 AND " & _
	'				"codigo_caja = '0000'"
	'	set rs = dbConexion.execute(sSQL)
	'	if err.number <> 0 then
	'		dbConexion.rollbacktrans
	'		response.Write "Ocurrió un error al actualizar el correlativo de voucher. 1" & err.Description
	'		response.End
	'	end if
	'	set rs = nothing
		
		dbConexion.begintrans		
				
		' verifica si existe la solicitud
		if trim(sSolicitud) = "" then
			sSQL = "InsertarSolicitud " & evaluarstr(sejecutivo) & ", " & evaluarstr(Session("CodigoCliente")) & ", " & _
						evaluarstr(date) & ", 'USD', 0, 0, 0, 0, 0, 0, " & formatonumerosql(ccur(trim(request.Form("txtMontoPesos")))) & ", " & _
						"0, 0, null, 0, 4, 1, 0, 1, 0, 0, 0, 2, null, '0000', " & evaluarstr(sEjecutivo) & ", " & _
						"'20071231', " & evaluarstr(Session("NombreOperador")) & ", 0, 0, null, null, null, null, " & _
						"null, 0, null, null, 1, " & formatonumerosql(ccur(trim(request.Form("txtMontoPesos")))) & ", null "
			set rs = dbConexion.execute(sSQL)
			if err.number <> 0 then
				dbConexion.rollbacktrans
				response.Write "Ocurrió un error al grabar la solicitud. " & err.Description
				response.End
			end if
			if err.number <> 0 then
				dbConexion.rollbacktrans
				response.Write "No se generó una solicitud. "
				response.End
			end if
			sSolicitud = rs("solicitud")
			set rs = nothing		
		end if
		
		'**** DetalleSolicitud
		sGlosa = sBancoOrigen & ";BANCO ORIGEN;" & sCuentaOrigen & ";;;;;;;;;;48;;" & date & ";;" & date & ";" & date & ";;;;;;"
		sSQL = "InsertarDetalleSolicitud " & sSolicitud & ", 2, " & sLineaSolicitud & ", null, 2, 'USD', 3, 1, 1, 2, null, " & _ 
					formatonumerosql(ccur(trim(request("Equivalente")))) & ", " & formatonumerosql(ccur(trim(request.Form("txtTipoCambio")))) & ", " & _
					formatonumerosql(ccur(trim(request.Form("txtMontoPesos")))) & ", 0, " & formatonumerosql(ccur(trim(request.Form("txtMontoPesos")))) & ", " & _
					"0, 0, 0, null, null, null, null, 1, 1, " & evaluarstr(sGlosa) & ", null, null, 0"						
		
		set rs = dbConexion.execute(sSQL)
		if err.number <> 0 then
			dbConexion.rollbacktrans
			response.Write "Ocurrió un error al grabar el detalle de solicitud. " & err.Description
			response.End
		end if		
		set rs = nothing
		' actualiza la fila 1 de la solicitud
		sSQL = ""
		sSQL = " update detalle_solicitud " & _
				" set monto_extranjera = monto_extranjera - " & formatonumerosql(ccur(trim(request("Equivalente")))) & ", " & _
					" monto_nacional = monto_nacional - " & formatonumerosql(ccur(trim(Request.form("txtmontopesos")))) & ", " & _
					" total_nacional = total_nacional - " & formatonumerosql(ccur(trim(Request.form("txtmontopesos")))) & _
				" where codigo_solicitud = " & sSolicitud & _
					" and numero_linea = 1 "					
		set rs = dbConexion.execute(sSQL)
		if err.number <> 0 then
			dbConexion.rollbacktrans
			response.Write "Ocurrió un error al actualizar el detalle de solicitud. " & err.Description
			response.End
		end if		
		set rs = nothing
		' elimina el detalle 1 si es que esta en 0
		sSQL = " delete detalle_solicitud " & _				
				" where codigo_solicitud = " & sSolicitud & _
					" and numero_linea = 1 " & _
					" and monto_extranjera <= 0 "
		set rs = dbConexion.execute(sSQL)
		if err.number <> 0 then
			dbConexion.rollbacktrans
			response.Write "Ocurrió un error al eliminar el detalle de solicitud. " & err.Description
			response.End
		end if		
		set rs = nothing
		
		'**** TRANSFER
		if trim(request.Form("cbxBancodestino")) = "" then
			sCodigoBanco = "null"
		else
			trim(request.Form("cbxBancodestino")) = sCodigoBanco
		end if
		sSQL = "InsertarTransferencia " & sBancoOrigen & ", " & evaluarstr(Session("CodigoCliente")) & ", " & _ 
					formatonumerosql(ccur(trim(request("Monto")))) & ", " & evaluarstr(date) & ", " & sEstadoTRF & ", " & sSolicitud & ", " & _ 
					"2, 1, null, " & evaluarstr(sCuentaOrigen) & ", " & sCodigoBanco & ", " & evaluarstr(trim(request.Form("txtBanco"))) & ", " & _
					evaluarstr(trim(request.Form("txtCtaCte"))) & ", " & evaluarstr(trim(request.Form("txtNombreExacto"))) & ", " & _
					evaluarstr(trim(request.Form("txtABA"))) & ", null, " & evaluarstr(trim(request.Form("txtCiudadBanco"))) & ", " & _
					evaluarstr(trim(request.Form("txtDireccionBanco"))) & ", " & _ 
					evaluarstr(sTiempo) & ", null, " & evaluarstr(trim(request("CodigoMoneda"))) & ", " & evaluarstr(sMonedaEquivalente) & ", " & _ 
					formatonumerosql(ccur(trim(request("Rate")))) & ", " & formatonumerosql(ccur(trim(request("Equivalente")))) & ", " & _ 
					evaluarstr(Session("NombreUsuarioOperador")) & ", null, 1, " & formatonumerosql(ccur(trim(request("Tarifa")))) & ", " & formatonumerosql(ccur(trim(request("Tarifa")))) & ", " & _ 
					"0, " & evaluarstr(trim(request.Form("txtBancoInt"))) & ", " & evaluarstr(trim(request.Form("txtCtaCteInt"))) & ", " & _ 
					evaluarstr(trim(request.Form("txtCiudadBancoInt"))) & ", " & evaluarstr(trim(request.Form("txtDireccionBancoInt"))) & ", " & _
					"1, null, null, null, " & evaluarstr(Request.Form("txtinvoice")) & ", null, " & evaluarstr(Request.Form("txtdireccionb"))		
	
'	Response.Write ssql
'	Response.end
	
		set rs = dbConexion.execute(sSQL)
		if err.number <> 0 then
			dbConexion.rollbacktrans
			response.Write "Ocurrió un error al grabar la transferencia. " & err.Description
			response.End
		end if
		set rs = nothing
		
		'***** Vouchers
		'Voucher 1	
		' Encabezado 1
		'if trim(sSolicitudOriginal) = "" then
		'	sSQL = "INSERT INTO voucher (numero_voucher, codigo_solicitud, codigo_plantilla, codigo_entidad, fecha_ingreso, " & _
		'								" tipo_voucher, codigo_moneda, glosa_voucher, monto_extranjera, tipo_cambio, monto_nacional, " & _
		'								" codigo_negocio, tipo_operacion, codigo_operacion, centralizacion_contabilidad, centralizacion_matriz, " & _
		'								" estado_voucher, codigo_usuario, hora_ingreso, correlativo_transferencia) " & _
		'						" VALUES  (" & evaluarstr(cvoucher) & ", " & ssolicitud & ", 2, null, " & evaluarstr(date) & ", 1, 'CLP', " & _
		'								EvaluarStr(Session("NombreOperador")) & ", 0, 0, " & formatonumerosql(ccur(trim(request.Form("txtMontoPesos")))) & ", " & _
		'								"1, 2, null, 1, 0, 1, " & EvaluarStr(sEjecutivo) & ", " & EvaluarStr(replace(time(), ":", "")) & ", Null)"
		'	set rs = dbConexion.execute(sSQL)
		'	if err.number <> 0 then
		'		dbConexion.rollbacktrans
		'		response.Write "Ocurrió un error algrabar el voucher. 1" & err.Description
		'		response.End
		'	end if
		'	set rs = nothing
		'		'detalle del voucher   Linea 1
		'		sSQL = "INSERT INTO detalle_voucher (numero_voucher, numero_linea, numero_cuenta, codigo_producto, numero_producto, " & _
		'										" debe_extranjera, haber_extranjera, debe_nacional, haber_nacional, tipo_documento, " & _
		'										" numero_documento, tipo_operacion, glosa_linea, codigo_item) " & _
		'							" VALUES  (" & evaluarstr(cvoucher) & ", 1, 210603001, 3, null, 0, " & formatonumerosql(ccur(trim(request("Equivalente")))) & ", " & _
		'										"0, " & formatonumerosql(ccur(trim(request.Form("txtMontoPesos")))) & ", 2, null, 2, null, 0)"
		'		set rs = dbConexion.execute(sSQL)
		'		if err.number <> 0 then
		'			dbConexion.rollbacktrans
		'			response.Write "Ocurrió un error al grabar el detalle de voucher. 1.1" & err.Description
		'			response.End
		'		end if
		'		set rs = nothing
		'		' Linea 2
		'		sSQL = "INSERT INTO detalle_voucher (numero_voucher, numero_linea, numero_cuenta, codigo_producto, numero_producto, " & _
		'										" debe_extranjera, haber_extranjera, debe_nacional, haber_nacional, tipo_documento, " & _
		'										" numero_documento, tipo_operacion, glosa_linea, codigo_item) " & _
		'							" VALUES  (" & evaluarstr(cvoucher) & ", 2, 910102001, 3, null, " & formatonumerosql(ccur(trim(request("Equivalente")))) & ", 0, " & _
		'										"0, 0, 2, null, 2, null, 0)"
		'		set rs = dbConexion.execute(sSQL)
		'		if err.number <> 0 then
		'			dbConexion.rollbacktrans
		'			response.Write "Ocurrió un error al grabar el detalle de voucher. 1.2" & err.Description
		'			response.End
		'		end if
		'		set rs = nothing
		'		' Linea 3
		'		sSQL = "INSERT INTO detalle_voucher (numero_voucher, numero_linea, numero_cuenta, codigo_producto, numero_producto, " & _
		'										" debe_extranjera, haber_extranjera, debe_nacional, haber_nacional, tipo_documento, " & _
		'										" numero_documento, tipo_operacion, glosa_linea, codigo_item) " & _
		'							" VALUES  (" & evaluarstr(cvoucher) & ", 3, 110101001, null, null, 0, 0, " & _
		'										formatonumerosql(ccur(trim(request.Form("txtMontoPesos")))) & ", 0, null, null, 2, null, 0)"
		'		set rs = dbConexion.execute(sSQL)
		'		if err.number <> 0 then
		'			dbConexion.rollbacktrans
		'			response.Write "Ocurrió un error al grabar el detalle de voucher. 1.3" & err.Description
		'			response.End
		'		end if
		'		set rs = nothing
		'end if
				
		''Voucher 2
		'' saca el numero de voucher
		'sSQL = "select isnull(correlativo_actual, 0) as voucher " & _
		'		"from   correlativo_documento " & _
		'		"where  tipo_documento = 3 and " & _
		'			" tipo_operacion = 5 AND " & _
		'			" codigo_caja = '0000'"
		'set rs = dbConexion.execute(sSQL)
		'if err.number <> 0 then
		'	dbConexion.rollbacktrans
		'	response.Write "Ocurrió un error al buscar el correlativo de voucher. 2" & err.Description
		'	response.End
		'end if
		'if rs.eof then
		'	dbConexion.rollbacktrans
		'	response.Write "No hay correlativo para el voucher. 2" & err.Description
		'	response.End
		'end if
		'cVoucher = rs("voucher")
		'set rs = nothing
		' actualiza el correlativo
		'sSQL = "update correlativo_documento set correlativo_actual = correlativo_actual + 1 " & _
		'		"where  tipo_documento = 3 and " & _
		'			"tipo_operacion = 5 AND " & _
		'			"codigo_caja = '0000'"
		'set rs = dbConexion.execute(sSQL)
		'if err.number <> 0 then
		'	dbConexion.rollbacktrans
		'	response.Write "Ocurrió un error al actualizar el correlativo de voucher. 2" & err.Description
		'	response.End
		'end if
		'set rs = nothing
		
		' Encabezado 2
		'sSQL = "INSERT INTO voucher (numero_voucher, codigo_solicitud, codigo_plantilla, codigo_entidad, fecha_ingreso, " & _
		'							" tipo_voucher, codigo_moneda, glosa_voucher, monto_extranjera, tipo_cambio, monto_nacional, " & _
		'							" codigo_negocio, tipo_operacion, codigo_operacion, centralizacion_contabilidad, centralizacion_matriz, " & _
		'							" estado_voucher, codigo_usuario, hora_ingreso, correlativo_transferencia) " & _
		'					" VALUES  (" & evaluarstr(cvoucher) & ", " & ssolicitud & ", null, null, " & evaluarstr(date) & ", 3, 'USD', " & _
		'							EvaluarStr("TRF, " & Session("NombreOperador")) & ", " & formatonumerosql(ccur(trim(request("Equivalente")))) & ", " & _
		'							formatonumerosql(ccur(trim(request.Form("txtTipoCambio")))) & ", " & formatonumerosql(ccur(trim(request.Form("txtMontoPesos")))) & ", " & _
		'							"1, 10, null, 1, 0, 1, " & EvaluarStr(sEjecutivo) & ", " & EvaluarStr(replace(time(), ":", "")) & ", Null)"
		'set rs = dbConexion.execute(sSQL)
		'if err.number <> 0 then
		'	dbConexion.rollbacktrans
		'	response.Write "Ocurrió un error algrabar el voucher. 2" & err.Description
		'	response.End
		'end if
		'set rs = nothing
			'detalle del voucher   Linea 1
		'	sSQL = "INSERT INTO detalle_voucher (numero_voucher, numero_linea, numero_cuenta, codigo_producto, numero_producto, " & _
		'									" debe_extranjera, haber_extranjera, debe_nacional, haber_nacional, tipo_documento, " & _
		'									" numero_documento, tipo_operacion, glosa_linea, codigo_item) " & _
		'						" VALUES  (" & evaluarstr(cvoucher) & ", 1, 510101010, null, null, " & formatonumerosql(ccur(trim(request("Equivalente")))) & ", 0, " & _
		'									formatonumerosql(ccur(trim(request.Form("txtMontoPesos")))) & ", 0, null, null, 10, " & _
		'									EvaluarStr("TRF, " & Session("NombreOperador")) & ", 1203)"
		'	set rs = dbConexion.execute(sSQL)
		'	if err.number <> 0 then
		'		dbConexion.rollbacktrans
		'		response.Write "Ocurrió un error al grabar el detalle de voucher. 1.1" & err.Description
		'		response.End
		'	end if
		'	set rs = nothing
			' Linea 2
		'	sSQL = "INSERT INTO detalle_voucher (numero_voucher, numero_linea, numero_cuenta, codigo_producto, numero_producto, " & _
		'									" debe_extranjera, haber_extranjera, debe_nacional, haber_nacional, tipo_documento, " & _
		'									" numero_documento, tipo_operacion, glosa_linea, codigo_item) " & _
		'						" VALUES  (" & evaluarstr(cvoucher) & ", 2, 110502002, null, null, 0, " & formatonumerosql(ccur(trim(request("Equivalente")))) & ", " & _
		'									"0, " & formatonumerosql(ccur(trim(request.Form("txtMontoPesos")))) & ", null, null, 10, " & _
		'									EvaluarStr("TRF, " & Session("NombreOperador")) & ", 0)"
		'	set rs = dbConexion.execute(sSQL)
		'	if err.number <> 0 then
		'		dbConexion.rollbacktrans
		'		response.Write "Ocurrió un error al grabar el detalle de voucher. 1.2" & err.Description
		'		response.End
		'	end if
		'	set rs = nothing		
			
		dbConexion.committrans
		
		set rs = nothing
		set dbConexion = nothing
		' ********************* Fin ********************
	
		sAsunto = "Enviar Transferencia"
		sDescripcion = sDescripcion & "desea enviar una transferencia con el siguiente detalle: " & _
							vbCrlf & vbCrlf 
		
		sDescripcion = sDescripcion & vbCrlf & "VALUTA  : " & trim(request("Valuta"))
		sDescripcion = sDescripcion & vbCrlf & "RATE    : " & trim(request("Rate"))		
		sDescripcion = sDescripcion & vbCrlf & "MONEDA  : " & trim(request("Moneda"))
		sDescripcion = sDescripcion & vbCrlf & "MONTO   : " & trim(request("Monto"))
		sDescripcion = sDescripcion & vbCrlf & "EQUIVALENTE US$ : " & trim(request("Equivalente"))
		sDescripcion = sDescripcion & vbCrlf & "TARIFA  : " & trim(request("Tarifa"))
		sDescripcion = sDescripcion & vbCrlf & "TOTAL   : " & trim(request("Total"))		
		sDescripcion = sDescripcion & vbCrlf & "TIPO CAMBIO : " & trim(request.Form("txtTipoCambio"))
		sDescripcion = sDescripcion & vbCrlf & "MONTO PESOS : " & trim(request.Form("txtMontoPesos"))
		sDescripcion = sDescripcion & vbCrlf & vbCrlf & "BANCO DESTINO "
		sDescripcion = sDescripcion & vbCrlf & "NOMBRE : " & trim(request.Form("txtBanco"))
		sDescripcion = sDescripcion & vbCrlf & "Nº CTA.CTE : " & trim(request.Form("txtCtaCte"))
		sDescripcion = sDescripcion & vbCrlf & "NOMBRE Cta. DEL BENEFICIARIO : " & trim(request.Form("txtNombreExacto"))
		sDescripcion = sDescripcion & vbCrlf & "DIRRECCIÓN : " & trim(request.Form("txtDireccionBanco"))
		sDescripcion = sDescripcion & vbCrlf & "CIUDAD : " & trim(request.Form("txtCiudadBanco"))
		sDescripcion = sDescripcion & vbCrlf & "ABA : " & trim(request.Form("txtABA"))
		sDescripcion = sDescripcion & vbCrlf & "CHIPS : " & trim(request.Form("txtCHIPS"))
		sDescripcion = sDescripcion & vbCrlf & "SWIFT : " & trim(request.Form("txtSWIFT"))
		sDescripcion = sDescripcion & vbCrlf & "IBAN : " & trim(request.Form("txtIBAN"))
		sDescripcion = sDescripcion & vbCrlf & "OTRO : " & trim(request.Form("txtOTRO"))
		sDescripcion = sDescripcion & vbCrlf & vbCrlf & "BANCO INTERMEDIARIO "
		sDescripcion = sDescripcion & vbCrlf & "NOMBRE : " & trim(request.Form("txtBancoInt"))
		sDescripcion = sDescripcion & vbCrlf & "Nº CTA.CTE : " & trim(request.Form("txtCtaCteInt"))
		sDescripcion = sDescripcion & vbCrlf & "NOMBRE Cta. DEL BENEFICIARIO : " & trim(request.Form("txtNombreExactoInt"))
		sDescripcion = sDescripcion & vbCrlf & "DIRRECCIÓN : " & trim(request.Form("txtDireccionBancoInt"))
		sDescripcion = sDescripcion & vbCrlf & "CIUDAD : " & trim(request.Form("txtCiudadBancoInt"))
		sDescripcion = sDescripcion & vbCrlf & "ABA : " & trim(request.Form("txtABAInt"))
		sDescripcion = sDescripcion & vbCrlf & "CHIPS : " & trim(request.Form("txtCHIPSInt"))
		sDescripcion = sDescripcion & vbCrlf & "SWIFT : " & trim(request.Form("txtSWIFTInt"))
		sDescripcion = sDescripcion & vbCrlf & "IBAN : " & trim(request.Form("txtIBANInt"))
		sDescripcion = sDescripcion & vbCrlf & "OTRO : " & trim(request.Form("txtOTROInt"))
	
	Case 7	' crear cliente
		Dim sID
		
		sID = Request.Form("txtID")
		If Request.Form("cbxID") = 1 Then
			sID = sID & " RUT" 
		else
			sID = sID & " PASAPORTE" 
		End if
		sAsunto = "Crear Cliente"
		sDescripcion = "Desde la página web se está registrando el siguiente Cliente: " & _
							vbCrlf & vbCrlf 
		
		
		sDescripcion = sDescripcion & vbCrlf & "ID : " & trim(sID)
		sDescripcion = sDescripcion & vbCrlf & "NOMBRE : " & trim(request.Form("txtPrimerNombre")) & " " & trim(request.Form("txtSegundoNombre"))
		sDescripcion = sDescripcion & vbCrlf & "APELLIDOS : " & trim(request.Form("txtApellidoPaterno")) & " " & trim(request.Form("txtApellidoMaterno"))
		sDescripcion = sDescripcion & vbCrlf & "DIRECCIÓN : " & trim(request.Form("txtDireccion"))
		sDescripcion = sDescripcion & vbCrlf & "EMAIL : " & trim(request.Form("txtEmail"))
		sDescripcion = sDescripcion & vbCrlf & "TELÉFONO : " & trim(request.Form("txtTelefono"))
		sDescripcion = sDescripcion & vbCrlf & "CELULAR : " & trim(request.Form("txtCelular"))		
		sDescripcion = sDescripcion & vbCrlf & "NOMBRE USUARIO : " & trim(request.Form("txtUsuario"))
		sDescripcion = sDescripcion & vbCrlf & "CLAVE : " & trim(request.Form("txtClave"))
		sDescripcion = sDescripcion & vbCrlf & "PREGUNTA : " & trim(request.Form("txtPreguntaSecreta"))
		sDescripcion = sDescripcion & vbCrlf & "RESPUESTA : " & trim(request.Form("txtRespuesta"))		
End Select

' envia el mail
EnviarEmail "AFEX", sMailEjecutivo, Session("emailcliente"), sAsunto, sDescripcion, 0

If err.number <> 0 Then
	response.Write "Ocurrió un error al enviar el mail. " & err.Description	
	response.End
else
	Select Case Request("Accion")
		Case 6	' Transfer
			response.redirect "enviartransfer.asp?Mensaje=Los datos han sido enviados. Gracias por operar con nosotros."
	end select
End If

%>