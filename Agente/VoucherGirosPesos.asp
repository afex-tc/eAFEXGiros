<%@ Language=VBScript %>
<%
	'Se asegura que la página no se almacene en la memoria cache
	Response.expires = 0
	If Session("CodigoCliente") = "" Then
		response.Redirect "../Compartido/TimeOut.htm"
		response.end
	End If
%>
<!-- #Include virtual="/Compartido/Rutinas.asp"-->
<!-- #Include virtual="/Compartido/Errores.asp"-->

<%

	dim sSQL, rs, sFecha, sEncabezadoFondo, sEncabezadoTitulo
	dim sConexionSucursal, sBDSucursal, cVoucher, iEstado, cMontoNacional, sHora, i
	dim nNumeroCuenta, cDebeNacional, cHaberNacional, sGlosaLinea, iCodigoItem
	dim sMensajeError
	
	sFecha = Request.Form("txtFecha")
	if sFecha = "" then sFecha = date
	
	' verifica la acción solicitada
	if request("Accion") = 1 then	' generar voucher en AFEXchange de la sucursal
		' saca la conexiín de la sucursal
		sSQL = "select ip, nombrebd from sucursales where codigoagente = " & evaluarstr(Session("CodigoAgente"))
		set rs = ejecutarsqlcliente(Session("afxCnxCorporativa"), sSQL)
		if err.number <> 0 then
			MostrarErrorMS "Buscar Conexión Sucursal"
		end if
		if rs.eof then
			set rs = nothing
			MostrarErrorMS "No se encontró la Conexión."
		end if
		' arma la conexión
		sConexionSucursal = "Provider=SQLOLEDB.1;Password=cambios;User ID=cambios;" 
		if rs("nombrebd") <> "" then 
			sConexionSucursal = sConexionSucursal & "Initial Catalog=" & trim(rs("nombrebd")) & ";Data Source=" & rs("ip") & ";"
		else
			sConexionSucursal = sConexionSucursal & "Initial Catalog=cambios;Data Source=" & trim(rs("ip")) & ";"
		end if
		set rs = nothing
		
		' verifica si el voucher ya existe
		sSQL = " select count(*) as Cantidad " & _
				" from voucher " & _
				" where glosa_voucher = 'Giros en Pesos .WEB' " & _
					" and fecha_ingreso = " & evaluarstr(formatofechasql(date))
		set rs = ejecutarsqlcliente(sConexionSucursal, sSQL)
		if err.number <> 0 then
			MostrarErrorMS "Buscar Voucher Cierre. "
		end if
		if not rs.eof then
			if rs("cantidad") > 0 then
				sMensaje = "El voucher ya existe."
					
			else	' el voucher no existe	
			
				' consulta los giros para el día seleccionado y generar el voucher
				sSQL = " exec mostrargirospesossucursal " & evaluarstr(Session("CodigoAgente")) & ", " & evaluarstr(formatofechasql(sFecha))	
				set  rsGirosVoucher = ejecutarsqlcliente(Session("afxCnxAFEXpress"), sSQL)
				if err.number <> 0 then
					MostrarErrorMS "Consultar Giros"
				end if
				if rsGirosVoucher.eof then
					set rsGirosVoucher = nothing
					MostrarErrorMS "No se encontraron Giros."
				end if
		
				'** Genera el voucher
					' saca el correlativo
				sSQL = " select isnull(correlativo_actual, 0) as correlativovoucher " & _
					   " from correlativo_documento " & _
					   " where codigo_caja='0000' and tipo_operacion = 6 "
				set rs = ejecutarsqlcliente(sConexionSucursal, sSQL)
				if err.number <> 0 then
					MostrarErrorMS "Obtener Correlativo Voucher."
				end if
				if rs.eof then
					MostrarErrorMS "No se encontraró Correlativo Voucher."
				end if
				cVoucher = rs("correlativovoucher")
					' actualiza el correlativo del voucher
				sSQL = " update correlativo_documento " & _
					   "	set correlativo_actual = correlativo_actual + 1" & _
					   " where codigo_caja='0000' and tipo_operacion = 6 "
				set rs = ejecutarsqlcliente(sConexionSucursal, sSQL)
				if err.number <> 0 then
					MostrarErrorMS "Actualizar Correlativo Voucher."
				end if
				set rs = nothing
		
				iEstado = 1
				cMontoNacional = 0
				sHora = right("00" & replace(time, ":", ""),6)
		
					' encabezado
				sSQL = " insert into voucher (numero_voucher, codigo_plantilla, codigo_entidad, fecha_ingreso, tipo_voucher, codigo_moneda, glosa_voucher, " & _
										" monto_nacional, tipo_cambio, monto_extranjera, codigo_negocio, tipo_operacion, estado_voucher, centralizacion_matriz, " & _
										" codigo_usuario, hora_ingreso) " & _
									" values (" & cVoucher & ", 0, 0, " & evaluarstr(formatofechasql(date)) & ", 3, 'CLP', 'Giros en Pesos .WEB', " & _
										replace(ccur(cMontoNacional), ",", ".") & ", 0, 0, 1, 10, " & iEstado & ", 0, " & _
										evaluarstr(Session("NombreUsuarioOperador")) & ", " & sHora & ") "
				set rs = ejecutarsqlcliente(sConexionSucursal, sSQL)
				if err.number <> 0 then
					MostrarErrorMS "Encabezado Voucher."
				end if
		
					'detalle
				i = 0
				do while not rsGirosVoucher.eof
					' inicializa las variables
					nNumeroCuenta = 0
					cDebeNacional = 0
					cHaberNacional = 0
					sGlosaLinea = ""
					iCodigoItem = 0
		
					' crea la línea de detalle
					if rsGirosVoucher("tipogiro") = 1 then	' giro pagado por la sucursal
						' linea 1
						sGlosaLinea = "PAGO Giro : " & rsGirosVoucher("codigo_giro")
						nNumeroCuenta = 110501026 'Cta. Cte. Giros
						iCodigoItem = 0
						cDebeNacional = ccur(rsGirosVoucher("monto_giro")) + ccur(rsGirosVoucher("comisionpagador"))
						cHaberNacional = 0	
						
						cMontoNacional = ccur(cMontoNacional) + ccur(cDebeNacional)
						
						i = i + 1
						sSQL = " insert into detalle_voucher (numero_voucher, numero_linea, numero_cuenta, tipo_operacion, debe_nacional, haber_nacional, glosa_linea, codigo_item) " & _
									" values (" & cVoucher & ", " & i & ", " & nNumeroCuenta & ", 10, " & cDebeNacional & ", " & cHaberNacional & ", " & evaluarstr(sGlosaLinea) & ", " & iCodigoItem & ") "
						' ejecuta el script con tres lienas, las que corresponden a un giro			
						set rs = ejecutarsqlcliente(sConexionSucursal, sSQL)
						if err.number <> 0 then				
							sMensajeError = err.Description 
							EliminarVoucher cVoucher, sConexionSucursal
							MostrarErrorMS "Detalle Voucher. 1" & sMensajeError
						end if
						
						' linea 2				
						nNumeroCuenta = 110101001 'Caja Pesos
						iCodigoItem = 0
						cDebeNacional = 0
						cHaberNacional = rsGirosVoucher("monto_giro")
						
						i = i + 1
						sSQL = " insert into detalle_voucher (numero_voucher, numero_linea, numero_cuenta, tipo_operacion, debe_nacional, haber_nacional, glosa_linea, codigo_item) " & _
									" values (" & cVoucher & ", " & i & ", " & nNumeroCuenta & ", 10, " & cDebeNacional & ", " & cHaberNacional & ", " & evaluarstr(sGlosaLinea) & ", " & iCodigoItem & ") "
						' ejecuta el script con tres lienas, las que corresponden a un giro			
						set rs = ejecutarsqlcliente(sConexionSucursal, sSQL)
						if err.number <> 0 then				
							sMensajeError = err.Description 
							EliminarVoucher cVoucher, sConexionSucursal
							MostrarErrorMS "Detalle Voucher. 2" & sMensajeError
						end if
						
						
						' linea 3
						nNumeroCuenta = 610103001 'Otros Ingresos Operacionales
						iCodigoItem = 6359
						cDebeNacional = 0
						cHaberNacional = rsGirosVoucher("comisionpagador")
						
						i = i + 1
						sSQL = " insert into detalle_voucher (numero_voucher, numero_linea, numero_cuenta, tipo_operacion, debe_nacional, haber_nacional, glosa_linea, codigo_item) " & _
									" values (" & cVoucher & ", " & i & ", " & nNumeroCuenta & ", 10, " & cDebeNacional & ", " & cHaberNacional & ", " & evaluarstr(sGlosaLinea) & ", " & iCodigoItem & ") "
						' ejecuta el script con tres lienas, las que corresponden a un giro			
						set rs = ejecutarsqlcliente(sConexionSucursal, sSQL)
						if err.number <> 0 then				
							sMensajeError = err.Description 
							EliminarVoucher cVoucher, sConexionSucursal
							MostrarErrorMS "Detalle Voucher. 3" & sMensajeError
						end if					
						
						
					elseif rsGirosVoucher("tipogiro") = 2 then	' giro enviado por la sucursal
						' linea 1
						sGlosaLinea = "ENVIO Giro : " & rsGirosVoucher("codigo_giro")
						
						nNumeroCuenta = 110101001 'Caja Pesos
						iCodigoItem = 0
						cDebeNacional = ccur(rsGirosVoucher("monto_giro")) + ccur(rsGirosVoucher("tarifacobrada"))
						cHaberNacional = 0
						
						cMontoNacional = ccur(cMontoNacional) + ccur(cDebeNacional)
						
						i = i + 1
						sSQL = " insert into detalle_voucher (numero_voucher, numero_linea, numero_cuenta, tipo_operacion, debe_nacional, haber_nacional, glosa_linea, codigo_item) " & _
									" values (" & cVoucher & ", " & i & ", " & nNumeroCuenta & ", 10, " & cDebeNacional & ", " & cHaberNacional & ", " & evaluarstr(sGlosaLinea) & ", " & iCodigoItem & ") "
						' ejecuta el script con tres lienas, las que corresponden a un giro			
						set rs = ejecutarsqlcliente(sConexionSucursal, sSQL)
						if err.number <> 0 then				
							sMensajeError = err.Description 
							EliminarVoucher cVoucher, sConexionSucursal
							MostrarErrorMS "Detalle Voucher. 4" & sMensajeError
						end if
						
						
						' linea 2
						nNumeroCuenta = 110501026 'Cta. Cte. Giros
						iCodigoItem = 0
						cDebeNacional = 0
						cHaberNacional = (ccur(rsGirosVoucher("monto_giro")) + ccur(rsGirosVoucher("tarifacobrada"))) - ccur(rsGirosVoucher("comisioncaptador"))				
						
						i = i + 1
						sSQL = " insert into detalle_voucher (numero_voucher, numero_linea, numero_cuenta, tipo_operacion, debe_nacional, haber_nacional, glosa_linea, codigo_item) " & _
									" values (" & cVoucher & ", " & i & ", " & nNumeroCuenta & ", 10, " & cDebeNacional & ", " & cHaberNacional & ", " & evaluarstr(sGlosaLinea) & ", " & iCodigoItem & ") "
						' ejecuta el script con tres lienas, las que corresponden a un giro			
						set rs = ejecutarsqlcliente(sConexionSucursal, sSQL)
						if err.number <> 0 then				
							sMensajeError = err.Description 
							EliminarVoucher cVoucher, sConexionSucursal
							MostrarErrorMS "Detalle Voucher. 5" & sMensajeError
						end if
						
						
						' linea 3
						nNumeroCuenta = 610102001 'Comisiones Ganadas
						iCodigoItem = 6215
						cDebeNacional = 0
						cHaberNacional = rsGirosVoucher("comisioncaptador")
						
						i = i + 1
						sSQL = " insert into detalle_voucher (numero_voucher, numero_linea, numero_cuenta, tipo_operacion, debe_nacional, haber_nacional, glosa_linea, codigo_item) " & _
									" values (" & cVoucher & ", " & i & ", " & nNumeroCuenta & ", 10, " & cDebeNacional & ", " & cHaberNacional & ", " & evaluarstr(sGlosaLinea) & ", " & iCodigoItem & ") "
						' ejecuta el script con tres lienas, las que corresponden a un giro			
						set rs = ejecutarsqlcliente(sConexionSucursal, sSQL)
						if err.number <> 0 then
							sMensajeError = err.Description 
							EliminarVoucher cVoucher, sConexionSucursal
							MostrarErrorMS "Detalle Voucher. 6" & sMensajeError
						end if
					end if
					
					rsGirosVoucher.MoveNext
				loop		
		
				' actualiza el monto del encabezado
				sSQL = " update voucher " & _
						" set monto_nacional = " & ccur(cMontoNacional) &  _
						" where numero_voucher = " & cVoucher
				set rs = ejecutarsqlcliente(sConexionSucursal, sSQL)
				if err.number <> 0 then
					sMensajeError = err.Description 
					EliminarVoucher cVoucher, sConexionSucursal
					MostrarErrorMS "Actualizar Monto Voucher. 6" & sMensajeError					
				end if
		
				sMensaje = "El Voucher se generó Correctamente."
			end if
		end if
	end if
	
	' consulta los giros para el día seleccionado
	sSQL = " exec mostrargirospesossucursal " & evaluarstr(Session("CodigoAgente")) & ", " & evaluarstr(formatofechasql(sFecha))	
	set rs = ejecutarsqlcliente(Session("afxCnxAFEXpress"), sSQL)
	if err.number <> 0 then
		MostrarErrorMS "Consultar Giros"
	end if
	
	sEncabezadoFondo = "Consultas"
	sEncabezadoTitulo = "Voucher de giros en pesos para AFEXchange"

	sub EliminarVoucher(ByVal Voucher, ByVal ConexionSucursal)
		dim sSQL, rs
		
		' elimina detalle
		sSQL = " delete detalle_voucher where numero_voucher = " & Voucher 
		set rs = ejecutarsqlcliente(ConexionSucursal, sSQL)
		if err.number <> 0 then
			MostrarErrorMS "Eliminar Detalle Voucher."
		end if
		
		' elimina encabezado
		sSQL = " delete voucher where numero_voucher = " & Voucher
		set rs = ejecutarsqlcliente(ConexionSucursal, sSQL)
		if err.number <> 0 then
			MostrarErrorMS "Eliminar Encabezado Voucher."
		end if
		
		set rs = nothing
	end sub

	Response.Expires = 0
%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<title></title>
<link rel="stylesheet" type="text/css" href="../Estilos/Reportes.css">
<link rel="stylesheet" type="text/css" href="../Estilos/Cliente.css">
</HEAD>
<BODY>
	
	<script language="vbscript">
	<!--
		sub window_onload()
			if "<%=sMensaje%>" <> "" then
				msgbox "<%=sMensaje%>", ,"AFEX" 
			end if
		end sub
	
		Sub imgAceptar_onClick()
			frm.action = "VoucherGirosPesos.asp"
			frm.submit()
			frm.action = ""
		End Sub
		
		Sub cmdGenerar_onClick()
			frm.action = "VoucherGirosPesos.asp?Accion=1"
			frm.submit()
			frm.action = ""
		End Sub
		
	-->
	</script>

	<form name="frm" method="post" action="">	
		<table width="100%">
			<tr>
				<td align="left">
					<!--#INCLUDE virtual="/Compartido/Encabezado.asp" -->
				</td>
			</tr>
			<tr>
				<td align="center">
					<table id="tabConsulta" class="borde" BORDER="0" cellpadding="4" cellspacing="0" style="HEIGHT: 195px" width="300px">
						<tr>
							<td class="Titulo" colspan="2" style="FONT-SIZE: 10pt; HEIGHT: 5px">Datos de la consulta</td>
						</tr>
						<tr>
							<td align="center">
								<table id="tabPeriodo" class="bordeinactivo" cellspacing="0" cellpadding="3">									
									<tr>
										<td colspan="2" class="tituloinactivo">Periodo&nbsp;&nbsp;<a style="font-size: 8pt">(ddmmyy)</a></td>
									</tr>
									<tr>
										<td>Fecha</td> 
										<td><input SIZE="8" VALUE="<%=sFecha%>" name="txtFecha"></td>
									</tr>								
								</table>
							</td>
						</tr>
						<tr align="middle">
							<td colspan="2">
								<img id="imgAceptar" src="../images/BotonAceptar.jpg" style="CURSOR: hand" WIDTH="70" HEIGHT="20">
							</td>
						</tr>						
					</table>
				</td>
			</tr>
			<tr>
				<td>	
					<table STYLE="COLOR: #505050; FONT-FAMILY: Verdana; FONT-SIZE: 10px; POSITION: relative; TOP: 0px" align="center" width="500px">				
						<tr class="Encabezado">			
							<td><b>Glosa</b></td>
							<td style="Display: none;"><b>Cuenta</b></td>							
							<td><b>Debe</b></td>
							<td><b>Haber</b></td>							
						</tr>						
						
						<%
						dim nTotalDebePago, nTotalHaberPago, nTotalDebeEnvio, nTotalHaberEnvio
						nTotalDebePago = 0
						nTotalHaberPago = 0
						nTotalDebeEnvio = 0
						nTotalHaberEnvio = 0
						do while not rs.eof%>							
							<%if rs("tipogiro") = 1 then	' giro pagado por la sucursal%>
								<tr><td>&nbsp;</td></tr>
								<!-- línea 1-->							
								<tr style="HEIGHT: 25px; display: none" onmouseover="javascript:this.bgColor='#a4dded';" bgColor="#dbf7ff" onmouseout="javascript:this.bgColor='#DAF6FF';" style="cursor: hand">
									<td><b><%="PAGO Giro : " & rs("codigo_giro")%></b></td>
									<td><b>Cta. Cte. Giros</b></td>
									<td align="right"><b><%=formatnumber(ccur(rs("monto_giro")) + ccur(rs("comisionpagador")),0)%></b></td>
									<td align="right"><b>0</b></td>
								</tr>
							
								<!-- línea 2-->
								<tr style="HEIGHT: 25px; display: none" onmouseover="javascript:this.bgColor='#a4dded';" bgColor="#dbf7ff" onmouseout="javascript:this.bgColor='#DAF6FF';" style="cursor: hand">
									<td><b></b></td>
									<td><b>Caja Pesos</b></td>
									<td align="right"><b>0</b></td>
									<td align="right"><b><%=formatnumber(rs("monto_giro"),0)%></b></td>
								</tr>
								
								<!-- línea 3-->
								<tr style="HEIGHT: 25px; display: none" onmouseover="javascript:this.bgColor='#a4dded';" bgColor="#dbf7ff" onmouseout="javascript:this.bgColor='#DAF6FF';" style="cursor: hand">
									<td><b></b></td>
									<td><b>Otros Ingresos Operacionales</b></td>
									<td align="right"><b>0</b></td>
									<td align="right"><b><%=formatnumber(rs("comisionpagador"),0)%></b></td>
								</tr>
								<%
									nTotalDebePago = ccur(nTotalDebePago) + (ccur(rs("monto_giro")) + ccur(rs("comisionpagador")))
									nTotalHaberPago = ccur(nTotalHaberPago) + (ccur(rs("monto_giro")) + ccur(rs("comisionpagador")))
								%>
								
							
							<%elseif rs("tipogiro") = 2 then	' giro enviado por la sucursal%>
								<tr><td>&nbsp;</td></tr>
								<!-- línea 1-->
								<tr style="HEIGHT: 25px; display: none"  onmouseover="javascript:this.bgColor='#a4dded';" bgColor="#dbf7ff" onmouseout="javascript:this.bgColor='#DAF6FF';" style="cursor: hand">
									<td><b><%="ENVIO Giro : " & rs("codigo_giro")%></b></td>
									<td><b>Caja Pesos</b></td>
									<td align="right"><b><%=formatnumber(ccur(rs("monto_giro")) + ccur(rs("tarifacobrada")),0)%></b></td>
									<td align="right"><b>0</b></td>
								</tr>
							
								<!-- línea 2-->
								<tr style="HEIGHT: 25px; display: none"  onmouseover="javascript:this.bgColor='#a4dded';" bgColor="#dbf7ff" onmouseout="javascript:this.bgColor='#DAF6FF';" style="cursor: hand">
									<td><b></b></td>
									<td><b>Cta. Cte. Giros</b></td>
									<td align="right"><b>0</b></td>
									<td align="right"><b><%=formatnumber((ccur(rs("monto_giro")) + ccur(rs("tarifacobrada"))) - ccur(rs("comisioncaptador")),0)%></b></td>
								</tr>
							
								<!-- línea 3-->
								<tr style="HEIGHT: 25px; display: none"  onmouseover="javascript:this.bgColor='#a4dded';" bgColor="#dbf7ff" onmouseout="javascript:this.bgColor='#DAF6FF';" style="cursor: hand">
									<td><b></b></td>
									<td><b>Comisiones Ganadas</b></td>
									<td align="right"><b>0</b></td>
									<td align="right"><b><%=formatnumber(rs("comisioncaptador"),0)%></b></td>
								</tr>
									<%
										nTotalDebeEnvio = ccur(nTotalDebeEnvio) + (ccur(rs("monto_giro")) + ccur(rs("tarifacobrada")))
										nTotalHaberEnvio = ccur(nTotalHaberEnvio) + (ccur(rs("monto_giro")) + ccur(rs("tarifacobrada")))
									%>						
								
							<%end if %>
							<%rs.MoveNext%>
						<%loop%>
						<%if (nTotalHaberPago <> 0 and nTotalHaberPago <> 0) or (nTotalHaberEnvio <> 0 and nTotalHaberEnvio <> 0) then%>							
							<tr style="HEIGHT: 25px;"  onmouseover="javascript:this.bgColor='#a4dded';" bgColor="#dbf7ff" onmouseout="javascript:this.bgColor='#DAF6FF';" style="cursor: hand">
								<td><b>Giros Pagados</b></td>								
								<td align="right"><b><%=formatnumber(nTotalDebePago,0)%></b></td>
								<td align="right"><b><%=formatnumber(nTotalHaberPago,0)%></b></td>
							</tr>							
							<tr style="HEIGHT: 25px;"  onmouseover="javascript:this.bgColor='#a4dded';" bgColor="#dbf7ff" onmouseout="javascript:this.bgColor='#DAF6FF';" style="cursor: hand">
								<td><b>Giros Enviados</b></td>								
								<td align="right"><b><%=formatnumber(nTotalDebeEnvio,0)%></b></td>
								<td align="right"><b><%=formatnumber(nTotalHaberEnvio,0)%></b></td>
							</tr>
							<tr><td>&nbsp;</td></tr>
							<tr style="HEIGHT: 25px;"  onmouseover="javascript:this.bgColor='#a4dded';" bgColor="#dbf7ff" onmouseout="javascript:this.bgColor='#DAF6FF';" style="cursor: hand">
								<td><b>TOTAL</b></td>								
								<td align="right"><b><%=formatnumber(nTotalDebePago + nTotalDebeEnvio,0)%></b></td>
								<td align="right"><b><%=formatnumber(nTotalHaberPago + nTotalHaberEnvio,0)%></b></td>
							</tr>
							<tr><td>&nbsp;</td></tr>
							<tr>
								<td colspan="4" align="right"><input type="button" name="cmdGenerar" value="Generar"></td>
							</tr>
							
						<%end if%>
					</table>		
				</td>
			</tr>
		</div>	
	</form>
</BODY>
</HTML>
<%set rs = nothing%>
