<%@ Language=VBScript %>
<%
	'Se asegura que la página no se almacene en la memoria cache
	Response.Expires = 0
	If Session("CodigoCliente") = "" Then
		response.Redirect "../Compartido/TimeOut.htm"
		response.end
	End If	
%>
<%
	dim sMoneda
	dim cParidad
	dim cRecargo
	dim sFechaTermino
	dim sHoraTermino	
	dim sUsuario
	dim cParidadOriginal
	
	sub cargarultimaparidad(Moneda)
		dim rs
		dim sSQL
		
		sSQL = " select * " & _
				 " from paridadesmoneda " & _
				 " where moneda = " & EvaluarStr(request.Form("cbxMoneda")) & _
				   " and estado = 1 "	
		
		set rs = EjecutarSQLCliente(Session("afxCnxAfexchange"), sSQL)
		if err.number <> 0 then
			set rs = nothing
			mostrarerrorms "Cargar Monedas Mantenedor"
		end if
		
		if not rs.eof then
		
			sMoneda = rs("moneda")
			cParidad = formatnumber(rs("paridad"), 8)
			cRecargo = formatnumber(rs("recargo"), 2)
			cParidadOriginal = formatnumber(rs("paridadoriginal"), 8)
			sFechaTermino = rs("fechatermino")
			sHoraTermino = rs("horatermino")
			sUsuario = rs("usuario")			
		
			rs.close
		end if
		
		set rs = nothing
	end sub
	if sHoraTermino = empty then sHoraTermino = "16:30:00"
%>
<!--#INCLUDE virtual="/compartido/Rutinas.asp" -->
<!--#INCLUDE virtual="/compartido/Rutinas.htm" -->
<!--#INCLUDE virtual="/compartido/Errores.asp" -->
<!--#INCLUDE virtual="/compartido/Transferencias.asp" -->
<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link rel="stylesheet" type="text/css" href="../Estilos/Principal.css">
<link rel="stylesheet" type="text/css" href="../Estilos/Cliente.css">
<link rel="stylesheet" type="text/css" href="../Estilos/Reportes.css">
</head>

<script language="vbscript">
<!--
	
	sub window_onLoad()
		<%	
			if request("Accion") = 1 then
				cargarultimaparidad(request.form("cbxMoneda"))
		%>		
				frmParidad.cbxMoneda.value = "<%=request.form("cbxMoneda")%>"
		<%		
			end if
		%>
		
	end sub
	
	sub cmdAceptar_onClick()
		' verifica los campos vacios
		if frmParidad.cbxMoneda.value = empty then
			msgbox "Debe seleccionar una moneda.",,"Paridades"
			frmParidad.cbxMoneda.focus
			exit sub
		end if
		if frmParidad.txtParidadFinal.value = empty then
			msgbox "Debe ingresar una paridad final.",,"Paridades"
			frmParidad.txtParidad.focus
			exit sub
		elseif frmParidad.txtParidadFinal.value <= 0 then
			msgbox "Debe ingresar una paridad final valida.",,"Paridades"
			frmParidad.txtParidad.focus
			exit sub
		end if
		if frmParidad.txtRecargo.value = empty then
			frmParidad.txtRecargo.value = "0"
		end if
	
		frmParidad.txtParidadFinal.disabled = false
		'frmParidad.txtFecha.disabled = false
	
		frmParidad.action = "GrabarParidadMoneda.asp"
		frmParidad.submit()
		frmParidad.action = ""
	end sub
	
	sub cbxMoneda_onChange()
		frmParidad.action = "MantenedorParidades.asp?accion=1"
		frmParidad.submit()
		frmParidad.action = ""
	end sub
	
	sub txtParidad_onBlur()
		ValidarParidad()
	end sub
	
	sub txtRecargo_onBlur()
		ValidarParidad()
	end sub
	
	sub ValidarParidad()
		if frmParidad.txtParidad.value = empty then frmParidad.txtParidad.value = "0"
		frmParidad.txtParidad.value = formatnumber(frmParidad.txtParidad.value, 8)
		if frmParidad.txtRecargo.value = empty then frmParidad.txtRecargo.value = "0"
		frmParidad.txtRecargo.value = formatnumber(frmParidad.txtRecargo.value, 2)
		
		
		if ccur(frmParidad.txtRecargo.value) > 0 then			
			cValor = formatnumber((frmParidad.txtParidad.value) * (1 + frmParidad.txtRecargo.value / 100), 8)
			'cValor = ccur(cValor) + ccur(frmParidad.txtParidad.value)			
		else
			cValor = frmParidad.txtParidad.value
		end if
		frmParidad.txtParidadFinal.value = cValor
	end sub
-->
</script>	

<body>
<form name="frmParidad" method="post" action="">
	<table>
		<tr>
			<td><b>Moneda</b><br>
				<select name="cbxMoneda">
					<% cargarmonedasmantenedor %>
				</select>
			</td>
			<td><b>Paridad</b><br>
				<input type="text" name="txtParidad" onKeyPress="IngresarTexto(1)" size="10" value="<%=cParidadOriginal%>">
			</td>
			<td><b>Recargo</b><br>
				<input type="text" name="txtRecargo" onKeyPress="IngresarTexto(1)" size="10" value="<%=cRecargo%>">
			</td>
			<td><b>Paridad Final</b><br>
				<input type="text" disabled name="txtParidadFinal" size="10" value="<%=cParidad%>">
			</td>
			<td><b>Fecha Término</b><br>
				<input type="text" name="txtFecha" size="10" value="<%=Date%>">
			</td>
			<td><b>Hora Término</b><br>
				<input type="text" name="txtHora" size="10" value="<%=sHoraTermino%>">
			</td>
			<%If sUsuario <> empty then%>
				<td><b>Usuario</b><br>
					<b><%=sUsuario%></b>
				</td>
			<%End if%>
		</tr>
		<tr>
			<td>
				<input type="button" name="cmdAceptar" value="Aceptar">
			</td>
		</tr>
	</table>
</form>	
</body>
</html>