<%@ Language=VBScript %>
<%
	Response.Expires = 0
	If Session("CodigoCliente") = "" Then
		response.Redirect "../Compartido/TimeOut.htm"
		response.end
	End If
%>
<% 
	response.Expires = 0
	response.Buffer = True
	response.Clear 
%>
<!--#INCLUDE virtual="/Compartido/TimeOut.asp" -->
<!--#INCLUDE virtual="/Compartido/Rutinas.asp" -->
<!--#INCLUDE virtual="/Compartido/Errores.asp" -->
<!--#INCLUDE virtual="/sucursal/Rutinas.asp" -->
<%

	Dim sPais, sCiudad, sSQL, sMoneda
	Dim rs
	
	sPais = Request.Form("cbxPais")
	sCiudad = Request.Form("cbxCiudad")
	sMoneda = Request.Form("cbxMoneda")
	sMonto = Request.Form("txtMonto")


	if Request("Grabar") = 1 then
		sSQL = " InsertarMatrizGiro " & evaluarstr(sPais) & ", " & evaluarstr(sCiudad) & ", " & _
										evaluarstr(sMoneda) & ", " & formatonumerosql(ccur(Request.Form("txtMonto")))
		set rs = EjecutarSQLCliente(Session("afxCnxcorporativa"), sSQL)
		if err.number <> 0 then
			set rs = nothing
			mostrarerrorms "Error al grabar la Matriz. "
		end if
	else

		if sPais = empty then sPais = "**"
		if sCiudad = empty then sCiudad = "***"
		sMonto = empty
		
		sSQL = "select monto from matrizgiro where pais = " & evaluarstr(sPais) & " and ciudad = " & evaluarstr(sCiudad)
		set rs = EjecutarSQLCliente(Session("afxCnxcorporativa"), sSQL)
		if err.number <> 0 then
			set rs = nothing
			mostrarerrorms "Error al cargar la Matriz."
		end if
		if not rs.eof then
			sMonto = rs("monto")
		end if		
	end if	
	
%>
<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<title>Configuración Ficha Cliente Giro</title>
<link rel="stylesheet" type="text/css" href="../Estilos/Cliente.css">
<link rel="stylesheet" type="text/css" href="../Estilos/AFX2004.css">
</head>

<script LANGUAGE="VBScript">
<!--
	On Error Resume Next

	Sub window_onload()
	End Sub

	Sub GuardarCambios_onclick()
		frmFicha.action = "http:GuardarFichaProducto.asp"
		frmFicha.submit 
		frmFicha.action = ""
	End Sub
	
	sub cbxPais_onChange()
		frmFicha.action = "configuracionfichagiro.asp"
		frmFicha.submit 
		frmFicha.action = ""
	end sub
	
	sub cbxCiudad_onChange()
		frmFicha.action = "configuracionfichagiro.asp"
		frmFicha.submit 
		frmFicha.action = ""
	end sub
	
	sub txtMonto_onBlur()
		frmFicha.txtMonto.value = formatnumber(frmFicha.txtMonto.value, 2)
	end sub	
	
	Sub GuardarCambios_onclick()
		if frmFicha.txtMonto.value = empty then exit sub
		
		frmFicha.action = "http:ConfiguracionFichaGiro.asp?Grabar=1"
		frmFicha.submit 
		frmFicha.action = ""
	End Sub	
	
//-->
</script>
<!--#INCLUDE virtual="/Compartido/Rutinas.htm" -->
<body id="bb" border="0" style="margin: 2 2 2 2" >
<form id="frmFicha" method="post">
	<table class="Borde" id="" BORDER="0" cellpadding="0" cellspacing="0" style="HEIGHT: 150px; width:100%; background-color: #f4f4f4">	
	<tr height="40" style="background-color: #ffeeaa; #ffdd77; #e1e1e1"><td colspan="3" style="font-size: 16pt">&nbsp;&nbsp;Configuración Ficha Cliente</td></tr>
	<tr height="1" style="background-color: silver "><td colspan="3" ></td></tr>
	<tr height="4"><td colspan="3" ></td></tr>
	<tr>
		<td colspan="3">
		<table width="100%">
		<tr>
			<td>
				<table style="font-size: 8pt">
					<tr>
					</tr>
				</table>
			</td>
			<td align="right">
				<table>
				<tr>
					<td colspan="2">
						<!--#INCLUDE virtual="/sucursal/MenuConfiguracionFichaGiro.htm" -->
					</td>
					</div>
				</tr>
				</table>
			</td>
		</tr>
		</table>
		</td>
	</tr>

	<tr height="10"><td></td></tr>

	<tr>
		<td colspan="3" align="center">
			<table cellpadding="6" cellspacing="1" style="background-color: silver;" >			
				<tr>
					<td>Pa&iacute;s<br>
						<select name="cbxPais">
							<%CargarUbicacion 1, "", sPais%>
						</select>
					</td>
					<td>Ciudad<br>
						<select name="cbxCiudad">
							<%CargarCiudadesPais sPais, sCiudad%> 
						</select>
					</td>
					<td>Moneda<br>
						<select name="cbxMoneda">
							<option value="CLP">Pesos Chilenos</option>
							<option value="USD">Dolares</option>
						</select>
					</td>
					<td>Monto US$<br>
						<input type="text" size="10" name="txtMonto" onkeypress="IngresarTexto(1)" maxlength="7" value="<%=formatnumber(sMonto, 2)%>" style="text-align: right">
					</td>
				</tr>
			</table>
		</td>
	</tr>


	<tr height="10"><td></td></tr>
	</table>
</form>
</body>
</html>