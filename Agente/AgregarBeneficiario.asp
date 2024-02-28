<%@ Language=VBScript %>
<!-- #INCLUDE virtual="/Compartido/TimeOut.asp" -->
<!-- #INCLUDE virtual="/Compartido/Errores.asp" -->
<!-- #INCLUDE virtual="/Compartido/Rutinas.asp" -->
<%
	dim sCliente, rs, sSQL
	

	sCliente = request("Cliente")
	if trim(sCliente) = "" then sCliente = Request.Form("txtCodigoCliente")
	if Trim(sCliente) = "" then 
		Response.Write "A este cliente no se le pueden asociar Beneficiarios, ya que no es Cliente de Giros."
		Response.End 
	end if
	
	' saca la lista de beneficiarios asociados al cliente
	sSQL = "select * from beneficiarios where codigocliente = " & evaluarstr(sCliente)
	'Response.Write ssql
	'Response.End 	
	set rs = ejecutarsqlcliente(session("afxCnxafexpress"), sSQL)
	if err.number <> 0 then
		Response.Write "Error al consultar los Beneficiarios. " & err.Description 
		Response.End
	end if
%>
<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link rel="stylesheet" type="text/css" href="../Estilos/Cliente.css">
</head>
<script LANGUAGE="VBScript">
<!--
	'Variables de módulo
	'Variables para encabezado	
	sub window_onLoad()		
		CargarCliente
		frmBeneficiario.txtNombres.focus()
	End Sub	

	Sub CargarCliente()
		frmBeneficiario.txtApellidos.value = "<%=request.form("txtApellidos")%>"
		frmBeneficiario.txtNombres.value = "<%=request.form("txtNombres")%>"
		frmBeneficiario.cbxCiudad.value = "<%=request.form("cbxCiudad")%>"
		frmBeneficiario.cbxPais.value = "<%=request.form("cbxPais")%>"
	End Sub	
		
	Sub cbxPais_onblur()
		Dim sCiudad
		
		If frmBeneficiario.cbxPais.value = "" Then Exit Sub
		If frmBeneficiario.cbxPais.value = "<%=request.form("cbxPais")%>" Then Exit Sub
		
		frmBeneficiario.action = "AgregarBeneficiario.asp"		
		frmBeneficiario.submit 
		frmBeneficiario.action = ""
	End Sub
	
	Sub LimpiarControles
		frmBeneficiario.txtApellidos.value = ""		
		frmBeneficiario.txtNombres.value = ""
		frmBeneficiario.cbxPais.value = ""
		frmBeneficiario.cbxCiudad.value = ""
	End Sub		
	
-->
</script>

<body >
<form id="frmBeneficiario" method="post">
<br>
<table CELLSPACING="0" border="0" style="position: relative; top: 0px; left: 2px;">
	<tr HEIGHT="15">
		<td colspan="3">&nbsp;</td>
	</tr>
</table>
<table class="Borde" align="center" ID="tabPaso1" CELLSPACING="0" border="0" style="position: relative; top: 0px; left: 2px;">	
	<tr HEIGHT="15">
		<td colspan="3" class="titulo">Nuevo Beneficiario</td>
	</tr>
	<tr HEIGHT="2">
		<td></td>
		<td COLSPAN="2"></td>
	</tr>
	<tr>
		<td width="1px"></td>
		<td>
			<table width="100%" border="0">		
				<tr>
					<td></td>
					<td colspan="2">Nombres<br>
						<input NAME="txtNombres" style="width: 350px" onkeypress="IngresarTexto(2)" onblur="frmBeneficiario.txtNombres.value=MayMin(frmBeneficiario.txtNombres.value)">*
					</td>
				</tr>
				<tr>
					<td></td>
					<td colspan="2">Apellidos<br>
						<input name="txtApellidos" style="width: 350px" onkeypress="IngresarTexto(2)" onblur="frmBeneficiario.txtApellidos.value=MayMin(frmBeneficiario.txtApellidos.value)">*
					</td>
				</tr>
				<tr>
					<td></td>
					<td>País<br>
						<select name="cbxPais" style="width: 170px">
							<%CargarUbicacion 1, "", sPais%>
						</select>
					</td>
					<td colspan="1">Ciudad<br>
						<select  name="cbxCiudad" style="width: 170px">
							<%CargarCiudadesPais request.form("cbxPais"), request.form("cbxCiudad")%>
						</select>
					</td>
				</tr>
				<tr>
					<td></td>
					<td colspan="2">Teléfono<br>
						<input name="txtFono" style="width: 350px">					
					</td>
				</tr>	
			</table>
		</td>
	</tr>
	<tr>
		<td></td>
		<td align="right"><input type="button" name="cmdAceptar" value="Aceptar"></td>
	</tr>
	<tr>
		<td>&nbsp;</td>
	</tr>
	<tr>
		<td colspan="2" style="font-size: 16px"><b>Beneficiarios</b><br>
			<table>
				<tr>
					<td><b>Nombres</b></td>
					<td><b>Apellidos</b></td>
					<td><b>País</b></td>
					<td><b>Ciudad</b></td>
					<td><b>Teléfono</b></td>
				</tr>	
			<%do while not rs.eof%>
				<tr>
					<td><%=rs("nombres")%></td>
					<td><%=rs("apellidos")%></td>
					<td><%=rs("pais")%></td>
					<td><%=rs("ciudad")%></td>
					<td><%=rs("numerocontacto")%></td>
				</tr>
	
				<%rs.movenext%>
			<%loop%>
			</table>
		</td>
	</tr>
</table>
	<input type="hidden" name="txtcodigocliente" value="<%=sCliente%>">
</form>
</body>
<!--#INCLUDE virtual="/Compartido/Rutinas.htm" -->

<script>

	Sub cmdAceptar_onClick()
		
		if Not ValidarDatos Then
			Exit Sub
		end If					
		
		frmBeneficiario.action = "GrabarBeneficiario.asp"
		frmBeneficiario.submit 
		frmBeneficiario.action = ""		
	End Sub

	Function ValidarDatos()
		
		ValidarDatos = False
		If Trim(frmBeneficiario.txtNombres.value) = "" Then
			MsgBox "Debe ingresar los nombres del beneficiario.",,"AFEX"
			frmBeneficiario.txtNombres.focus 
			Exit Function
		End If
		If Trim(frmBeneficiario.txtApellidos.value) = "" Then
			MsgBox "Debe ingresar los apellidos beneficiario.",,"AFEX"
			frmBeneficiario.txtApellidos.focus
			Exit Function
		End If
		If Trim(frmBeneficiario.cbxPais.value) = "" Then
			MsgBox "Debe ingresar el pais del cliente",,"AFEX"
			frmBeneficiario.cbxPais.focus 
			Exit Function
		End If
		If Trim(frmBeneficiario.cbxCiudad.value) = "" Then
			MsgBox "Debe ingresar la ciudad del cliente",,"AFEX"
			frmBeneficiario.cbxCiudad.focus 
			Exit Function
		End If

		ValidarDatos = True
	End Function
	
</script>
</html>