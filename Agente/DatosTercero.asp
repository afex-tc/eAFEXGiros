<%@ Language=VBScript %>
<%
	'Se asegura que la página no se almacene en la memoria cache
	Response.Expires = 0
	If Session("CodigoCliente") = "" Then
		response.Redirect "../Compartido/TimeOut.htm"
		response.end
	End If
%>
<!-- #INCLUDE virtual="/Compartido/Errores.asp" -->
<!-- #INCLUDE virtual="/Compartido/Rutinas.asp" -->
<!-- #INCLUDE virtual="/agente/Constantes.asp" -->
<%
	Dim sNombres, sApellidos, sRut, sPasaporte, sPaisPasaporte
	
	sNombres = Request("Nombres")
	sApellidos = Request("Apellidos")
	sRut = Request("Rut")
	sPasaporte = Request("Pasaporte")
	sPaisPasaporte = Request("PaisPasaporte")
	If sPaisPasaporte = "" Then
		sPaisPasaporte = Session("PaisCliente")
	End If
	
%>
<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<title>Datos del Tercero</title>
<link rel="stylesheet" type="text/css" href="../Estilos/Cliente.css">
</head>

<script LANGUAGE="VBScript">
<!--
	window.dialogWidth = 26
	window.dialogHeight = 16
	window.dialogLeft = 240
	window.dialogTop = 220
	window.defaultstatus = ""
	
	Sub window_onLoad()
		If txtPasaporte.value <> "" Then
			optPasaporte_onClick 
		End If
		<% If Session("Categoria") = 4 Then %>
				txtPasaporte.Select
				txtPasaporte.focus
		<% Else %>
				txtRut.Select
				txtRut.focus
		<% End If %>
				
	End Sub
	
	Sub optRut_onClick()
		window.optpasaporte.checked=False
		window.optRut.checked = True
		window.txtpasaporte.style.display="none"
		window.cbxPaisPasaporte.style.display="none"
		window.txtpasaporte.value = ""
		window.cbxPaisPasaporte.value = ""
		lblPaisPasaporte.style.display="none"
		window.txtRut.style.display = ""		
	End Sub
	
	Sub optPasaporte_onClick()
		window.optRut.checked=False
		window.optpasaporte.checked=True
		window.txtRut.style.display = "none"
		window.txtRut.value = ""
		window.txtpasaporte.style.display=""
		window.cbxPaisPasaporte.style.display=""
		lblPaisPasaporte.style.display=""
	End Sub

	Sub txtRut_onBlur()
		Dim sRut
		
		If txtRut.value = "" Then Exit Sub
		sRut = ValidarRut(txtRut.value)
		If sRut = Empty Then
			msgbox "El número de Rut no es válido"
			txtRut.focus()
		Else
			txtRut.value = sRut
		End If
		
	End Sub

	
	Sub imgAceptar_onClick()
		
		window.returnvalue = window.txtRut.value & ";" & window.txtPasaporte.value & ";" & _
							 window.cbxPaisPasaporte.value & ";" & _	
							 window.txtNombres.value & ";" & _
							 window.txtApellidos.value
			
		window.close
		
	End Sub		

//-->
</script>
<body>
<center>
<input type="hidden" id="txtCodigoCliente">
<table id="tabConsulta" class="border" BORDER="0" cellpadding="4" cellspacing="0" style="HEIGHT: 195px" width="350px">	
<!--<tr><td class="Titulo" colspan="3" style="FONT-SIZE: 10pt; HEIGHT: 5px">Datos del Tercero</td></tr>-->
<tr HEIGHT="40">
	<td VALIGN="center" colspan="3"><br>
		<table border="0" cellpadding="1">
		<tr>
			<td>
				<% If Session("Categoria") = 4 Then %>
						<input TYPE="radio" name="optRut" style="display: none">
						<input TYPE="radio" name="optPasaporte" CHECKED>Id
				<% Else %>
						<input TYPE="radio" name="optRut" CHECKED>Rut
						<input TYPE="radio" name="optPasaporte">Pasaporte
				<% End If %>
			</td>
			<td id="lblPaisPasaporte" style="display: none">Pais</td>
		</tr>
		<tr>
			<td>
				<% If Session("Categoria") = 4 Then %>
						<input id="txtRut" style="width: 150px; text-align:right; display: none" value="<%=sRut%>">
						<input id="txtPasaporte" style="width: 150px" value="<%=sPasaporte%>">
				<% Else %>
						<input id="txtRut" style="width: 150px; text-align:right" value="<%=sRut%>">
						<input id="txtPasaporte" style="width: 150px; display: none" value="<%=sPasaporte%>">
				<% End If %>
			</td>
			<td>
				<% If Session("Categoria") = 4 Then %>
						<select name="cbxPaisPasaporte" style="width: 150px">									
				<% Else %>
						<select name="cbxPaisPasaporte" style="width: 150px; display: none">									
				<% End If %>
					<%
						CargarPaisPasaporte sPaisPasaporte ' INTERNO-1831 - JFMG 28-07-2014
					%>
				</select>
			</td>
		</tr>
		</table>
			</td>
</tr>
<tr align="center">
	<td title="Giro" style align="left">
	<table>
	<tr><td id="tdNombres">Nombres<br><input SIZE="40" id="txtNombres" style="width: 300px" onKeyPress="IngresarTexto(2)" onBlur="txtNombres.value=MayMin(txtNombres.value)" value="<%=sNombres%>"></td></tr>
	<tr><td id="tdApellidos">Apellidos<br><input SIZE="40" id="txtApellidos" style="width: 300px" onKeyPress="IngresarTexto(2)" onBlur="txtApellidos.value=MayMin(txtApellidos.value)" value="<%=sApellidos%>"></td></tr>
	</table>
	</td>
<tr align="middle">
	<td colspan="2"><img id="imgAceptar" src="../images/BotonAceptar.jpg" style="CURSOR: hand" WIDTH="70" HEIGHT="20"></td>
</tr>
</table>
</center>
<!--#INCLUDE virtual="/Compartido/Rutinas.htm" -->
</body>
</html>