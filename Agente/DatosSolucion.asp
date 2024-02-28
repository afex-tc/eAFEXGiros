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
	Dim sNombres, sApellidos, sDireccion
	Dim sAreaFono, sPaisFono, sFono, sDescripcion
	
	sNombres = Request("Nombres")
	sApellidos = Request("Apellidos")
	sDireccion = Request("Direccion")
	sAreaFono = Request("AreaFono")
	sPaisFono = Request("PaisFono")
	sFono = Request("Fono")
	sDescripcion = Request("Descripcion")
	
	
%>
<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<title>Datos del Tercero</title>
<link rel="stylesheet" type="text/css" href="../Estilos/Cliente.css">
</head>

<script LANGUAGE="VBScript">
<!--
	window.dialogWidth = 25
	window.dialogHeight = 19
	window.dialogLeft = 160
	window.dialogTop = 220
	window.defaultstatus = ""
	
	Sub window_onLoad()
	End Sub
	
	
	Sub imgAceptar_onClick()
		
		window.returnvalue = window.txtNombres.value & ";" & _
							 window.txtApellidos.value & ";" & _
							 window.txtDireccion.value  & ";" & _	
							 window.txtFono.value & ";" & _
							 window.txtdescripcion.value
		window.close
		
	End Sub		

//-->
</script>
<body>
<center>
<input type="hidden" id="txtCodigoCliente">
<table id="tabConsulta" class="border" BORDER="0" cellpadding="4" cellspacing="0" style="HEIGHT: 195px" width="350px">	
<!--<tr><td class="Titulo" colspan="3" style="FONT-SIZE: 10pt; HEIGHT: 5px">Datos del Tercero</td></tr>-->
<tr align="center">
	<td title="Giro" style align="left">
	<table>
	<tr><td>Nombres<br><input SIZE="40" id="txtNombres" style="width: 300px" onKeyPress="IngresarTexto(2)" onBlur="txtNombres.value=MayMin(txtNombres.value)" value="<%=sNombres%>"></td></tr>
	<tr><td>Apellidos<br><input SIZE="40" id="txtApellidos" style="width: 300px" onKeyPress="IngresarTexto(2)" onBlur="txtApellidos.value=MayMin(txtApellidos.value)" value="<%=sApellidos%>"></td></tr>
	<tr><td>Direccion<br><input SIZE="40" id="txtDireccion" style="width: 300px" onBlur="txtDireccion.value=MayMin(txtDireccion.value)" value="<%=sDireccion%>"></td></tr>
	<tr>
		<td colspan="1">Teléfono<br>
		<input disabled id="txtPaisFono" style="width: 40px" value="<%=sPaisFono%>">
		<input disabled id="txtAreaFono" style="width: 40px" value="<%=sAreaFono%>">
		<input name="txtFono" style="width: 90px" onkeypress="IngresarTexto(1)" onblur="txtFono.value=MayMin(txtFono.value)" value="<%=sFono%>">
		</td>
	</tr>
	<tr><td>Descripcion de la solución<br><input SIZE="40" id="txtDescripcion" style="width: 300px" onBlur="txtDescripcion.value=MayMin(txtDescripcion.value)" value="<%=sDescripcion%>"></td></tr>
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