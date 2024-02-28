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
	sFono = "(" & Request("PaisFono") & Request("AreaFono") & ") " & Request("Fono")
	sDescripcion = Request("Descripcion")
	
	
%>
<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<title>Comprobante de Pago</title>
<link rel="stylesheet" type="text/css" href="../Estilos/Cliente.css">
</head>

<script LANGUAGE="VBScript">
<!--
	window.dialogWidth = 40
	window.dialogHeight = 24
	window.dialogLeft = 160
	window.dialogTop = 220
	window.defaultstatus = ""
	
	Sub window_onLoad()
	End Sub
	
	
	Sub imgAceptar_onClick()
		
		window.close
		
	End Sub		

//-->
</script>
<body>
<center>
<input type="hidden" id="txtCodigoCliente">
<table id="tabConsulta" class="border" BORDER="0" cellpadding="4" cellspacing="0" style="HEIGHT: 90%px; font-size: 14pt" width="90%">	
<!--<tr><td class="Titulo" colspan="3" style="FONT-SIZE: 10pt; HEIGHT: 5px">Datos del Tercero</td></tr>-->
<tr align="center">
	<td title="Giro" style align="left">
	<table>
	<tr><td colspan="3"><b>Beneficiario</b><br>
		&nbsp;&nbsp;&nbsp;Nombre:&nbsp;<%=Request("NombresB")%>&nbsp;<%=Request("ApellidosB")%><br>
		&nbsp;&nbsp;&nbsp;Direccion:&nbsp;<%=Request("DireccionB")%><br>
		&nbsp;&nbsp;&nbsp;Ciudad:&nbsp;<%=Request("CiudadB")%><br>
		</td>
		<td colspan="2"><br>
		&nbsp;&nbsp;&nbsp;Rut:&nbsp;<%=Request("RutB")%><br>
		&nbsp;&nbsp;&nbsp;Telefono:&nbsp;<%=Request("FonoB")%>
		</td>
	</tr>
	<tr>
		<td colspan="3"><b>Giro</b><br>
		&nbsp;&nbsp;&nbsp;Codigo:&nbsp;<%=Request("Codigo")%><br>
		&nbsp;&nbsp;&nbsp;Agente Captador:&nbsp;<%=Request("Captador")%><br><br>
		</td>
		<td colspan="2"><br>
		&nbsp;&nbsp;&nbsp;Referencia:&nbsp;<%=Request("Invoice")%><br>
		&nbsp;&nbsp;&nbsp;Agente Pagador:&nbsp;<%=Request("Pagador")%><br><br>
		</td>
	</tr>
	<tr>
		<td colspan="3"><b>Remitente</b><br>
		&nbsp;&nbsp;&nbsp;Nombre:&nbsp;<%=sNombres%>&nbsp;<%=sApellidos%><br>
		&nbsp;&nbsp;&nbsp;Direccion:&nbsp;<%=sDireccion%><br>
		&nbsp;&nbsp;&nbsp;Ciudad:&nbsp;<%=Request("Ciudad")%><br>
		&nbsp;&nbsp;&nbsp;Mensaje:&nbsp;<%=Request("Mensaje")%><br><br>		
		</td>
		<td colspan="2"><br><br><br>
		&nbsp;&nbsp;&nbsp;Pais:&nbsp;<%=Request("Pais")%>
		</td>
	</tr>
	<tr>
		<td><b>Receptor</b><br>
			&nbsp;&nbsp;&nbsp;Recibi Conforme:&nbsp;<%=Request("NombreRetira")%><br>
			&nbsp;&nbsp;&nbsp;Rut:&nbsp;<%=Request("RutRetira")%><br><br>
		</td>
		<td></td>
	</tr>
	<tr>
		<td colspan="5" align="right"><br>
			<%=Request("Prefijo")%>&nbsp;<%=Request("Monto")%><br>
		</td>
	</tr>
	</table>
	</td>
<tr align="middle">
	<td colspan="2"><img id="imgAceptar" src="../images/BotonCerrar.jpg" style="CURSOR: hand" WIDTH="70" HEIGHT="20"></td>
</tr>
</table>
</center>
<!--#INCLUDE virtual="/Compartido/Rutinas.htm" -->
</body>
</html>