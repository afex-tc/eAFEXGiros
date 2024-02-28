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
%>
<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<title>Datos del Reclamo</title>
<link rel="stylesheet" type="text/css" href="../Estilos/Cliente.css">
</head>

<script LANGUAGE="VBScript">
<!--
	window.dialogWidth = 24
	window.dialogHeight = 14
	window.dialogLeft = 240
	window.dialogTop = 220
	window.defaultstatus = ""
	
	Sub window_onLoad()
	End Sub
	
	Sub imgAceptar_onClick()
		
		window.returnvalue = window.cbxTipo.value & ";" & window.cbxTipo.item(window.cbxTipo.selectedIndex).text & ":" & window.txtMotivo.value
			
		window.close
		
	End Sub		

//-->
</script>
<body>
<center>
<input type="hidden" id="txtCodigoCliente">
<table id="tabConsulta" class="border" BORDER="0" cellpadding="4" cellspacing="0" style="height: 120px">	
	<tr height="10"><td></td></tr>
	<tr><td>Tipo</td>
		<td><select name="cbxTipo" style="width: 300px">
				<%
					CargarReclamo
				%>
			</select>
		</td>
	</tr>
	<tr><td>Motivo</td>
		<td><input SIZE="40" id="txtMotivo" style="width: 300px"></td>
	</tr>
	<tr align="middle">
		<td colspan="2"><img id="imgAceptar" src="../images/BotonAceptar.jpg" style="CURSOR: hand" WIDTH="70" HEIGHT="20"></td>
	</tr>
</table>
</center>
<!--#INCLUDE virtual="/Compartido/Rutinas.htm" -->
</body>
</html>