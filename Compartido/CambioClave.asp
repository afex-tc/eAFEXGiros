<%@ Language=VBScript %>
<%
	'Se asegura que la página no se almacene en la memoria cache
	Response.Expires = 0
	If Session("CodigoCliente") = "" Then
		response.Redirect "TimeOut.htm"
		response.end
	End If
%>
<!--#INCLUDE virtual="/compartido/Errores.asp" -->
<!--#INCLUDE virtual="/compartido/Rutinas.asp" -->
<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link rel="stylesheet" type="text/css" href="../Estilos/Cliente.css">
</head>
<body>
<script LANGUAGE="VBScript">
<!--

	Const sEncabezadoFondo = "Servicios"
	Const sEncabezadoTitulo = "Cambio de Clave"
	Const sClass = "TituloPrincipal"

	Sub tdAceptar_onClick()
		If frmClave.txtClaveNueva.value <> frmClave.txtClaveNuevaR.value Then
			MsgBox "La clave reingresada es distinta a la nueva"
			Exit Sub
		End If
		If Not ValidarDatos() Then
			Exit Sub
		End If
		HabilitarControles
		frmClave.action = "GrabarCambioClave.asp"
		frmClave.submit()
		frmClave.action = ""		
	End Sub
	
	Function ValidarDatos()
		ValidarDatos = False
		If Trim(frmClave.txtClaveAnterior.value) = "" Then
			MsgBox "Debe ingresar la clave actual",,"AFEX"
			Exit Function
		End If	
		If Trim(frmClave.txtClaveNueva.value) = "" Then
			MsgBox "Debe ingresar la clave nueva",,"AFEX"
			Exit Function
		End If	
		If Trim(frmClave.txtClaveNuevaR.value) = "" Then
			MsgBox "Debe reingresar la clave nueva",,"AFEX"
			Exit Function
		End If	
		ValidarDatos = True
	End Function

-->
</script>
<!--#INCLUDE virtual="/compartido/Encabezado.htm" -->
<form id="frmClave" method="post">
<table class="Borde" HEIGHT="300" ID="tabPaso1" CELLSPACING="0" CELLPADDING="1" BORDER="0" WIDTH="560" STYLE="position: relative; LEFT: 4px;">
	<tr bgcolor="#dcf9ff">
		<td WIDTH="50"></td>
		<td HEIGHT="80" COLSPAN="2">
			 Aquí poner indicaciones acerca del cambio de clave
		</td>	
	</tr>
	<tr HEIGHT="15">
		<td colspan="3" height="15" bgcolor="steelblue"><font face="Verdana,Helvetica" color="white" size="2"><b>Datos de la operación</b></font></td>
	</tr>
	<tr HEIGHT="20">
		<td COLSPAN="3"></td>		
	</tr>
	<tr HEIGHT="30">
		<td></td>
		<td width="170">Ingrese la clave anterior</td>
		<td align="left">
			<input TYPE="password" STYLE="HEIGHT: 22px; WIDTH: 123px" SIZE="10" ALIGN="right" name="txtClaveAnterior">
		</td>
	</tr>
	<tr HEIGHT="30">
		<td></td>
		<td width="170">Ingrese la clave nueva</td>
		<td align="left">
			<input TYPE="password" STYLE="HEIGHT: 22px; WIDTH: 123px" SIZE="10" ALIGN="right" name="txtClaveNueva">
		</td>
	</tr>
	<tr HEIGHT="30">
		<td></td>
		<td width="170">Reingrese la clave nueva</td>
		<td align="left">
			<input TYPE="password" STYLE="HEIGHT: 22px; WIDTH: 123px" SIZE="10" ALIGN="right" name="txtClaveNuevaR">
		</td>
	</tr>
	<tr HEIGHT="50" align="right">
		<td COLSPAN="3"><img border="0" id="tdAceptar" onclick src="../images/BotonAceptar.jpg" style="cursor: hand" WIDTH="70" HEIGHT="20">&nbsp;&nbsp;&nbsp;&nbsp;</td>
	</tr>
</table>
</form>
</body>
<!--#INCLUDE virtual="/compartido/Rutinas.htm" -->
</html>
