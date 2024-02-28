<%@ Language=VBScript %>
<%
sw=request.QueryString ("sw")
%>
<!--#INCLUDE virtual="/sucursal/Constantes.asp" -->
<!--#INCLUDE virtual="/Compartido/Rutinas.htm" -->
<!--#INCLUDE virtual="/Compartido/Transferencias.asp" -->
<!--#INCLUDE virtual="/Compartido/Rutinas.asp" -->
<!--#INCLUDE virtual="/Compartido/Errores.asp" -->

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link rel="stylesheet" type="text/css" href="../Estilos/Principal.css">
<link rel="stylesheet" type="text/css" href="../Estilos/Cliente.css">
<link rel="stylesheet" type="text/css" href="../Estilos/Reportes.css">
</HEAD>
<script LANGUAGE="VBScript">
<!--
	'Variables de módulo
	'Variables para encabezado
	Dim sEncabezadoFondo, sEncabezadoTitulo
	
	sEncabezadoFondo = "Agregar"
	sEncabezadoTitulo = "Nuevo Cliente"
	sub imgAceptar_onclick()
		frmExiste.action = "http:NuevoCliente.asp"
		frmExiste.submit 
		frmExiste.action = ""
	end sub
-->
</script>
<BODY>
<!-- #INCLUDE virtual="/Compartido/Encabezado.htm" -->
<form id="frmExiste" method="post" >
<table align="left" STYLE="BORDER-BOTTOM: steelblue 2px solid; BORDER-LEFT: steelblue 2px solid; BORDER-RIGHT: steelblue 2px solid; BORDER-TOP: steelblue 2px solid;  HEIGHT: 80px; LEFT: 4px; POSITION: relative; TOP: 0px; WIDTH: 466px">
	<tr>
		<td align="center">
		<font size="2" face="verdana"  color="SteelBlue"><b>Cliente ya Existe!!</b></font><br><br>
		<img align="center" alt border="0" hspace="0" id="imgAceptar" src="../images/BotonAceptar.jpg" style="cursor: hand;" WIDTH="70" HEIGHT="20">
		</td>
	</tr>
</form>	
</BODY>
</HTML>
