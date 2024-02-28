<%@ Language=VBScript %>

<!--#INCLUDE virtual="/agente/Constantes.asp" -->
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
	
	sEncabezadoFondo = "Anular Giro"
	sEncabezadoTitulo = "Anular Giro"
	sub imgAceptar_onclick()
		frmAnu.action = "http:AtencionClientes.asp"
		frmAnu.submit 
		frmAnu.action = ""
	end sub
-->
</script>
<BODY>
<!-- #INCLUDE virtual="/Compartido/Encabezado.htm" -->
<form id="frmAnu" method="post" >
<table align="left" STYLE="BORDER-BOTTOM: steelblue 2px solid; BORDER-LEFT: steelblue 2px solid; BORDER-RIGHT: steelblue 2px solid; BORDER-TOP: steelblue 2px solid;  HEIGHT: 80px; LEFT: 4px; POSITION: relative; TOP: 0px; WIDTH: 466px">
	<tr>
		<td align="center">
		<font size="2" face="verdana" color="SteelBlue"><b>Su Petición de Anulación de Giro ha sido enviada.<br>
															ATENCION: ANTES DE REALIZAR CUALQUIER DEVOLUCIÓN DE DINERO AL CLIENTE<BR>
															VERIFIQUE EN LA HISTORIA(DEL GIRO) QUE EFECTIVAMENTE SE ENCUENTRE ANULADO O RECLAMADO,<BR> 
															YA QUE SERÁ SU RESPONSABILIDAD SI ESTE NO SE ALCANZA A DETENER.</b></font><br><br>
		<img align="center" alt border="0" hspace="0" id="imgAceptar" src="../images/BotonAceptar.jpg" style="cursor: hand;" WIDTH="70" HEIGHT="20">
		</td>
	</tr>
</form>	
</BODY>
</HTML>
