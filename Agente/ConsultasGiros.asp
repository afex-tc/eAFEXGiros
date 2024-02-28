<%@ Language=VBScript %>
<%
	'Variables de módulo
	'Variables para encabezado
	Dim sEncabezadoFondo
	Dim sEncabezadoTitulo
	
	Dim sTitulo
	Dim nTipo
		
	sTitulo = Request("Titulo")
	
	If Trim(sTitulo) = "" Then sTitulo = "Giros"
	sEncabezadoFondo = "Consultas"
	sEncabezadoTitulo = sTitulo

%>
<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<title>AFEX</title>
<link rel="stylesheet" type="text/css" href="../Estilos/Principal.css">
<link rel="stylesheet" type="text/css" href="../Estilos/Cliente.css">
<link rel="stylesheet" type="text/css" href="../Estilos/Reportes.css">
</head>
<script LANGUAGE="VBScript">
<!--
	Sub imgAceptar_onClick()
		window.navigate "ListaGiros.asp?Tipo=1&Titulo=" + "<%=sTitulo%>"
	End Sub		

	Function imgAceptar_onMouseOver()
		window.imgAceptar.style.cursor = "Hand"		
	End Function

//-->
</script>
<body>
<!--#INCLUDE virtual="/Compartido/Encabezado.asp" -->
<br><br><br>
<marquee STYLE="HEIGHT: 346px; LEFT: 9px; POSITION: absolute; TOP: 80px; WIDTH: 511px" BEHAVIOR="slide" DIRECTION="up" SCROLLAMOUNT="50" SCROLLDELAY="100">		
<center>
<table BORDER="0" CELLSPACING="0" CELLPADDING="5" style="BORDER-BOTTOM: steelblue 2px solid; BORDER-LEFT: steelblue 2px solid; BORDER-RIGHT: steelblue 2px solid; BORDER-TOP: steelblue 2px solid; HEIGHT: 195px; WIDTH: 343px" id="tabConsulta">	
	<tr>
		<td CLASS="titulo">Datos de la consulta</td>
	</tr>
	<tr height="20"><td></td></tr>
	<tr>
		<td>	
		<table CELLSPACING="0" CELLPADDING="2" style="BORDER-BOTTOM: silver 1px solid; BORDER-LEFT: silver 1px solid; BORDER-RIGHT: silver 1px solid; BORDER-TOP: silver 1px solid; HEIGHT: 98px; WIDTH: 152px" id="tabConsulta">
        <tbody>		
		<tr>
			<td colspan="2" class="titulo" style="BACKGROUND-COLOR: silver">Periodo</td>
		</tr>
		<tr>
			<td>Desde el</td> 
			<td><input SIZE="8" VALUE="01-01-2002" id="text2" name="text2"></td>
		</tr>
		<tr>
			<td>Hasta el</td>
			<td><input SIZE="8" VALUE="01-01-2002" id="text1" name="text1" style="LEFT: 58px; TOP: 7px"></td>
			</td>
		</tr>
		</table></td></tr>
	<tr align="middle">
		<td>
			<img height="25" id="imgAceptar" src="../images/BotonAceptar.jpg" width="80">
		</td>
	</tr></tbody></table>
</center>
</marquee>		
</body></p></marquee></tr></tbody></table>
</html>
