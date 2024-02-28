<%@ Language=VBScript %>
<%
   Dim sDescription   
   Dim sDetalle
   Dim sTitulo
   
   sTitulo = Request("Titulo")
   sDetalle = Request("detalle")
   If sTitulo = "" Then
		sTitulo = "Información"
   End If   
	
	Select Case cInt(0 & Request("Tipo"))
	Case 0	'Nada
		sDetalle= sDetalle
	Case 1	'Bloque de la BD
		sDetalle = "Debido a la ejecución de algunos procesos en la matriz, ha sido necesario bloquear la base de datos por unos instantes.<br>Intente nuevamante esta operación en unos minutos más."
	
   End Select
	
	'sDetalle = sDetalle & Request("Detalle")	
   sDetalle = sDetalle & Request.Form("sHTML")
%>
<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<title>AFEX Ltda.</title>
<link rel="stylesheet" type="text/css" href="../Estilos/Principal.css">
<script language="vbscript" id="vbMetodos">
<!--
	Dim sDisplay
	sDisplay = "none"
-->
</script>
</head>
<body>

<table id="tabError" class="bordeinactivo" cellspacing="0" cellpadding="3" style="LEFT: 5px; POSITION: relative; TOP: 5px; WIDTH: 500px">
<tr>
	<td colspan="3" class="tituloinactivo"><%=sTitulo%></td>
</tr>
<tr height="8"><td></td></tr>
<tr>
	<td>
		<table cellspacing="0" cellpadding="0" class style="BORDER-BOTTOM: gray 2px solid; BORDER-LEFT: gray 2px solid; BORDER-RIGHT: gray 2px solid; BORDER-TOP: gray 2px solid;    TEXT-DECORATION: blink" width="35" height="30">
		<tr>
			<td align="middle" style="COLOR: gray; FONT-FAMILY: courier; FONT-SIZE: 20pt; FONT-WEIGHT: bold">i</td>
		</tr>
		</table>
	</td>
	<td style="COLOR: black"><%=sDetalle%></td>
	<td></td>
</tr>

<tr height="10"><td></td></tr>
<tr id="trDetalle" style="COLOR: black; DISPLAY: none">
	<td colspan="3">
	<table id="tabDetalle" style="BORDER-BOTTOM-COLOR: black; BORDER-LEFT-COLOR: black; BORDER-RIGHT-COLOR: black; BORDER-TOP: silver 1px solid" width="100%">
	<tr><td>
		<%=sDetalle%>
	</td></tr>
	</table>
	</td>
</tr>
</table>

<script>
	
	Sub tdBoton_onClick()
		
		If sDisplay = "none" Then 
			sDisplay = ""			
			tdboton.innerText = "<< Detalle"
		Else
			sDisplay = "none"
			tdboton.innerText = "Detalle >>"
		End If	
		
		window.trDetalle.style.display = sDisplay
	End Sub
	
</script>
</body>
</html>
