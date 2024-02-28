<%@ Language=VBScript %>
<%
   Dim sError 
   Dim i 
   Dim sDescription   
   Dim sDetalle
   Dim sTitulo
   
   sTitulo = Request("Titulo")
   If sTitulo = "" Then
		sTitulo = "Error"
   End If   
   
   sError = Split(Request("Description"), "^", 2)   
   For i = 0 To UBound(sError)
      If i = 0 Then
         sDescription = sError(i)
      Else
         sDetalle = sDetalle & sError(i)
      End If
   Next
   'sDescription = sError(0)
   sDetalle = replace(sDetalle, "^", "<br>")
   Err.Clear 
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
<body onload="Javascript:history.go(1);" onunload="Javascript:history.go(1);">

<table id="tabError" class="bordeinactivo" cellspacing="0" cellpadding="3" style="LEFT: 5px; POSITION: relative; TOP: 5px; WIDTH: 400px">
<tr>
	<td colspan="3" class="tituloinactivo"><%=sTitulo%></td>
</tr>
<tr height="8"><td></td></tr>
<tr>
	<td>
		<table cellspacing="0" cellpadding="0" class style="BORDER-BOTTOM: gray 2px solid; BORDER-LEFT: gray 2px solid; BORDER-RIGHT: gray 2px solid; BORDER-TOP: gray 2px solid;    TEXT-DECORATION: blink" width="35" height="30">
		<tr>
			<td align="middle" style="COLOR: gray; FONT-FAMILY: arialblack; FONT-SIZE: 20pt; FONT-WEIGHT: bold">X</td>
		</tr>
		</table>
	</td>
	<td style="COLOR: black"><%=sDescription%></td>
	<td></td>
</tr>
<tr>
	<td colspan="2" width="100%"></td>
	<td>
		<table id="tabDetalle" class="bordeinactivo" style="BORDER-BOTTOM-COLOR: black; BORDER-LEFT-COLOR: black; BORDER-RIGHT-COLOR: black; BORDER-TOP-COLOR: black; CURSOR: hand" width="100">
		<tr>
			<td id="tdBoton" align="middle" onmouseover="window.tdBoton.style.color='blue'" onmouseout="window.tdBoton.style.color='black'" style="COLOR: black">Detalle &gt;&gt;</td>
		</tr>
		</table>
	</td>
</tr>
<tr height="10"><td></td></tr>
<tr id="trDetalle" style="COLOR: black; DISPLAY: none">
	<td colspan="3">
	<table id="tabDetalle" style="BORDER-BOTTOM-COLOR: black; BORDER-LEFT-COLOR: black; BORDER-RIGHT-COLOR: black; BORDER-TOP: silver 1px solid" width="100%">
	<tr><td>
		Este detalle de los errores ocurridos es información valiosa para AFEXinformática <br><br>
		ErrorAFX <%=Request("Number")%> en <%=Request("Source")%> <br> <%=sDescription%><br>
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
