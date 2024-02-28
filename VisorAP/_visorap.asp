<!-- _visorap.htm -->
<%
	Response.Expires = 0
	Response.buffer = true	
	
	Dim sChecked, sCaption
	
	If cInt(0 & Request("vpr2")) = 1 Then
		sChecked = "checked"
		nHeight = 61
		sCaption = "Ocultar variaciones del día"
	Else
		nHeight = 32
		sCaption = "Mostrar variaciones del día"
	End If
%>
<html>
<head>
<title>AFEX Ltda.&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</title>
</head>
<script LANGUAGE="VBScript">

	On Error Resume Next
	Dim nVpr2, nSeg
		
	Sub window_onload()
		nSeg = 60
		<% If Request("vpr2")="" Then %>		
				window.resizeto 185, 258
		<% Else %>
				window.resizeTo 185, 316
		<% End If %>
		window.setInterval "CargarVisor", 1000, "vbscript"
	End Sub	
	
	Sub CargarVisor
		nSeg = nSeg - 1
		If nSeg = 0 Then
			window.navigate "http:_visorAP.asp?vpr2=<%=Request("vpr2")%>"		
		End If
		bSeg.innerText = nSeg		
	End Sub

	Sub ActualizarVisor		
		<% If Request("vpr2")="" Then %>		
				window.navigate "http:_visorAP.asp?vpr2=1"		
		<% Else %>
				window.navigate "http:_visorAP.asp"
		<% End If %>
	End Sub
	
</script>
<!--#INCLUDE FILE="Boton.htm" -->

<body border="0" style="margin: -20 -1 -1 -1" background="AFEX.jpg" bgcolor="#f6f6f6" style="background-repeat: no-repeat">
		<table cellspacing="2" cellpadding="1" style="BORDER-BOTTOM: lightgrey 0px solid">
			<tr height="27px" align="left">						
				<td colspan="2" width="50px"></td>				
			</tr>
			<tr>
				<td>
					<object align="left" id="objUSD" style="HEIGHT: <%=nHeight%>px; WIDTH: 169px" type="text/x-scriptlet" width="174" border="0" VIEWASTEXT><param NAME="Scrollbar" VALUE="0"><param NAME="URL" VALUE="http:VisorAP.asp?mn=usd&vpr1=1&vpr2=<%=Request("vpr2")%>&co=2&ve=2"></object>
				</td>
			</tr>
			<tr>
				<td>
					<object align="left" id="objEUR" style="HEIGHT: <%=nHeight%>px; WIDTH: 169px" type="text/x-scriptlet" width="174" border="0" VIEWASTEXT><param NAME="Scrollbar" VALUE="0"><param NAME="URL" VALUE="http:VisorAP.asp?mn=eur&vpr1=1&vpr2=<%=Request("vpr2")%>&co=3&ve=2"></object>
				</td>
			</tr>
			<tr>
				<td width="" style="background-color: white; font-size: 8pt;" class="boton" onClick="ActualizarVisor" onMouseOver="MouseOver()" onMouseOut="MouseOut()"><%=sCaption%></td>
			</tr>
			<tr>
				<td width="" style="background-color: white; font-size: 8pt;" class="boton" onClick="window.navigate 'http:_visorAP.asp?vpr2=<%=Request("vpr2")%>'" onMouseOver="MouseOver()" onMouseOut="MouseOut()">Actualizar en <b id="bSeg">60</b> seg.</td>
			</tr>
			<tr>
				<td align="center" style="background-color: LightYellow; border: 1px solid #bcbcbc; font-family: Tahoma, Arial; color: #666666 ; font-size: 8pt">
					<b>IMPORTANTE</b><br>
					* Precios referenciales para operaciones en efectivo sólo con Casa Matriz.<br>
					* Precios sujetos a cambio.<br>
				</td>
			</tr>
			<tr>
				<td align="center" style="cursor: hand; color: blue; text-decoration: underline ; font-family: Tahoma; font-size: 8pt;" onClick="window.open '../Contactenos.asp', 'Contactenos', 'height=460,width=500,top=10, left=10,status=no,toolbar=no,menubar=no,resize=yes,location=no,channelmode=no'">Llámenos al 636-9011</td>
			</tr>
		</table>
</body>
</html>