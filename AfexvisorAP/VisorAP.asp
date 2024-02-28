<%@ Language=VBScript %>
<%
	Response.Expires = 0
	Response.buffer = true	
	If Not Session("SesionActiva") Then
		response.Redirect "http:Compartido/ErrorSesion.htm"
		response.end
	End If
%>
<HTML>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<title>VisorAP</title>
<link rel="stylesheet" type="text/css" href="Compartido/VisorAP.css">
</head>
<body bgproperties="fixed" style="margin: 0 0 0 0" scroll="yes" sbackground="images/Afex VisorAP.gif" style="background-repeat: no-repeat">
<table border=0 cellspacing="0" cellpadding="0" style="width: 100%">
<!--<tr height="39px"align="right"><td><img src="images/Afex.jpg">&nbsp;<td></td></tr>
<tr align="left"><td class="tituloXP" nowrap style="font-family: MS Sans Serif; height: 15; font-size: 8pt">&nbsp;<%=Session("AliasSC")%> - <%=Session("NombreEmpleado")%></td></tr>
-->
<tr>
	<td>
		<object align="center" id="objMenu" style="margin: 0 0 0 0; HEIGHT: 390px; WIDTH: 240px" type="text/x-scriptlet" border="0"  VIEWASTEXT><param NAME="Scrollbar" VALUE="0">
			<param NAME="URL" VALUE="VisorAP/VisorAP.asp?mn=<%=Request("mn")%>&sc=<%=Request("sc")%>&fch=<%=Request("fch")%>&itv=<%=Request("itv")%>&mut=<%=Request("mut")%>">
		</object>
	<td>
</tr>
</BODY>
</HTML>
