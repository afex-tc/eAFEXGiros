<%@ LANGUAGE = VBScript %>
<%
	Dim sCliente
	
	sCliente = Request("Cliente")

%>

<html>
<head>

<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<title>AFEX</title>
<link rel="stylesheet" type="text/css" href="../Estilos/Principal.css">
<script ID="clientEventHandlersVBS" LANGUAGE="VBScript">
<!--
	Function imgPrincipal_onclick()
		window.open "../Cabecera.htm", "Cabecera"
		window.open "../Menu.htm", "Menu"
		window.open "../Principal.htm", "Principal"
	End Function

	Function CambiarCursor(ByVal sControl)
		document.all.item(sControl).style.cursor = "Hand"		
	End Function

-->
</script>
</head>
<body bgcolor="steelblue" COLOR="#66cc99">
<h1 STYLE="COLOR: #407db4; FONT-SIZE: 150pt; HEIGHT: 151px; LEFT: 65px; POSITION: absolute; TOP: -96px; WIDTH: 412px">AFEX</h1>
<h1 STYLE="COLOR: #3d78ad; FONT-SIZE: 80pt; LEFT: -6px; POSITION: absolute; TOP: -50px">En línea</h1>
<h1 STYLE="COLOR: #205987; FONT-FAMILY: Verdana; FONT-SIZE: 30pt; LEFT: 17px; POSITION: absolute; TOP: 8px">En línea</h1>
<h1 STYLE="COLOR: aliceblue; FONT-FAMILY: Verdana; FONT-SIZE: 30pt; LEFT: 14px; POSITION: absolute; TOP: 5px">En línea</h1>
<marquee STYLE="HEIGHT: 34px; LEFT: 251px; POSITION: absolute; TOP: 59px; WIDTH: 424px" BEHAVIOR="slide" SCROLLAMOUNT="1" SCROLLDELAY="30" DIRECTION="up">
	<a ALIGN="Right" href target="Principal" STYLE="COLOR: #205987; FONT-SIZE: 14pt; HEIGHT: 25px; LEFT: 0px; TEXT-DECORATION: none; TOP: 0px; WIDTH: 403px" title="Actualizar datos de <%=sCliente%>"><%=sCliente%></a>
</marquee>
<marquee STYLE="HEIGHT: 34px; LEFT: 249px; POSITION: absolute; TOP: 57px; WIDTH: 424px" BEHAVIOR="slide" SCROLLAMOUNT="1" SCROLLDELAY="30" DIRECTION="up">
	<a ALIGN="Right" href target="Principal" STYLE="COLOR: aliceblue; FONT-SIZE: 14pt; HEIGHT: 25px; TEXT-DECORATION: none; WIDTH: 403px" title="Actualizar datos de <%=sCliente%>"><%=sCliente%></a>
</marquee>
<p>
<object height="20" id="objTab" style="HEIGHT: 20px; LEFT: 0px; POSITION: absolute; TOP: 60px; WIDTH: 187px" type="text/x-scriptlet" VIEWASTEXT><param NAME="Scrollbar" VALUE="0"><param NAME="URL" VALUE="http:../Scriptlets/Tab.htm"></object>
</p>
<!--<p><IMG alt=Principal border=no height=20 id=imgPrincipal src="../images/BotonHome.jpg" style="CURSOR: hand; LEFT: 0px; POSITION: absolute; TOP: 60px" width=97 ><A href="..\Contactenos.asp" target=Principal >	<IMG alt=Contáctenos border=no height=20 src="../images/BotonContactenosCliente.jpg" style="CURSOR: hand; LEFT: 98px; POSITION: absolute; TOP: 60px" width=97 ></A></p>-->
</body>
<script language="vbscript">

	Sub window_onload()
		objtab.agregar "Principal", "../Cliente/Principal.asp", "Principal", 2
		objtab.agregar "Contáctenos", "../Contactenos.asp", "Principal", 2
	End Sub
	
</script>
</html>
