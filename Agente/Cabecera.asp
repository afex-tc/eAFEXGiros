<%@ LANGUAGE = VBScript %>
<!--#INCLUDE virtual="/Compartido/Constantes.asp" -->
<%
	Dim sCliente, sUsuario
	
	sCliente = Request("Cliente") 
	If TRIM(Session("NombreOperador")) <> "" Then 
		sUsuario = " - " & TRIM(Session("NombreOperador"))
	Else
		sUsuario = ""
	End If

%>
 
<html>
<head>

<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<title>AFEX</title>
<link rel="stylesheet" type="text/css" href="../Estilos/Principal.css">
<script ID="clientEventHandlersVBS" LANGUAGE="VBScript">
<!--
	Function imgPrincipal_onclick()
		window.open "../AtencionClientes", "Principal"
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
<%	If Session("ModoPrueba") Then %>
		<h1 STYLE="COLOR: aliceblue; FONT-FAMILY: Verdana; FONT-SIZE: 30pt; LEFT: 14px; POSITION: absolute; TOP: 5px">En línea.</h1>
<% Else %>
		<h1 STYLE="COLOR: aliceblue; FONT-FAMILY: Verdana; FONT-SIZE: 30pt; LEFT: 14px; POSITION: absolute; TOP: 5px">En línea</h1>
<% End If %>
<p>
<object height="20" id="objTab" style="HEIGHT: 20px; LEFT: 0px; POSITION: absolute; TOP: 60px; WIDTH: 310px" type="text/x-scriptlet" VIEWASTEXT><param NAME="Scrollbar" VALUE="0"><param NAME="URL" VALUE="http:../Scriptlets/Tab.htm"></object>
</p>
<marquee STYLE="HEIGHT: 34px; LEFT: 321px; POSITION: absolute; TOP: 59px; WIDTH: 424px" BEHAVIOR="slide" SCROLLAMOUNT="1" SCROLLDELAY="30" DIRECTION="up">
	<a ALIGN="Right" href target="Principal" STYLE="COLOR: #205987; FONT-SIZE: 12pt; HEIGHT: 25px; LEFT: 0px; TEXT-DECORATION: none; TOP: 0px; WIDTH: 403px" title="<%=sCliente%>"><%=sCliente%><FONT SIZE=1>&nbsp;<%=sUsuario%></FONT></a>
</marquee>
<marquee STYLE="HEIGHT: 34px; LEFT: 320px; POSITION: absolute; TOP: 57px; WIDTH: 424px" BEHAVIOR="slide" SCROLLAMOUNT="1" SCROLLDELAY="30" DIRECTION="up">
	<a ALIGN="Right" href target="Principal" STYLE="COLOR: aliceblue; FONT-SIZE: 12pt; HEIGHT: 25px; TEXT-DECORATION: none; WIDTH: 403px" title=" <%=sCliente%>"><%=sCliente%><FONT SIZE=1>&nbsp;<%=sUsuario%></FONT></a>
</marquee>
</body>
<script language="vbscript">

	dim nSeg

	Sub window_onload()
		objtab.agregar "AtenciónClientes", "../Agente/AtencionClientes.asp", "Principal", 2
		objtab.agregar "ClienteActual", "#ATC", "Principal", 2		
		objtab.agregar "Contáctenos", "../Contactenos.asp", "Principal", 2
		
		
		nSeg = 600
		window.setInterval "CargarPagina", 1000, "vbscript"		
	End Sub
		
	Sub CargarPagina
		nSeg = nSeg - 1
		If nSeg = 0 Then
			
			window.navigate "cabecera.asp"		
		End If		
	End Sub	

	Sub objTab_OnScriptletEvent(strEventName, varEventData)

	   Select Case strEventName
	   
			Case "linkClick"
				If Right(varEventData, 3) = "ATC" Then
					window.open "../Agente/AtencionClientes.asp?Accion=<%=afxAccionClienteActual%>", "Principal"
				Else
					window.open varEventData, "Principal"
				End If
								
		End Select
		
	End Sub
	
</script>
</html>
