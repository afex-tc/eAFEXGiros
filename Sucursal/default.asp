<%@ LANGUAGE = VBScript %>
<%
	'Se asegura que la página no se almacene en la memoria cache
	Response.Expires = 0
	If Session("CodigoCliente") = "" Then
		response.Redirect "../Compartido/TimeOut.htm"
		response.end
	End If
	Session.Timeout = 40
%>

<!--#INCLUDE virtual="/Compartido/Rutinas.asp" -->
<!--#INCLUDE virtual="/Compartido/Errores.asp" -->

<%
	Dim sCliente, sMenu, nTipo
	'BuscarCliente
	sMenu = ObtenerMenuCliente(Request("Codigo"), sCliente)
	 nTipo = cInt(0 & Request("tp"))
	Select Case nTipo
	Case 98
		sMenu = "10;11"
	Case 99
		sMenu = "10;11;20;21;22;23"
	Case Else
	End Select
		
%>

<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<title>AFEX En Línea</title>
</head>
<script language="VBScript">
<!--
		
		window.resizeTo 800, 600
		window.moveTo 0, 0
		window.defaultStatus="Afex Money Web - En Línea"
	
		sub window_onunload()
			'elima la sesion en la bd
			window.showmodaldialog("../compartido/eliminarsesionusuario.asp")		
		end sub

-->
</script>
<frameset COLS="190,*" BORDER="0">
	<frame NAME="Menu" SRC="Menu.asp?Menu=<%=sMenu%>">
	<frameset ROWS="80,*">
		<frame Name="Cabecera" SRC="Cabecera.asp?Cliente=<%=sCliente%>" scrolling="NO">
		<frame Name="Principal" SRC="ListaClientes.asp">
		<noframes>
			<p>Solo puede entrar si su navegador soporta marcos.</p>
		</noframes>
	</frameset>
</frameset>
</html>
