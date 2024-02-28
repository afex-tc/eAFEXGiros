<%@ LANGUAGE = VBScript %>
<%
	'Se asegura que la página no se almacene en la memoria cache
	Response.Expires = 0
	If Session("CodigoCliente") = "" Then
		response.Redirect "../Compartido/TimeOut.htm"
		response.end
	End If
	Session.Timeout = 20
%>
<!--#INCLUDE VIRTUAL="/Compartido/Rutinas.asp" -->
<%
	Dim sCliente, sMenu
	'BuscarCliente

	sMenu = ObtenerMenuCliente(Request("Codigo"), sCliente)
		
%>
<!--#INCLUDE VIRTUAL="/Compartido/Errores.asp" -->

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
		'window.open("../compartido/eliminarsesionusuario.asp")
		
	end sub

	
-->
</script>

<!--<frameset COLS="190,*" BORDER="0"  >-->
<frameset COLS="205,*" BORDER="0">
	<frameset ROWS="170, *">
	<frame NAME="Logo" SRC="Logo.htm" scrolling="no">
	<frame NAME="Menu" SRC="Menu.asp?Menu=<%=sMenu%>">
	</frameset>
	<frameset ROWS="80,*">
		<frame Name="Cabecera" SRC="Cabecera.asp?Cliente=<%=sCliente%>" scrolling="no">
		<frame Name="Principal" SRC="AtencionClientes.asp">
	</frameset>
</frameset>
</html>
