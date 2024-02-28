<%@ LANGUAGE = VBScript %>
<%
	'option explicit	
	'Se asegura que la página no se almacene en la memoria cache
	Response.Expires = 0
	If Session("CodigoCliente") = "" Then
		response.Redirect "../Compartido/TimeOut.htm"
		response.end
	End If
%>
<!--#INCLUDE virtual="/Compartido/Rutinas.asp" -->
<%
	Dim sCliente, sMenu
	'BuscarCliente
	sMenu = ObtenerMenuCliente(Request("Codigo"), sCliente)	
	session.Timeout = 20
	

'Este cambio de ODBC es sólamente para pruebas en explotación.
'Una vez terminadas las pruebas se deben comentar estas líneas
'Inicio
' 	Session("afxCnxCorporativa") = "DSN=wAFEXCorporativa;UID=corporativa;PWD=afxsqlcor;"
'	Session("afxCnxAFEXchange") = "DSN=wAFEXchange;UID=cambios;PWD=cambios;"
'	Session("afxCnxAFEXpress") = "DSN=wAFEX_giros;UID=giros;PWD=giros;"	
'	Session("afxCnxAFEXweb") = "DSN=wAFEXweb;UID=cambios;PWD=cambios;"
'Fin 
	
%>
<!--#INCLUDE virtual="/Compartido/Errores.asp" -->

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<TITLE>AFEX En Línea</TITLE>
</HEAD>
<script language="VBScript">
<!--

		window.resizeTo 800, 600
		window.moveTo 0, 0
-->
</script>
<FRAMESET COLS="190,*" BORDER=0>
	<FRAME NAME="Menu" SRC="Menu.asp?Menu=<%=sMenu%>">
	<FRAMESET ROWS="80,*" >
		<FRAME Name="Cabecera"  SRC="Cabecera.asp?Cliente=<%=sCliente%>" scrolling="NO">
		<FRAME Name="Principal" SRC="Principal.asp">
		<NOFRAMES>
			<P>Solo puede entrar si su navegador soporta marcos.</p>
		</NOFRAMES>
	</FRAMESET>
</FRAMESET>
</HTML>
