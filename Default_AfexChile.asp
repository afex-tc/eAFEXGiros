<%@ Language=VBScript %>
<%

	Session("StringConexion") = "DSN=AFEXchange;UID=cambios;PWD=cambios;"
	
%>

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<TITLE>AFEX Ltda.</TITLE>
</HEAD>
<script language="VBScript">
<!--

		'window.resizeTo 800, 600
		'window.moveTo 0, 0
-->
</script>

<frameset ROWS="110, *" BORDER="0">	
	<frame Name="Cabecera" SRC="topChileAfex.html" scrolling="NO">
	<frame Name="Principal" SRC="HomeChileAfex.asp">	
</frameset>

<NOFRAMES>
	<P>Solo puede entrar si su navegador soporta marcos.</p>
</NOFRAMES>

<!--
<FRAMESET COLS="190,*" BORDER=0>
	<FRAME NAME="Menu" SRC="Menu.asp">
	<FRAMESET ROWS="80, *" >
		<FRAME Name="Cabecera"  SRC="Cabecera.asp" scrolling="NO">
		<FRAME Name="Principal" SRC="Principal.asp">
		<NOFRAMES>
			<P>Solo puede entrar si su navegador soporta marcos.</p>
		</NOFRAMES>
	</FRAMESET>
</FRAMESET>
-->

</HTML>
