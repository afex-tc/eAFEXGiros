<%@ Language=VBScript %>
<!-- #INCLUDE virtual="/compartido/Constantes.asp" -->
<%
%>
<html>
<head>
<META name=VI60_defaultClientScript content=VBScript>

<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<title>AFEX</title>
<link rel="stylesheet" type="text/css" href="Estilos/Principal.css">

<SCRIPT LANGUAGE=vbscript>
<!--
	Sub Version()
		window.open "PrincipalEnglish/default.asp", "_parent"
	End Sub
-->
</SCRIPT>
</head>

<body bgcolor="#66cc99" "#107c6f" >
<h1 STYLE="FONT-SIZE: 150pt; LEFT: 58px; WIDTH: 412px; COLOR: #59c88e; POSITION: absolute; TOP: -96px; HEIGHT: 151px">AFEX</h1>
<h1 STYLE="FONT-SIZE: 50pt; LEFT: -238px; WIDTH: 1046px; COLOR: #4dc087; POSITION: absolute; TOP: -17px; HEIGHT: 104px">Associated Foreign Exchange</h1>
<h1 STYLE="FONT-SIZE: 24pt; LEFT: 15px; COLOR: #3e9b78; FONT-FAMILY: Verdana; POSITION: absolute; TOP: 7px">Associated Foreign Exchange</h1>
<!--
-->
<h1 STYLE="FONT-SIZE: 24pt; LEFT: 12px; COLOR: aliceblue; FONT-FAMILY: Verdana; POSITION: absolute; TOP: 4px">Associated Foreign Exchange</h1>
<p>
<OBJECT id=objTab 
style="LEFT: 0px; WIDTH: 569px; POSITION: absolute; TOP: 60px; HEIGHT: 20px" 
type=text/x-scriptlet height=20 VIEWASTEXT><PARAM NAME="Scrollbar" VALUE="0"><PARAM NAME="URL" VALUE="http:Scriptlets/Tab.htm"></OBJECT>
</p>
<table style="LEFT: 600px; POSITION: absolute; TOP: 60px">
<!--	<IMG id=imgUSA style="WIDTH: 35px; CURSOR: hand; HEIGHT: 19px" alt="" hspace=0 src="images/us.jpg" width=35 height=19 align=absMiddle useMap="" border=0 onClick="Version">-->
<!--	<A href="PrincipalEnglish/default.asp" target=_parent>English Version</A>-->
</table>
</body>
<script language="vbscript">

	Sub window_onload()
		objTab.agregar "Principal", "../Principal.asp", "Principal", "1"
		objTab.agregar "Giros", "../Giros.asp", "Principal", "1"
		objTab.agregar "Cambios", "../Cambios.asp", "Principal", "1"
		objTab.agregar "Transferencias", "../Transferencia.asp", "Principal", "1"
		objTab.agregar "Contáctenos", "../Contactenos.asp", "Principal", "1"
		objTab.agregar "HágaseCliente", "../HagaseCliente.asp", "Principal", "1"
		<% If Request("Opcion") <> "" Then %>
				objTab.activar "td" & cInt(0 & "<%=Request("Opcion")%>")
		<% End If %>
	End Sub

</script>
</html>
