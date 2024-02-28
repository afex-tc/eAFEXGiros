<%@ Language=VBScript %>
<!--#INCLUDE virtual="/afexvisorap/Compartido/_nocache.asp" -->
<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<title>VisorAP</title>
<link rel="stylesheet" type="text/css" href="Compartido/VisorAP.css">
</head>
<script language="vbscript">	
	<% If Session("SesionActiva") Then %>
			'window.close 
			'window.open "VisorAP.asp", Null, "height=420,width=256,status=no,toolbar=no,menubar=no, scrollbars=yes, location=no, channelmode=no, resizable=yes"
			window.navigate "VisorAP.asp"
	<% End If %>
</script>
<body bgproperties="fixed" style="margin: 0 0 0 0" scroll="no" background="images/Afex%20VisorAP.gif" style="background-repeat: no-repeat">
<table border="0" cellspacing="0" cellpadding="0" style="width: 100%">
<!--<tr></td><img src="images/Afex VisorAP.gif" style WIDTH="792" HEIGHT="43"></td></tr>-->
<tr height="39px" align="right"><td><img src="images/Afex.jpg" WIDTH="30" HEIGHT="30">&nbsp;</td></tr>
<% Select Case Session("TipoSesion") %>
<% Case 1		'Sesion Sucursal %>
		<tr align="left"><td class="tituloXP2">&nbsp;Personalizar Sesion &nbsp;para <%=Session("NombreSC")%></td></tr>
<%	Case Else	'Sesion Cliente %>
		<tr align="left"><td class="tituloXP">&nbsp;Inicio de Sesion</td></tr>
<% End Select %>
<tr align="center">
	<td>
		<object align="center" id="objMenu" style="HEIGHT: 150px; LEFT: -10px; POSITION: relative; TOP: 0px; WIDTH: 190px" type="text/x-scriptlet" border="0" VIEWASTEXT>
			<param NAME="Scrollbar" VALUE="0">
			<param NAME="URL" VALUE="InicioSesion.asp">
		</object>
	</td>
</tr>
<tr>
	<td>
<!--	
	<applet  align="texttop" width="630" height="17" codebase="http://www.economiaynegocios.cl/noticias/lib" code="hscroll.class" VIEWASTEXT>
     <param name="bgcolor" value="#00AAAA">
     <param name="fontstyle" value="">
     <param name="hlcolor" value="240,240,150">
     <param name="Notice" value="Valorfuturo.com">
     <param name="pausemargin" value="5">
     <param name="scrolldelay" value="10">
     <param name="direction" value="left">
     <param name="scrolljump" value="1">
     <param name="fontface" value="MS Sans Serif">
     <param name="size" value="12">
     <param name="yoffset" value="-1">
     <param name="onsbtext" value="">
     <param name="offsbtext" value="">
     <param name="loadwhere" value="_blank">
                                
     <param name="desc0" value="|Martes  22 de marzo. Dólar Observado:">
     <param name="textcolor0" value="#FFFFFF">
                                
     <param name="desc1" value="|$586,86  ">
     <param name="textcolor1" value="#FFffCC">
     <param name="fontface1" value="Verdana">
     <param name="size1" value="16">
     <param name="fontstyle1" value="bold">
                                
     <param name="desc2" value="|UF:">
     <param name="textcolor2" value="#FFFFFF">
                                
     <param name="desc3" value="|$17.203,78">
     <param name="textcolor3" value="#FFffCC">
                                
     <param name="desc4" value="|UTM">
     <param name="textcolor4" value="#FFFFFF">
                                
     <param name="desc5" value="|$30.186">
     <param name="textcolor5" value="#FFffCC">
                                
     <param name="desc6" value="|IPSA">
     <param name="textcolor6" value="#FFFFFF">
                                
     <param name="desc7" value="|1.950,11 ">
     <param name="textcolor7" value="#FFffCC">
                                
   </applet>
-->
	</td>
</tr>
</table>
<!--#INCLUDE virtual="/afexvisorap/Compartido/Pie.htm" -->
</body>
</html>
