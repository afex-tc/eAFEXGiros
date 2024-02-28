<%@ Language=VBScript %>
<!--#INCLUDE virtual="/afexvisorap/Compartido/_nocache.asp" -->
<%
	Dim sTitulo
	
	sTitulo = "Inicio de Sesion"
%>
<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<title>VisorAP - Inicio Sesion</title>
<link rel="stylesheet" type="text/css" href="Compartido/VisorAP.css">
</head>
<script language="vbscript">
	Sub window_onload()
		frmLogin.us.focus 
	End Sub

	Sub KeyEnter
		If window.event.keyCode <> 13 Then Exit Sub
		If Trim(frmLogin.us.value) <> Empty And Trim(frmLogin.pw.value) <> Empty Then IniciarSesion			
	End Sub
		
	Sub IniciarSesion

		<% Select Case Session("TipoSesion") %>
		<%	Case 1  %>
				frmLogin.action = "Login.asp"
								
		<% Case Else %>
				frmLogin.action = "LoginID.asp"
		<% End Select %>
		frmLogin.submit
		frmLogin.action = ""
	End Sub
</script>
<!--#INCLUDE virtual="/afexvisorap/Compartido/Boton.htm" -->
<body onKeyPress="KeyEnter">
<form id="frmLogin" method="post">
<input type="hidden" name="ap" value="6">
<table cellpadding="0" cellspacing="0" style="border: 1px solid silver; height: 130px; width: 180px">
	<tr style="color: white; height: 20px; background-color: '#00AAAA'; '#6699ff'; font-weight: bold; font-family: MS Sans Serif; font-size: 8pt"><td>&nbsp;<%=sTitulo%></td></tr>
	<tr align="center">
		<td>
		<table cellpadding="0" cellspacing="10">
			<tr><td>Usuario&nbsp;<input type="textbox" maxlength="12" id="us" name="us" style="width: 80px"></td></tr>
			<tr><td>Password&nbsp;<input type="password" maxlength="8" id="pw" name="pw" style="width: 70px"></td></tr>
			<tr height="20px" align="center">
				<td class="boton" onClick="IniciarSesion" onMouseOver="MouseOver()" onMouseOut="MouseOut()">Iniciar Sesion</td>
			</tr>
		</table>
		</td>
	</tr>
</table>
</form>
</body>

</html>
