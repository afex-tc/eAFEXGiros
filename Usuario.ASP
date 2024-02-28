<%@ Language=VBScript %>
<!--#INCLUDE virtual="/Compartido/Rutinas.htm" -->
<%
dim sSucursal
sSucursal = session("NombreCliente")
	
%>
<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<title>Usuario</title>
<link rel="stylesheet" type="text/css" href="Estilos/Principal.css">
</head>

<body onKeypress="KeyEnter" LANGUAGE="VBScript">
<script LANGUAGE="VBScript">
<!--
	'window.dialogWidth = 20
	'window.dialogHeight = 14
	'window.dialogLeft = 300
	'window.dialogTop = 70
	Sub KeyEnter()
		If window.event.keyCode <> 13 Then Exit Sub
		If Trim(txtUsuario.value) <> Empty And Trim(txtPassword.value) <> Empty Then
			imgAceptar_onclick
		End If
	End Sub
	
	Sub Window_onload()
		txtUsuario.focus
	End Sub
	
	Sub imgAceptar_onclick()
		If TRIM(txtUsuario.value) = "" Then 
			window.close 
			Exit Sub
		End If
		If TRIM(txtPassword.value) = "" Then 
			window.close 
			Exit Sub
		End If
		
		Usuario.NombreUsuario.value = LimpiarCampoTxt(txtUsuario.value)
		Usuario.Password.value  = LimpiarCampoTxt(txtPassword.value)
		txtUsuario.value = ""
		txtPassword.value = ""
		Usuario.action = "ValidarUsuario.asp?tp=<%=Request("tp")%>"
		Usuario.submit
		Usuario.action = ""
		'window.close 
	End Sub
	

	    
	sub ValidarTecla()
		dim iTecla
					
		iTecla = window.event.keyCode 		
		if iTecla = 32 or iTecla = 34 or iTecla = 39 then		
			window.event.keyCode = 0		
		end if
		
	end sub
	

		
-->
</script>

<form id="Usuario" method="post">
	<input type="hidden" name="NombreUsuario">
	<input type="hidden" name="Password">
</form>
<center>
<table style="LEFT: 2px; WIDTH: 310px; POSITION: relative; TOP: 2px; HEIGHT: 110px" bgcolor="#c1c1c1" cellSpacing="1" cellPadding="0" border="0">  
  <tr height="40" bgcolor="steelblue">
    <td colspan="2" vAlign=center>
      <P align=center><strong><font style="FONT-SIZE: 16pt; TEXT-ALIGN: center"  color=white 
     ><%=sSucursal%></font></strong></P></td>
  </tr>
  <tr height="10" bgcolor="#e1e1e1">
    <td colspan="2"><font style="FONT-SIZE: 7.5pt" scolor="white">&nbsp;Ingrese su nombre de usuario y password personal</font>  </td>
  </tr>
  <tr><td>
  <table bgcolor="#f1f1f1" border="0" cellspacing="0" cellpadding="2" width="320" style="WIDTH: 320px; HEIGHT: 110px" id=TABLE1 height=110 
     >
  <tr>
    <td align="right" style="HEIGHT: 1px" height=1 
           >Usuario</td>
    <td style="HEIGHT: 2px" height=2><input onKeypress="ValidarTecla()" id="txtUsuario" size="10"></td>
  </tr>
  <tr>
    <td align="right" style="HEIGHT: 1px" height=1 
           >Password</td>
    <td style="HEIGHT: 1px" height=1><input type="password" onKeypress="ValidarTecla()" id="txtPassword" size="10" 
            style="LEFT: 2px; TOP: 2px"></td>
  </tr>
  <tr>
		<td colspan="2" align="middle" style="HEIGHT: 1px" height=1 
         ><IMG id=imgAceptar style="CURSOR: hand" onclick ="" height=20 src="images/BotonAceptar.jpg" width=70 border=0 ></td>
  </tr>
  </table>
  </td></tr>
</table>
</center>
</body>
</html>
