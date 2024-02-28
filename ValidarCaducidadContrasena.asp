<%@ Language=VBScript %>
<!--#INCLUDE virtual="/Compartido/rutinas.asp" -->
<!--#INCLUDE virtual="/Compartido/errores.asp" -->
<%
	dim sSQL
	dim rsContrasena
	dim sMensaje
	dim iTipoPantalla
	dim sContrasena

	iTipoPantalla = request("Tipo")
	if iTipoPantalla = "" then iTipoPantalla = 1 ' verificar contraseña
	
	if iTipoPantalla = 1 then
		' verifica si se debe cambiar la contraseña
		sSql = " exec validarcaducidadcontraseña " & EvaluarStr(Trim((Session("NombreUsuarioOperador"))))
		Set rsContrasena = EjecutarSQLCliente(session("afxCnxAFEXpress"),sSql)
		If Err.Number <> 0 Then
			MostrarErrorMS "Validar caducidad contraseña"
		End If
		If rsContrasena("caducidad") Then
			iTipoPantalla = 2 ' se debe actualizar contraseña
		
		else 
			iTipoPantalla = 5 ' no se debe actualizar la contraseña
		End If	
	end if

	
	if iTipoPantalla = 3 then ' grabar actualización
		' actualiza la nueva contraseña
		sSQL = " exec actualizarcontraseña " & evaluarstr(Session("NombreUsuarioOperador")) & ", " & evaluarstr(request("CT"))
		set rs = ejecutarsqlcliente(session("afxCnxAFEXpress"), sSQL)
		if err.number <> 0 then
			MostrarErrorMS "Actualizar contraseña"
		end if
		Session("ContrasenaOperador") = request("CT")
		iTipoPantalla = 4 ' contraseña actualizada sin problemas
	end if
%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<title>Usuario</title>
<link rel="stylesheet" type="text/css" href="Estilos/Principal.css">
</HEAD>
<script language="vbscript">
<!--
	
	'window.dialogwidth = 25
	'window.dialogheight = 20
	'window.dialogtop = 100
	'window.dialogleft = 200	
	
	Sub KeyEnter()
		If window.event.keyCode <> 13 Then Exit Sub
		if txtContrasena.value <> "" and txtConfirmacion.value <> "" then
			imgAceptar_onclick
		End If
	End Sub
	
	sub imgAceptar_onclick()
		if txtContrasena.value = "" or txtConfirmacion.value = "" then exit sub
		
		if txtContrasena.value <> txtConfirmacion.value then 
			msgbox "La contraseña y su confirmación deben ser iguales.",, "AFEX"
			exit sub
		end if
		
		window.navigate "ValidarCaducidadContrasena.asp?Tipo=3&CT=" & trim(txtContrasena.value) & "&URL=<%=request("URL")%>"
	end sub
	
	sub window_onload()		
		txtContrasena.focus 	
		<%if iTipoPantalla = "4" or iTipoPantalla = "5" then%>			
			window.open "<%=replace(request("URL"), ";-;", "&")%>"
			window.close
		<%end if%>
	end sub
-->
</script>

<BODY  onKeypress="KeyEnter" LANGUAGE="VBScript">	
	<table align="center" style="LEFT: 2px; WIDTH: 310px; POSITION: relative; TOP: 2px; HEIGHT: 110px" bgcolor="#c1c1c1" cellSpacing="1" cellPadding="0" border="0">		
		<tr height="40" bgcolor="steelblue">
			<td colspan="2" vAlign=center>
				<P align=center>
					<strong>
						<font style="FONT-SIZE: 16pt; TEXT-ALIGN: center"  color=white ><%=Session("NombreOperador")%></font>
					</strong>
				</P>
			</td>
		</tr>
		<tr height="10" bgcolor="#e1e1e1">
			<td>
				<font style="FONT-SIZE: 7.5pt" scolor="white">
				&nbsp;Su password ha caducado y la debe actualizar
				</font>
			</td>
		</tr>		
		<tr>
			<td>
				<table bgcolor="#f1f1f1" border="0" cellspacing="0" cellpadding="2" width="320" style="WIDTH: 320px; HEIGHT: 110px" id=TABLE1 height=110>
					<tr>
						<td align="right" style="HEIGHT: 1px" height=1>
							Password Nueva
						</td>
						<td style="HEIGHT: 2px" height=2>
							<INPUT type="password" size="10" name="txtContrasena" style="LEFT: 2px; TOP: 2px">
						</td>
					</tr>
					<tr>
						<td align="right" style="HEIGHT: 1px" height=1>
							Confirmación
						</td>
						<td style="HEIGHT: 2px" height=2>
							<INPUT type="password" size="10" name="txtConfirmacion" style="LEFT: 2px; TOP: 2px">
						</td>
					</tr>
					<tr>
						<td colspan="2" align="middle" style="HEIGHT: 1px" height=1>
							<IMG id=imgAceptar style="CURSOR: hand" onclick ="" height=20 src="images/BotonAceptar.jpg" width=70 border=0 >							
						</td>
					</tr>
				</table>
			</td>
		</tr>
	</table>
</BODY>
</HTML>