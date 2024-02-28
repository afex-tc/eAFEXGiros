<%@ Language=VBScript %>
<!--#INCLUDE virtual="/Compartido/Constantes.asp" -->
<%

	ValidarCliente

	Sub ValidarCliente()
		Dim afxCliente
		Dim rs, sString, sSQL
		Dim sUsuario, sPassword
		Dim sCodigo, nTipo, sCodigoSistemas, nCampo		
		
		' JFMG 21-08-2009 se agrega para validar si se muestra la página o no
		Dim swMostrarPagina 
		swMostrarPagina = 1
		' ************* FIN JFMG 21-08-2009

		'On Error Resume Next
		'Set afxCliente = Server.CreateObject("AfexCorporativo.Cliente")
		Set rs = Server.CreateObject("ADODB.recordset")
		sUsuario = Request.Form("NombreUsuario")
		sPassword = Request.Form("Password")
		'Set rs = afxCliente.ObtenerCliente(Session("afxCnxCorporativa"), Request("NombreUsuario"), Request("Password"))
		'sSQL = "SELECT * FROM Usuario WHERE codigo_usuario = '" & Request.Form("NombreUsuario") & "' AND password_usuario='" & Request.Form("Password") & "' AND codigo_agente IS NULL "
		
		' JFMG 08-07-2009 se cambia por un procedimiento almacenado
		'sSQL = "SELECT * FROM Usuario WHERE codigo_usuario = '" & Request.Form("NombreUsuario") & "' AND password_usuario='" & Request.Form("Password") & "'"
		sSQL = " exec auditoria.mostrardatosusuarioconexion " & evaluarstr(Request.Form("NombreUsuario")) & ", " & _
									evaluarstr(Request.Form("Password")) & ", " & evaluarstr(session("CodigoAgente"))
		' ********** FIN JFMG 08-07-2009
		

		if Request("tp")=99 then
			' JFMG 08-07-2009 se cambia por un procedimiento almacenado
			'sSQL = sSQL &" AND web_clientes = 1 "
			sSQL = sSQL & ", 2 "
			' ********** FIN JFMG 08-07-2009
		else
			' JFMG 08-07-2009 se cambia por un procedimiento almacenado
			'sSQL = sSQL &" AND uso_web = 1 "
			sSQL = sSQL & ", 1 "
			' ********** FIN JFMG 08-07-2009
		end If
		Set rs = EjecutarSQLCliente(Session("afxCnxAFEXpress"), sSQL)
		If Err.Number <> 0 Then
			rs.Close
			Set rs = Nothing
			'Set afxCliente = Nothing
			'MostrarErrorAFEX afxCliente, "AFEX En Línea"
			Response.Redirect "Compartido/Informacion.asp?Detalle=" & Err.Description
		End If
		If rs.EOF Then
			rs.Close
			Set rs = Nothing
			'Set afxCLiente = Nothing
			%>
				<HTML>
				<TITLE>AFEX En Línea</TITLE>
				<SCRIPT LANGUAGE="vbscript">
					'window.moveTo 1500, 1500
					'window.resizeTo 1, 1
					Msgbox "No se encontró el usuario o la password no corresponde."
					window.close
				</SCRIPT>				
				</HTML>
			<%
			'Response.Redirect "Compartido/Informacion.asp?Detalle=No se encontró el usuario o la password no corresponde"
			'Response.End 
			Exit Sub		
		End If
		'If afxCliente.ErrNumber <> 0 Then			
		'	Set rs = Nothing
		'	MostrarErrorAFEX afxCliente, "AFEX En Línea"
		'End If


		' JFMG 08-07-2009 verifica si la ip del cliente está habilitada
		' si es sucursal valida la ip del usuario
		if Session("Categoria") <> 3 and Session("Categoria") <> 4 and session("CodigoAgente") <> "ZI" then ' JFMG 02-11-2009 se agrega AFEX ESTADO para que no valide IP
			
			' JFMG 19-08-2009 si el cliente tiene autorización de ingreso, no se valida su ip
			if rs("swUsuarioAutorizadoIngreso") = 0 then
			' ********** FIN JFMG 19-08-2009 **************

				' se desglosa la ip del agente para saber si la del usuario concuerda en los 3 primeros bloques
				dim ipUsuario, ipsucursal
				
				ipusuario = split(request.servervariables("REMOTE_ADDR"), ".")
				ipsucursal = split(rs("ipagente"), ".")			


'response.write rs("swUsuarioAutorizadoIngreso") '& "-" & ipusuario  & "/" & ipsucursal 
'response.end 
			
				' verifica la ip
				if ipusuario(0) <> ipsucursal(0) or ipusuario(1) <> ipsucursal(1) or ipusuario(2) <> ipsucursal(2) then
					'Set rs = Nothing
					' si no está autorizada se sale
					%>
						<HTML>
						<TITLE>AFEX En Línea</TITLE>
						<SCRIPT LANGUAGE="vbscript">					
							Msgbox "Ud. no se encuentra AUTORIZADO para ingresar.",,"AFEX"
							window.close
						</SCRIPT>				
						</HTML>
					<%
					swMostrarPagina = 0
				end if

			' JFMG 19-08-2009 si el cliente tiene autorización de ingreso, no se valida su ip
			elseif rs("swUsuarioAutorizadoIngreso") = 2 then
				%>
				<HTML>
					<TITLE>AFEX En Línea</TITLE>
					<SCRIPT LANGUAGE="vbscript">					
						Msgbox "Ud. no se encuentra AUTORIZADO para ingresar.",,"AFEX"
						window.close
					</SCRIPT>				
				</HTML>
				<%
				swMostrarPagina = 0	
			end if
			' ********** FIN JFMG 19-08-2009 **************

		end if
		
		' inserta la auditoria de conexión
		dim rsAuditoria, sSQLAuditoria
		
		sSQLAuditoria = " exec auditoria.insertarconexionusuario " & evaluarstr(rs("codigo_usuario")) & ", " & _
										evaluarstr(session("CodigoAgente")) & ", " & _
										evaluarstr(request.servervariables("REMOTE_ADDR"))
'		response.write sSQLAuditoria
'		response.end
		set rsAuditoria = ejecutarsqlcliente(session("afxCnxAFEXpress"), sSQLAuditoria)
		if err.number <> 0 then
'			Set rsAuditoria = Nothing
			set rs = nothing
			%>
			<HTML>
			<TITLE>AFEX En Línea</TITLE>
			<SCRIPT LANGUAGE="vbscript">					
				Msgbox "Problemas con el registro. 1. " & "<%=err.description%>"
				window.close
			</SCRIPT>		
			</HTML>
			<%
			
			swMostrarPagina = 0		

		elseif rsAuditoria.eof then
'			Set rsAuditoria  = Nothing
			set rs = nothing
			%>
			<HTML>
			<TITLE>AFEX En Línea</TITLE>
			<SCRIPT LANGUAGE="vbscript">					
				Msgbox "Problemas con el registro. 2"
				window.close
			</SCRIPT>		
			</HTML>
			<%
			
			swMostrarPagina = 0
		end if

		' verifica si esta autorizado o no
		if rsAuditoria("swAutorizacion") = 0 then
			'Set rsAuditoria  = Nothing
			'set rs = nothing
			%>
			<HTML>
			<TITLE>AFEX En Línea</TITLE>
			<SCRIPT LANGUAGE="vbscript">					
				Msgbox "Ud. ya tiene una Sesión abierta en la página.",,"AFEX"
				window.close
			</SCRIPT>		
			</HTML>
			<%
			
			swMostrarPagina = 0
		end if

		' asigna el código del registro a una session
		Session("CodigoConexionUsuario") = rsAuditoria("idConexionUsuario")
		set rsAuditoria = nothing
		' ********** FIN JFMG 08-07-2009 ***************

		' JFMG 21-08-2009 se agrega para validar si se muestra la página o no
		if swMostrarPagina = 1 then
		' ************* FIN 21-08-2009 *****************************
	
			Session("NombreOperador") = rs("nombre_usuario")
			Session("NombreUsuarioOperador") = rs("codigo_usuario")
	
			' JFMG 04-12-2009 datos para enviar a AFEXchangeWEB
			Session("ContrasenaOperador") = Request.Form("Password")
			' *********** FIN JFMG 04-12-2009
	
	        ' JFMG 10-03-2011
	        Session("UsuarioAutorizadoEnviarTransferencia") = rs("autorizadoenviartransferencia")
	        ' FIN JFMG 10-03-2011
	
	' JFMG 27-04-2011 datos para mensajeria cliente
			Session("SolicitudMensajeriaClienteActiva") = rs("MensajeriaClienteActiva")
			Session("AgentePagadorMensajeriaCliente") = rs("CodigoAgentePagadorMensajeria")
			' FIN JFMG 27-04-2011
	
			'Set afxCliente = Nothing
			rs.Close
			Set rs = Nothing
			
			Dim sURL
		
			nTipo = cInt(0 & Request("tp"))
			'response.Redirect "compartido/error.asp?description=" & Session("CodigoCliente") & ", " & Session("NombreCiente")
			If nTipo = 98 Then
				sURL = "http:Sucursal/Default.asp?Codigo=" & Session("CodigoCliente") & "&tp=" & nTipo
			ElseIf nTipo = 99 Then
				sURL = "http:Sucursal/Default.asp?Codigo=" & Session("CodigoCliente") & "&tp=" & nTipo
			ElseIf Session("CodigoCliente") = 98 then
				sURL = "EnLineaAfex.asp"
			Else
				sURL = "http:Agente/Default.asp?Codigo=" & Session("CodigoCliente")
			End If
			'Response.Redirect sURL
			%>
			<script language="vbscript">
				' JFMG 19-01-2010 se agrega validación de caducidad contraseña
				If <%=nTipo%> = 98 or <%=nTipo%> = 99 Then					
					window.open "ValidarCaducidadContrasena.asp?Tipo=1&URL=<%=replace(sURL,"&", ";-;")%>", "", "top=250,height=180,width=330,left=250"
				else 
					window.open "<%=sURL%>"				
				end if
				' ******** FIN JFMG 19-01-2010 *******
			
				'window.open "<%=sURL%>"
				window.close
			</script>		
			<%

		' JFMG 21-08-2009 se agrega para validar si se muestra la página o no
		end if
		' ************* FIN 21-08-2009 ***************************** 		

	End Sub	
%>
<!--#INCLUDE virtual="/Compartido/Rutinas.asp" -->
<!--#INCLUDE virtual="/Compartido/Errores.asp" -->
