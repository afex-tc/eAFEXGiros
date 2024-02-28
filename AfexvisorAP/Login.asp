<%@ Language=VBScript %>
<!--#INCLUDE virtual="/afexvisorap/Compartido/_nocache.asp" -->
<!--#INCLUDE virtual="/afexvisorap/Compartido/Rutinas.asp" -->
<%
	
	Session("SesionActiva") = IniciarSesion
	'Response.Write Session("SesionActiva")
	Response.Redirect "Default.asp"

	Function IniciarSesion()
		Dim rs, sConexion 
		
		IniciarSesion = False	
		sConexion = "Provider=SQLOLEDB.1;Password=afxsqlint;User ID=intranet;Initial Catalog=intranet;Data Source=alerce;"
		
		Set rs = ValidarInicioSesion(sConexion, Request("us"), Request("pw"), Request("ap"))
		If Err.Number <> 0 Then 
			Set rs = Nothing
			Exit Function		
		End If
		If rs.EOF Then 
			Set rs = Nothing
			Exit Function		
		End If
		Session("StringOpciones") = rs("string_opciones")
		Session("Perfil") = rs("correlativo_perfil")
		Session("NombrePerfil") = rs("nombre_perfil")
		Session("NombreEmpleado") = rs("nombre_completo")
		Session("NombreUsuario") = Request("us")
		IniciarSesion = True	
	End Function
		
	Function ValidarInicioSesion(Byval Conexion, ByVal Usuario, Byval Pasword, ByVal Aplicacion)
	   Dim rs
		Dim sSQL

	   Set ValidarInicioSesion = Nothing

	   On Error Resume Next

		sSQL = "ValidarInicioSesion " & EvaluarStr(Usuario) & ", " & _
	   			EvaluarStr(Pasword) & ", " & Aplicacion
	   Set rs = EjecutarSQLCliente(Conexion, sSQL)

	   If Err.Number <> 0 Then 
			Set rs = Nothing
			Exit Function
		End If
	   
	   Set ValidarInicioSesion = rs

	   Set rs = Nothing
	End Function
	
%>
