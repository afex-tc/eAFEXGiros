<%@ Language=VBScript %>
<!--#INCLUDE virtual="/afexvisorap/Compartido/_nocache.asp" -->
<!--#INCLUDE virtual="/afexvisorap/Compartido/Rutinas.asp" -->
<%
	
	Session("SesionActiva") = IniciarSesion	
	Response.Redirect "Default.asp"

	Function IniciarSesion()
		Dim rs, sConexion, afxCliente 
		
		IniciarSesion = False	
		sConexion = "Provider=SQLOLEDB.1;Password=afxsqlcor;User ID=corporativa;Initial Catalog=corporativa;Data Source=alerce;"
		
		'Conexion
		Set afxCliente = Server.CreateObject("AfexCorporativo.Cliente")
		Set rs = afxCliente.ObtenerCliente(sConexion, Request("us"), Request("pw"))
		
		'Control de errores
		If afxCliente.ErrNumber <> 0 Or Err.number Then 
			Set rs = Nothing
			Set afxCliente = Nothing
			Exit Function		
		End If
		
		'Si no encontró el cliente
		If rs.EOF Then 
			Set rs = Nothing
			Set afxCliente = Nothing
			Exit Function		
		End If
		
		Session("TipoCliente") = rs("tipo")
		Session("CodigoSC") = rs("codigo_agente")
		Session("NombreSC") = Trim(rs("nombre"))
		Session("AliasSC") = Replace(Session("NombreSC"), "AFEX", "")
		Session("AliasSC") = Trim(Session("AliasSC"))
		Select Case Session("TipoCliente")
		Case 4 ' Sucursal 
				Session("TipoSesion") = 1	'Sesion Sucursal
		Case Else 
				Session("TipoSesion") = 2	'Sesion Cliente
		End Select
				
		'Fin
		IniciarSesion = (Session("TipoCliente") <> 4)
		Set afxCliente = Nothing
		Set rs = Nothing
	End Function
		
	
%>