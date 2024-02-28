<%@ LANGUAGE = VBScript %>
<%
	'Se asegura que la página no se almacene en la memoria cache
	Response.Expires = 0
	If Session("CodigoCliente") = "" Then
		response.Redirect "../Compartido/TimeOut.htm"
		response.end
	End If
%>
<!--#INCLUDE virtual="/Compartido/Constantes.asp" -->
<!--#INCLUDE virtual="/Compartido/Rutinas.asp" -->
<!--#INCLUDE virtual="/Compartido/Errores.asp" -->
<%	
	Dim nNegocio
	Dim sCodigo
	Dim nTipoCliente
	Dim sNuevo, sEliminar
		
	On Error	Resume Next
	
	sNuevo = Request("Nuevo")
	sEliminar = Request("Eliminar")
	
	Session("ATCAFEXpress") = sNuevo
	Session("IdCliente") = 1
	AsociarCliente
	
	Select Case Request("Tipo")
	Case "DetalleGiro"
		Response.Redirect  "DetalleGiro.asp?AFEXpress=" & sNuevo & "&Codigo=" & Request("Giro") & "&AFEXchange=" & Request("AFEXchange") & "&Accion=" & Request("Accion")
		
	Case Else
	
		Response.Redirect "AtencionClientes.asp?Accion=" & afxAccionBuscar & "&Campo=" & afxCampoCodigoExpress & "&Argumento=" & sNuevo
	End Select
	
	'Funciones y Procedimientos	
	Sub AsociarCliente()
		Dim afx, bOk
	 	'mostrarerrorms sNuevo & ", " & sEliminar & ", " & Session("NombreUsuario")
		'Set afx = Server.CreateObject("AfexwebXP.web")
		Set afx = Server.CreateObject("AfexClienteXP.Cliente")
		bOK = afx.Asociar(Session("afxCnxAFEXpress"), sNuevo, sEliminar, Session("NombreUsuarioOperador"))
	 							 
		If Err.number <> 0 Then
			Set afx = Nothing
			MostrarErrorMS "Error en Asociar Cliente Giros"
		End If
		If afx.ErrNumber <> 0 Then
			MostrarErrorAFEX afx, "Error en Asociar Cliente Giros"
		End If
		If Not bOk Then
			response.Redirect "../compartido/error.asp?titulo=Error en Asociar Cliente Giros&description=Se produjo un error desconocido al asociar los clientes"		
		End If
		Set afx = Nothing
		'mostrarerrorms sNuevo & ", 2 " & sEliminar

	End Sub	
			
	Sub BuscarNuevo()
		Dim rs
		Set rs = BuscarCliente(nCampo, sArgumento, "", "")
			
		If Not rs.EOF Then
			Session("ATCAFEXchange") = rs("exchange")
		End If
		Set rs = Nothing
	End Sub
%>
