<%@ LANGUAGE = VBScript %>
<%
	'option explicit	
	'Se asegura que la página no se almacene en la memoria cache
	Response.Expires = 0
	If Session("CodigoCliente") = "" Then
		response.Redirect "TimeOut.htm"
		response.end
	End If
%>
<!--#INCLUDE virtual="/compartido/Errores.asp" -->
<!--#INCLUDE virtual="/compartido/Constantes.asp" -->
<%
	Dim rs, afxCl
	Set rs = ObtenerCliente(afxCampoCodigo, Session("CodigoCliente"))
	'mostrarerrorms afxCampoCodigo & ", " & Session("CodigoCliente")
	
	'Response.Redirect "Error.asp?Titulo=Cambio de Clave&description=" & rs("password") & ", " & Request.Form("txtClaveAnterior")
	If Trim(rs("password")) <> Trim(Request.Form("txtClaveAnterior")) Then
		Set rs = Nothing
		Response.Redirect "Error.asp?Titulo=Cambio de Clave&description=La clave anterior no coincide con la de nuestros registros"
	End If
	Set rs = Nothing
	
	Set afxCl = Server.CreateObject("AfexCorporativo.Cliente")
	
	bOk = afxCl.Actualizar(Session("AfxCnxCorporativa"), Session("CodigoCliente"), "WB",,,,,,,,,,,,,,,,,,,,, _
				Trim(Request.Form("txtClaveNueva")))
	If Err.number <> 0 Then
		Set afxCl = Nothing		
		MostrarErrorMS "Grabar Cambio de Clave 1"
	End If
	If afxCl.ErrNumber <> 0 Then
		MostrarErrorAFEX afxCl, "Grabar Cambio de Clave 2"
	End If		
	Set afxCl = Nothing			
	If Not bOk Then
		response.Redirect "../compartido/error.asp?Titulo=Grabar Cambio de Clave 3&description=No se pudo efectuar el cambio de clave"
	End If
	Set afxCl = Nothing

	Response.Redirect "Resultado.asp"	
%>
<!--#INCLUDE virtual="/compartido/Rutinas.asp" -->