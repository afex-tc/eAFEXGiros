<%@ LANGUAGE = VBScript %>
<!--#INCLUDE virtual="/Compartido/Rutinas.asp" -->
<%
	'Se asegura que la página no se almacene en la memoria cache
	Response.Expires = 0
	If Session("CodigoCliente") = "" Then
		response.Redirect "../Compartido/TimeOut.htm"
		response.end
	End If
	
	dim sSQL, rs, sMensaje, sFono		
	 
	On Error Resume Next
	
	sFono = Request.Form("txtFono")
	if sFono = "" then sFono = "null"

	' agrega el beneficiario
	sSQL = "insertarbeneficiario " & evaluarstr(Request.Form("txtCodigoCliente")) & ", " & evaluarstr(Request.Form("txtNombres")) & ", " & _
									evaluarstr(Request.Form("txtApellidos")) & ", " & evaluarstr(Request.Form("cbxPais")) & ", " & _
									evaluarstr(Request.Form("cbxCiudad")) & ", " & sFono
	
	set rs = ejecutarsqlcliente(session("afxCnxAFEXpress"), sSQL)	
	If Err.number <> 0 Then
		Set rs = Nothing
		Response.Write "Error al agregar el beneficiario. " & err.Description 
		Response.End 
	End If	
	Set rs = Nothing
	Response.Redirect "agregarbeneficiario.asp?Cliente=" & Request.Form("txtCodigoCliente")
%>