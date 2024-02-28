<%@ LANGUAGE = VBScript %>
<%
	'Se asegura que la página no se almacene en la memoria cache
	Response.Expires = 0
	If Session("CodigoCliente") = "" Then
		response.Redirect "http:../Compartido/TimeOut.htm"
		response.end
	End If
%>

<%	
	Dim sSQL
	Dim rs
		
	On Error	Resume Next
	
	sSQL = " execute InsertarParidadMoneda " & EvaluarStr(request.Form("cbxMoneda")) & ", " & FormatoNumeroSQL(request.Form("txtParidadFinal")) & _
														", " & FormatoNumeroSQL(request.Form("txtRecargo")) & ", " & FormatoNumeroSQL(request.Form("txtParidad")) & _
														", " & EvaluarStr(request.Form("txtFecha")) & ", " & EvaluarStr(request.Form("txtHora")) & ", " & _
														EvaluarStr(Session("NombreUsuarioOperador"))
	
	set rs = EjecutarSQLCliente(Session("afxCnxAfexchange"), sSQL)
	if Err.number <> 0 Then
		Set rs = Nothing
		MostrarErrorMS "Grabar Paridad Moneda 1"
	End If						
	
	set rs = nothing
	response.Redirect "MantenedorParidades.asp"

 	
						 	
%>
<!--#INCLUDE virtual="/compartido/Rutinas.asp" -->
<!--#INCLUDE virtual="/compartido/Errores.asp" -->