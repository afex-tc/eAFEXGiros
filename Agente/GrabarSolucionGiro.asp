<%@ LANGUAGE = VBScript %>
<%
	'Se asegura que la página no se almacene en la memoria cache
	Response.Expires = 0
	If Session("CodigoCliente") = "" Then
		response.Redirect "../Compartido/TimeOut.htm"
		response.end
	End If
%>

<%	
	Dim afx, Giro
	Dim sCodigo
	
	On Error	Resume Next
	
	SolucionarGiro
			
	Set afx = Nothing
	'Response.Redirect "AtencionClientes.asp?Accion=1&Campo=6&Argumento=" & Request.Form("txtCodigoBeneficiario")
	Response.Redirect "DetalleGiro.asp?Codigo=" & Request.Form("txtCodigoGiro")
	
	Sub SolucionarGiro()
		Set afx = Server.CreateObject("AfexGiro.Giro")
		If Err.number <> 0 Then
			Set afx = Nothing
			MostrarErrorMS "Grabar Solucion Giro 1"
		End If
		Giro = afx.Solucionar(Session("afxCnxAFEXpress"), Request.Form("txtCodigoGiro"), _
							  Request.Form("txtDescripcionSolucion"), Session("NombreUsuarioOperador"), _
							 Request.Form("txtNombreSolucion"), Request.Form("txtApellidoSolucion"), _
							 Request.Form("txtDIreccionSolucion"), Request.Form("txtFonoSolucion"))
		If Err.number <> 0 Then
			Set afx = Nothing
			MostrarErrorMS "Grabar Solucion Giro 2"
		End If						
		If afx.ErrNumber <> 0 Then
			MostrarErrorAFEX afx, "Grabar Solucion Giro 3"
		End If
		If Not Giro Then
			response.Redirect "../compartido/error.asp?Titulo=Grabar Pago Giro 4&description=Se produjo un error desconocido al intentar solucionar el giro"
		End If
	End Sub
 							 	
%>
<!--#INCLUDE virtual="/Compartido/Rutinas.asp" -->
<!--#INCLUDE virtual="/Compartido/Errores.asp" -->
