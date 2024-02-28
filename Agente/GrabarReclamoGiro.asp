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
	
	ReclamarGiro
			
	Set afx = Nothing
	Response.Redirect "AtencionClientes.asp?Accion=1&Campo=6&Argumento=" & Request.Form("txtCodigoBeneficiario")
					
	
	Sub ReclamarGiro()
		Set afx = Server.CreateObject("AfexGiroXP.Giro")
		If Err.number <> 0 Then
			Set afx = Nothing
			MostrarErrorMS "Grabar Reclamo Giro 1"
		End If
		'mostrarerrorms Session("afxCnxAFEXpress") & " 2, " & Request.Form("txtCodigoGiro") & ", " & Request.Form("txtTipoReclamo")  & ", " &  Request.Form("txtDescripcionReclamo") & ", " & Session("CodigoAgente") & ", " & Session("NombreUsuario")
		
		Giro = afx.Reclamar(Session("afxCnxAFEXpress"), Trim(Request.Form("txtCodigoGiro")), _
							 Request.Form("txtTipoReclamo"), Request.Form("txtDescripcionReclamo"),  _
							 Session("CodigoAgente"), Session("NombreUsuarioOperador"))		
		If Err.number <> 0 Then
			Set afx = Nothing
			MostrarErrorMS "Grabar Reclamo Giro 2"
		End If						
		If afx.ErrNumber <> 0 Then
			MostrarErrorAFEX afx, "Grabar Reclamo Giro 3"
		End If
		If Not Giro Then
			response.Redirect "../compartido/error.asp?Titulo=Grabar Pago Giro 4&description=Se produjo un error desconocido al intentar reclamar el giro"
		End If
	End Sub
 							 	
%>
<!--#INCLUDE virtual="/Compartido/Rutinas.asp" -->
<!--#INCLUDE virtual="/Compartido/Errores.asp" -->
