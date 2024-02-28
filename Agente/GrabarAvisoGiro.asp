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
<%	
	Dim afx, Giro
	Dim sCodigo
	
	On Error	Resume Next
	AvisarGiro
			
	Set afx = Nothing
	'Response.Redirect "AtencionClientes.asp?Accion=1&Campo=6&Argumento=" & Request.Form("txtCodigoBeneficiario")
	'response.Redirect "javascript:history.back()"
	Response.Redirect "ListaGiros.asp?Tipo=" & afxGirosAviso & "&Agente=" & Session("CodigoAgente") & "&Pagador=" & Session("CodigoAgente")
	'"&Cliente=" & Request("Cliente")
	
'Métodos	
	Sub AvisarGiro()
		Set afx = Server.CreateObject("AfexGiro.Giro")
		If Err.number <> 0 Then
			Set afx = Nothing
			MostrarErrorMS "Grabar Aviso Giro 1"
		End If
'		Response.Redirect "../compartido/error.asp?description=" & _
'		Session("afxCnxAFEXpress") & ", " & Request.Form("txtCodigoGiro") & ", " &  _
'							 Request.Form("txtTipoAviso") & ", " &  Request.Form("txtDescripcionAviso") & ", " &  _
'							 1 & ", " &  "WB" & ", " &  Request.Form("txtParentesco") & ", " &  Request.Form("txtNombreParentesco")
		Giro = afx.Avisar(Session("afxCnxAFEXpress"), Trim(Request.Form("txtCodigoGiro")), _
							 Request.Form("txtTipoAviso"), Request.Form("txtDescripcionAviso"), _
							 Session("NombreUsuarioOperador"), Request.Form("txtParentesco"), Request.Form("txtNombreParentesco"))
		'MostrarErrorMS "Grabar Pago Giro 100"
		If Err.number <> 0 Then
			Set afx = Nothing
			MostrarErrorMS "Grabar Aviso Giro 2"
		End If						
		If afx.ErrNumber <> 0 Then
			MostrarErrorAFEX afx, "Grabar Aviso Giro 3"
		End If
		If Giro <> "Ok" Then
			response.Redirect "../compartido/error.asp?Titulo=Grabar Pago Giro 4&description=Se produjo un error desconocido al intentar avisar el giro"
		End If
	End Sub
 							 	
%>
<!--#INCLUDE virtual="/Compartido/Rutinas.asp" -->
<!--#INCLUDE virtual="/Compartido/Errores.asp" -->
