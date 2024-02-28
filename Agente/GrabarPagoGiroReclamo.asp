<%@ Transaction=required LANGUAGE = VBScript %>
<%
	'Se asegura que la página no se almacene en la memoria cache
	Response.Expires = 0
	If Session("CodigoCliente") = "" Then
		response.Redirect "../Compartido/TimeOut.htm"
		response.end
	End If
%>
<%	
	Dim afx, Giro, afxPago
	Dim sAFEXpress
	Dim sCodigo
	
	On Error	Resume Next

	SolucionarGiro	
	Set afx = Nothing	
	PagarGiro

	Set afxPago = Nothing
	'Response.Redirect "AtencionClientes.asp?Accion=1&Campo=6&Argumento=" & Request.Form("txtCodigoBeneficiario")
	Response.Redirect "DetalleGiro.asp?Codigo=" & Request.Form("txtCodigoGiro")
				
	
	'Métodos	
	Sub PagarGiro()
		Dim bVoucher
		
		bVoucher = True
		Set afxPago = Server.CreateObject("AfexGiro.Giro")
		If Err.number <> 0 Then
			Set afxPago = Nothing
			MostrarErrorMS "Grabar Pago Giro 1"
		End If
		
		If Session("Categoria") = 3 And Session("CodigoAgente") <> Session("CodigoMoneyBroker") Then
			bVoucher = False
		End If
'		Response.Redirect "../compartido/error.asp?description=" & _
'		Session("afxCnxAFEXpress") & ", " & Request.Form("txtCodigoGiro") & ", " &  Session("CodigoAgente") & ", " &  _
'							 Request.Form("txtMonedaPago") & ", " &  Request.Form("txtRetira") & ", " &  Request.Form("txtNombresRetira") & ", " &  _
'							 Request.Form("txtApellidosRetira") & ", " &  Request.Form("txtRutRetira") & ", " &  _
'							 Request.Form("txtPassRetira") & ", " &  Request.Form("txtPaisPassRetira") & ", " &  "WEB"
		Giro = afxPago.Pagar(Session("afxCnxAFEXpress"), Request.Form("txtCodigoGiro"), Session("CodigoAgente"), _
							 Request.Form("txtMonedaPago"), cInt(0 & Request.Form("txtRetira")), Request.Form("txtNombresRetira"), _
							 Request.Form("txtApellidosRetira"), Request.Form("txtRutRetira"), _
							 Request.Form("txtPassRetira"), Request.Form("txtPaisPassRetira"), Session("NombreUsuarioOperador"),,bVoucher)
		If Err.number <> 0 Then
			Set afxPago = Nothing
			MostrarErrorMS "Grabar Pago Giro 2"
		End If						
		If afxPago.ErrNumber <> 0 Then
			MostrarErrorAFEX afxPago, "Grabar Pago Giro 3"
		End If
		If Not Giro Then
			response.Redirect "../compartido/error.asp?Titulo=Grabar Pago Giro 4&description=Se produjo un error desconocido al intentar pagar el giro"
		End If
	End Sub


	Sub SolucionarGiro()
		Set afx = Server.CreateObject("AfexGiro.Giro")
		If Err.number <> 0 Then
			Set afx = Nothing
			MostrarErrorMS "Grabar Solucion Giro 1"
		End If
		Giro = afx.Solucionar(Session("afxCnxAFEXpress"), Request.Form("txtCodigoGiro"), _
							  "Solucionado por agente pagador", Session("NombreUsuarioOperador"), _
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
