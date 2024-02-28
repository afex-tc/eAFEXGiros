<%@ LANGUAGE = VBScript %>
<!--#INCLUDE virtual="/Compartido/Constantes.asp" -->
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
	Dim sAFEXpress
	Dim sCodigo
	Dim nEstadoGiro
	Dim b
	Dim sSQL
	Dim rs
	
	On Error Resume Next
	
	nEstadoGiro = cInt(0 & Request("eg"))
	
	b = ValidarBDGiros
	If Not b Then Response.Redirect "http:../compartido/informacion.asp?Tipo=1"
	
	PagarGiro

	Set rs = nothing
	'Set afx = Nothing
	'Response.Redirect "AtencionClientes.asp?Accion=1&Campo=6&Argumento=" & Request.Form("txtCodigoBeneficiario")
	Response.Redirect "DetalleGiro.asp?Codigo=" & Request.Form("txtCodigoGiro")
				
	
'Métodos	
	Sub PagarGiro()
		Dim bVoucher
		dim sPaisR
		sPaisR= Request("Pais")
		
		'bVoucher = True
		'Set afx = Server.CreateObject("AfexGiro.Giro")
		'If Err.number <> 0 Then
		'	Set afx = Nothing
		'	MostrarErrorMS "Grabar Pago Giro 1"
		'End If
		
		'If nEstadoGiro = afxEstadoGiroEnvio Then
		'	bVoucher = False
		'End If 
		
		'If Session("Categoria") = 3 And Session("CodigoAgente") <> Session("CodigoMoneyBroker") _
		'And Request.Form("txtTipoGiro") = "1" Then
		'	bVoucher = False
		'End If

		'Response.Redirect "../compartido/error.asp?description=" & _
		'Session("afxCnxAFEXpress") & ", " & Trim(Request.Form("txtCodigoGiro")) & ", " &  Session("CodigoAgente") & ", " &  _
		'					 Request.Form("txtMonedaPago") & ", " &  Request.Form("txtRetira") & ", " &  Request.Form("txtNombresRetira") & ", " &  _
		'					 Request.Form("txtApellidosRetira") & ", " &  Request.Form("txtRutRetira") & ", " &  _
		'					 Request.Form("txtPassRetira") & ", " &  Request.Form("txtPaisPassRetira") & ", " &  "WEB" & ",," & "False"

		'Giro = afx.Pagar(Session("afxCnxAFEXpress"), Trim(Request.Form("txtCodigoGiro")), Trim(Session("CodigoAgente")), _
		'					 Trim(Request.Form("txtMonedaPago")), cInt(0 & Request.Form("txtRetira")), Trim(Request.Form("txtNombresRetira")), _
		'					 Trim(Request.Form("txtApellidosRetira")), Trim(Request.Form("txtRutRetira")), _
		'					 Trim(Request.Form("txtPassRetira")), Trim(Request.Form("txtPaisPassRetira")), Trim(Session("NombreUsuarioOperador")),,bVoucher)
		
		

		' se cambia el uso del componente AFX por este procedimiento		
		sSQL = " execute PagoGiro " & EvaluarStr(Trim(Request.Form("txtCodigoGiro"))) & ", " & EvaluarStr(Trim(Session("CodigoAgente"))) & ", " & _
												EvaluarStr(Trim(Request.Form("txtMonedaPago"))) & ", " & cInt(0 & Request.Form("txtRetira")) & ", " & _
												EvaluarStr(Trim(Request.Form("txtRutRetira"))) & ", " & EvaluarStr(Trim(Request.Form("txtNombresRetira"))) & ", " & _
												EvaluarStr(Trim(Request.Form("txtApellidosRetira"))) & ", " & EvaluarStr(Trim(Request.Form("txtPassRetira"))) & ", " & _
												EvaluarStr(Trim(Request.Form("txtPaisPassRetira"))) & ", " & EvaluarStr(Trim(Session("NombreUsuarioOperador"))) & ", " & _
												EvaluarStr(trim(sPaisR))
		Set rs = EjecutarSQLCliente(Session("afxCnxAFEXpress"), sSQL)
		If Err.number <> 0 Then
			'Set afx = Nothing
			MostrarErrorMS "Grabar Pago Giro 2"
		End If						
		'If afx.ErrNumber <> 0 Then
		'	MostrarErrorAFEX afx, "Grabar Pago Giro 3"
		'End If
		'If Not Giro Then
		'	response.Redirect "../compartido/error.asp?Titulo=Grabar Pago Giro 4&description=Se produjo un error desconocido al intentar pagar el giro"
		'End If
	End Sub
 							 	
%>
<!--#INCLUDE virtual="/Compartido/Rutinas.asp" -->
<!--#INCLUDE virtual="/Compartido/Errores.asp" -->
