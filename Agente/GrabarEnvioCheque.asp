<%@ LANGUAGE = VBScript %>
<%
	'Se asegura que la página no se almacene en la memoria cache
	Response.Expires = 0
	If Session("CodigoCliente") = "" Then
		response.Redirect "../Compartido/TimeOut.htm"
		response.end
	End If
%>
<!--#INCLUDE virtual="/compartido/Errores.asp" -->
<%
	
	Dim sAFEXchange, sAFEXpress
	Dim sCodigo, nTipoCliente, bExtranjero
	
	On Error	Resume Next
	'Verifica si el usuario puede grabar
	'VerificarBloqueoUsuario	
	
	sAFEXpress = Trim(Request.Form("txtExpress"))
	sAFEXchange = Trim(Request.Form("txtExchange"))
		
	If sAFEXchange = "" Then
		If Request.Form("optPersona") = "on" Then
			nTipoCliente = 1
		Else
			nTipoCliente = 2
		End If
		sAFEXchange = AgregarClienteXC(nTipoCliente)
	End If
	GrabarEnvio
	If Err.number <> 0 Then
		MostrarErrorMS "Grabar Envio Cheque 4"
	End If
	Response.Redirect "AtencionClientes.asp"

	
	Sub GrabarEnvio()
		Dim nEstado
		Dim afxTVouIngreso			
		Dim afxCheque
		Dim cMontoCheque
		Dim sMoneda
		Dim cParidad
		Dim cMontoEquivalente
		Dim iValuta
		Dim cComision
		Dim cTarifaSugerida
		Dim cTarifaCobrada
		Dim iTipoParidad
		Dim cNumeroAba	
		Dim iBancoOrigen
		Dim sCuentaOrigen
		Dim sBancoDestino
		Dim sCuentaDestino
		Dim sNombreBeneficiario
		Dim sNombreCiudadDestino
		Dim sDireccionDestino
		Dim sTiempoDespacho
		Dim sBancoInterm
		Dim sCuentaInterm
		Dim sCiudadInterm
		Dim sDireccionInterm
		Dim sMensaje
		Dim sGlosaVoucher

		cMontoCheque = ccur(Cdbl(0 & Request.Form("txtMonto")))

		sMoneda = Request.Form("cbxMoneda")
		cParidad = Cdbl(0 & Request.Form("txtParidad"))
		cMontoEquivalente = ccur(Cdbl(0 & Request.Form("txtEquivalente")))

		'iValuta = Request.Form("cbxValuta")	
		'cNumeroAba = Request.Form("txtAba")
	
		'If cNumeroAba = Empty Then
		'	cNumeroAba = 0
		'End If
	
		' verifica el tipo de paridad
		iTipoParidad = cInt(0 & request.Form("cbxRate"))
		iBancoOrigen = cLng(6)
		sCuentaOrigen = "6550628510"
		sNombreBeneficiario = Trim(" " & Request.Form("txtBeneficiario"))
		sNombreCiudadDestino = Trim(" " & Request.Form("txtCiudadB"))
	
		'**Montos**
		cComision = cCur(0 & Request.Form("txtGastoExtranjera"))
		cTarifaSugerida = cCur(0 & Request.Form("txtTarifaSugerida"))
		cTarifaCobrada = cCur(0 & Request.Form("txtTarifaCobrada"))
		If cComision = Empty Then 
			cComision = 0
		End If
		If cTarifaSugerida = Empty Then 
			cTarifaSugerida = 0
		End If
		If cTarifaCobrada = Empty Then 
			cTarifaCobrada = 0
		End If
	
	
		' arma la glosa del voucher, dependiendo del tipo de cliente
		If ucase(Request.Form("optPersona")) = "ON" Then			' Persona
			sGlosaVoucher = Trim(Request.Form("txtNombres")) & " "
			sGlosaVoucher = sGlosaVoucher & Trim(Request.Form("txtApellidoP")) & " "
			sGlosaVoucher = sGlosaVoucher & Trim(Request.Form("txtApellidoM"))
	
		Else		'Empresa
			sGlosaVoucher = Trim(Request.Form("txtRazonSocial"))		
		End If
		sGlosaVoucher = Trim(sGlosaVoucher) & " (AFEX Money Web)"
		sGlosaVoucher = left(sGlosaVoucher, 50)

		'Chequeencia para AFEXchange							  
		'Set afxCheque = Server.CreateObject("AfexWebXP.Web")
		Set afxCheque = Server.CreateObject("AFEXProducto.Cheque")
		If Err.number <> 0 Then
			Set afxCheque = Nothing		
			MostrarErrorMS "Grabar Envio Cheque 1"
		End If

		'response.Redirect "../compartido/error.asp?description=" & _
		'Session("afxCnxAFEXchange") & ", " & Session("afxCnxAFEXpress")  & ", " &  _
		'						  iBancoOrigen & ", " &  sAFEXchange & ", " &  _
		'						  cMontoCheque & ", " &  Date & ", " &  1 & ", " &  0 & ", " &  8 & ", " &  0 & ",, " &  sCuentaOrigen & ", " &  _
		'						   ", " &  sBancoDestino & ", " &  sCuentaDestino & ", " &   sNombreBeneficiario & ", " &  _
		'						  cNumeroAba & ",, " & sNombreCiudadDestino & ", " &  sDireccionDestino & ", " & _
		'						  sTiempoDespacho & ", " &  sMensaje & ", " &  _
		'						  sMoneda & ", " &  "USD" & ", " &  cParidad & ", " &  _
		'						  cMontoEquivalente & ", " &  cComision & ", " &  Session("CodigoAgente") & ", " &  _
		'						  "WEB" & ", " &  iTipoParidad & ", " &  cTarifaSugerida & ", " &  _
		'						  cTarifaCobrada & ", " &  sBancoInterm & ", " &  sCuentaInterm & ", " &  sCiudadInterm & ", " &  _
		'						  sDireccionInterm & ", " &  sGlosaVoucher
		'mostrarerrorms sNombreBeneficiario
		'mostrarerrorms Session("afxCnxAFEXchange") & ", " &  Session("afxCnxAFEXpress") & ", " &  _
		'						  Session("CodigoAgente") & ", " &  afxChequeAfex & ",, " &  _
		'						  iBancoOrigen & ", " &  sCuentaOrigen & ", " &  sAFEXchange & ", " &  Date & ", " &  Date & ", " &  sMoneda & ", " &  cMontoCheque & ", " &  "USD" & ", " &  cMontoEquivalente & ", " &  _
		'						  cParidad & ", " &  iTipoParidad & ", " &  cTarifaSugerida & ", " &  cTarifaCobrada & ", " &  _
		'						  left(Request.Form("cbxPais"), 2) & ", " &  sNombreBeneficiario & ", " &  _
		'						  afxChequeSolicitado & ", " &  Session("NombreUsuario") & ", " &  cComision & ", " &  sGlosaVoucher
		Dim bOk, nCheque
		nCheque = afxCheque.Enviar(Session("afxCnxAFEXchange"), Session("afxCnxAFEXpress"), _
								  Session("CodigoAgente"), 1, "", _
								  iBancoOrigen, sCuentaOrigen, sAFEXchange, Date, Date, sMoneda, cMontoCheque, "USD", cMontoEquivalente, _
								  cParidad, iTipoParidad, cTarifaSugerida, cTarifaCobrada, _
								  left(Request.Form("cbxPaisB"), 2), sNombreBeneficiario, _
								  8, Session("NombreUsuario"), cComision, sGlosaVoucher)
		'mostrarerrorms "13"
		If Err.number <> 0 Then
			Set afxCheque = Nothing		
			MostrarErrorMS "Grabar Envio Cheque 2"
		End If
		If afxCheque.ErrNumber <> 0 Then
			MostrarErrorAFEX afxCheque, "Grabar Envio Cheque 3"
		End If		
		
		If nCheque = 0 Then
			Set afxCheque = Nothing				
			response.Redirect "http:../compartido/error.asp?Titulo=Grabar Envio Cheque 6&description=Se produjo un error desconocido y no se pudo enviar el Cheque"
		End If
		Set afxCheque = Nothing			
		'Fin Chequeencia para AFEXchange
	End Sub	


%>
<!--#INCLUDE virtual="/compartido/Rutinas.asp" -->