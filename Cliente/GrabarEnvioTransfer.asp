<%@ LANGUAGE = VBScript %>
<!--#INCLUDE virtual="/compartido/TimeOut.asp" -->
<!--#INCLUDE virtual="/compartido/Errores.asp" -->
<!--#INCLUDE virtual="/cliente/Constantes.asp" -->
<%
	Dim sAFEXchange, sAFEXpress
	Dim sCodigo, nTipoCliente, bExtranjero
	
	On Error	Resume Next
	'Verifica si el usuario puede grabar
	'VerificarBloqueoUsuario	
	
	sAFEXchange = Session("AFEXchange")
		
	GrabarEnvio
	Response.Redirect "Resultado.asp"
	
	Sub GrabarEnvio()
		Dim nEstado
		Dim afxTVouIngreso			
		Dim afxTransfer
		Dim cMontoTransfer
		Dim sMoneda
		Dim cParidad
		Dim cMontoEquivalente
		Dim iValuta
		Dim cComision
		Dim cTarifaSugerida
		Dim cTarifaCobrada
		Dim iTipoParidad
		Dim cNumeroAba	
		Dim iBancoOrigen, sBancoOrigen
		Dim sCuentaOrigen
		Dim sBancoDestino
		Dim sCuentaDestino
		Dim sNombreBeneficiario
		Dim sNombreCiudadDestino
		Dim sDireccionDestino
		Dim nTiempoDespacho
		Dim sBancoInterm
		Dim sCuentaInterm
		Dim sCiudadInterm
		Dim sDireccionInterm
		Dim sMensaje
		Dim sGlosaVoucher, bOk, sStringDetalle
		Dim nOperacion, nFormaPago, cTipoCambio, cMontoNacional
		Dim nDestinoBoleta

		cMontoTransfer = ccur(Cdbl(0 & Request.Form("txtMonto")))
		sMoneda = Request.Form("cbxMoneda")
		cParidad = Cdbl(0 & Request.Form("txtParidad"))
		cMontoEquivalente = ccur(Cdbl(0 & Request.Form("txtEquivalente")))
		iValuta = Request.Form("cbxValuta")	
		cNumeroAba = Request.Form("txtAba")
		iTipoParidad = 1
		iBancoOrigen = 6
		sBancoOrigen = "Bank Of America NT & SA"
		sCuentaOrigen = "6550628510"
		sBancoDestino = Trim(" " & Request.Form("txtBancoDestino"))
		sCuentaDestino = Trim(" " & Request.Form("txtCuentaDestino"))
		sNombreBeneficiario = Trim(" " & Request.Form("txtNombre"))
		sNombreCiudadDestino = Trim(" " & Request.Form("txtCiudad"))
		sDireccionDestino = Trim(" " & Request.Form("txtDireccion"))
		nTiempoDespacho = cInt(0 & Request.Form("cbxValuta"))
		sBancoInterm = Trim(" " & Request.Form("txtBancoIntermedio"))
		sCuentaInterm = Trim(" " & Request.Form("txtCuentaIntermedio"))
		sCiudadInterm = Trim(" " & Request.Form("txtCiudadIntermedio"))
		sDireccionInterm = Trim(" " & Request.Form("txtDireccionIntermedio"))
		cTarifaSugerida = cCur(0 & Request.Form("txtTarifa"))
		cTarifaCobrada = cCur(0 & Request.Form("txtTarifa"))
		cComision = cCur(0 & Request.Form("txtTarifa"))	
		cMontoNacional = cCur(cDbl(0 & Request.Form("txtMontoNacional")))
		cTipoCambio = cCur(cDbl(0 & Request.Form("txtTipoCambio")))
		nFormaPago = cInt(0 & Request.Form("cbxFormaPago"))
		If Request.Form("optEnviarBoleta") = "on" Then
			nDestinoBoleta = 1
		End If
		If Request.Form("optGuardarBoleta") = "on" Then
			nDestinoBoleta = 2
		End If
		Select Case nFormaPago
		Case afxEfectivoUSD
			nOperacion = afxOperacionCanje
			
		Case afxEfectivoCLP
			nOperacion = afxOperacionVenta
		
		Case afxDepositoUSD
			nOperacion = afxOperacionCanje
			
		Case afxDepositoCLP
			nOperacion = afxOperacionVenta
			
		Case afxCustodiaUSD
			nOperacion = afxOperacionCanje

		End Select

		sStringDetalle = iBancoOrigen & ";" & sBancoOrigen & ";" & sCuentaOrigen  & ";;" &  _
								sBancoDestino & ";" & sDireccionDestino & ";;" & _
								sNombreCiudadDestino & ";" &  sCuentaDestino & ";" & _
								sNombreBeneficiario & ";" & cNumeroABA & ";;" & _
								nTiempoDespacho & ";;;;;;;" & sBancoInterm & ";" & _
								sCuentaInterm & ";;" & sDireccionInterm & ";" & _
								sCiudadInterm
		'response.Redirect "../compartido/error.asp?description=" & sAFEXchange & ", " & _
		'	nOperacion & ", " & afxProductoTransferencia & ", " & sMoneda & ", " & cMontoTransfer & ", " & cTipoCambio & ", " & cMontoNacional  & ", " & cTarifaCobrada & ", " & sStringDetalle
		'Transferencia para AFEXchange							  
		Set afxTransfer = Server.CreateObject("AfexWebXP.Web")
		If Err.number <> 0 Then
			Set afxTransfer = Nothing		
			MostrarErrorMS "Grabar Envio Transfer 1"
		End If
		'response.Redirect "../compartido/error.asp?description=" & sStringDetalle
		'response.Redirect "../compartido/error.asp?description=" & Session("afxCnxAFEXchange") & ", " & _
		'						  sAFEXchange & ", " &  Date() & ", " &  Time() & ", " &  nOperacion & ", " &  _
		'						  afxProductoTransferencia & ", " &  sMoneda & ", " &   _
		'						  cMontoTransfer & ", " &  cTipoCambio & ", " &  cMontoNacional & ", " &  _
		'						  cTarifaCobrada & ", " &  1 & ",,, " &  Session("afxCnxCorporativa") & ", " &  Session("CodigoCliente") & ", " & sStringDetalle & ", " &   nDestinoBoleta
								  
		bOk = afxTransfer.AgregarPreSP(Session("afxCnxAFEXchange"), _
								  sAFEXchange, Date(), Time(), nOperacion, _
								  afxProductoTransferencia, sMoneda,  _
								  cMontoTransfer, cTipoCambio, cMontoNacional, _
								  cTarifaCobrada, 1, , , sStringDetalle,True, Session("afxCnxCorporativa"), Session("CodigoCliente"), nDestinoBoleta)
								  
		If Err.number <> 0 Then
			Set afxTransfer = Nothing		
			MostrarErrorMS "Grabar Envio Transfer 2"
		End If
		If afxTransfer.ErrNumber <> 0 Then
			MostrarErrorAFEX afxTransfer, "Grabar Envio Transfer 3"
		End If		
		Set afxTransfer = Nothing			
		If Not bOk Then
			response.Redirect "../compartido/error.asp?Titulo=Grabar Envio Transfer 6&description=No se pudo enviar la transferencia "
		End If
	End Sub	

%>
<!--#INCLUDE virtual="/compartido/Rutinas.asp" -->