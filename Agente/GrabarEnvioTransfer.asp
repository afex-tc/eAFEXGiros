<%@ Transaction=required LANGUAGE = VBScript %>
<%
	'Se asegura que la página no se almacene en la memoria cache
	Response.Expires = 0
	If Session("CodigoCliente") = "" Then
		response.Redirect "http:../Compartido/TimeOut.htm"
		response.end
	End If
%>
<!--#INCLUDE VIRTUAL="/compartido/Errores.asp" -->
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


'	ValidacionTransfer
	GrabarEnvio
	If Err.number <> 0 Then
		MostrarErrorMS "Grabar Envio Transfer 4"
	End If
	' Jonathan Miranda G. 30-03-2007
	'Response.Redirect "AtencionClientes.asp?Accion=" & afxAccionBuscar & "&Campo=" & afxCampoCodigoExchange & "&Argumento=" & sAFEXchange
	Response.Redirect "AtencionClientes.asp"
	'--------------------- Fin --------------------

	Sub ValidacionTransfer
		'mostrarerrorms sAFEXchange & ", " & Trim(" " & Request.Form("txtNombreB")) & ", " &  Request.Form("cbxMoneda") & ", " &  ccur(Cdbl(0 & Request.Form("txtMonto"))))
		If ValidarTransfer(sAFEXchange, Trim(" " & Request.Form("txtNombreB")), Request.Form("cbxMoneda"), ccur(Cdbl(0 & Request.Form("txtMonto")))) Then 
			Err.Raise 2004, "AfexWeb.GrabarTransfer", "Existe una transferencia creada con datos similares a la que ahora desea grabar. Por razones de seguridad no se creará esta transferencia.<br><br>Si necesita mayor información comuníquese con el departamento de transferencias."
			MostrarErrorMS "Grabar Envio de Transfer"
		End If
	End Sub
	
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
		Dim nUso
		Dim sCodigoAgente 
		Dim spago
		Dim sforma1
		Dim sforma
		Dim sformaP

		dim sDireccionBeneficiario
		
		cMontoTransfer = ccur(Cdbl(0 & Request.Form("txtMonto")))

		sMoneda = Request.Form("cbxMoneda")
		
		spago=request.Form("cbxpago")
		sforma=request.Form("formas")
		sforma1=request.Form("formas1")	
		'lugar= request.Form("lugar")
		
		if sforma<>0 then
			sformaP = sforma
		else
			sformaP = Sforma1
		end if	
		
		cParidad = Cdbl(0 & Request.Form("txtParidad"))
		cMontoEquivalente = ccur(Cdbl(0 & Request.Form("txtEquivalente")))
		iValuta = Request.Form("cbxValuta")	
		cNumeroAba = Request.Form("txtAba")
	
		If cNumeroAba = Empty Then
			cNumeroAba = 0
		End If
	
		' verifica el tipo de paridad
		iTipoParidad = cInt(0 & request.Form("cbxRate"))
		iBancoOrigen = 6
		sCuentaOrigen = "6550628510"
		sBancoDestino = Trim(" " & Request.Form("txtBancoDestino"))
		sCuentaDestino = Trim(" " & Request.Form("txtCuentaDestino"))
		sNombreBeneficiario = Trim(" " & Request.Form("txtNombreB"))
		sNombreCiudadDestino = Trim(" " & Request.Form("txtCiudadB"))
		sDireccionDestino = Trim(" " & Request.Form("txtDireccionB"))
		sTiempoDespacho = cInt(0 & Request.Form("cbxValuta"))
		sBancoInterm = Trim(" " & Request.Form("txtBancoIntermedio"))
		sCuentaInterm = Trim(" " & Request.Form("txtCuentaIntermedio"))
		sCiudadInterm = Trim(" " & Request.Form("txtCiudadIntermedio"))
		sDireccionInterm = Trim(" " & Request.Form("txtDireccionIntermedio"))
		sDireccionBeneficiario = Trim(" " & Request.Form("txtDireccionBeneficiario"))
	
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
	
		nEstado = Request.Form("cbxEstado")
			
		' arma la glosa del voucher, dependiendo del tipo de cliente
		If ucase(Request.Form("optPersona")) = "ON" Then			' Persona
			sGlosaVoucher = Trim(Request.Form("txtNombres")) & " "
			sGlosaVoucher = sGlosaVoucher & Trim(Request.Form("txtApellidoP")) & " "
			sGlosaVoucher = sGlosaVoucher & Trim(Request.Form("txtApellidoM"))
		Else		'Empresa
			sGlosaVoucher = Trim(Request.Form("txtRazonSocial"))		
		End If
		sGlosaVoucher = sGlosaVoucher & " (AFEX Money Web)"
		sGlosaVoucher = left(sGlosaVoucher, 50)
	
		'Transferencia para AFEXchange							  
		'Set afxTransfer = Server.CreateObject("AfexWebXP.Web")
		Set afxTransfer = Server.CreateObject("AFEXProducto.Transferencia")
		If Err.number <> 0 Then
			Set afxTransfer = Nothing		
			MostrarErrorMS "Grabar Envio Transfer 1"
		End If		
		
		Dim bOk, nTransfer
		Dim sSucursalXP
		
		sCodigoAgente = Session("CodigoAgente")
		sSucursalXP = Session("CodigoAgente")
		
		Select Case Session("Categoria") 
			Case 1, 2
				If Trim(Session("CodigoAgente")) = "AW" Then
					nUso = 15
					sSucursalXP = ""
				Else
					nUso = 3
				End If
				
			Case 3
				'nUso = 4
				nUso = 2
				If Trim(Session("CodigoAgente")) = "MO" Then
					'sCodigoAgente = "MO"
					sSucursalXP = "MG"
				End If
				
			Case Else
				nUso = 3
		End Select
		
'		response.write Session("afxCnxAFEXchange") & ", " & Session("afxCnxAFEXpress") & ", " & _
'								  iBancoOrigen & ", " & sAFEXchange & ", " & _
'								  cMontoTransfer & ", " & date & ", 1, 0, 8, 0, , " & sCuentaOrigen & ",, " & _
'								  sBancoDestino & ", " & sCuentaDestino & ", " & sNombreBeneficiario & ", " & _
'								  cNumeroAba & ",, " & sNombreCiudadDestino & ", " & sDireccionDestino & ", " & _
'								  sTiempoDespacho & ", " & sMensaje & ", " & _
'								  sMoneda & ", " & "USD" & ", " & cParidad & ", " & _
'								  cMontoEquivalente & ", " & cComision & ", " & sCodigoAgente & ", " & _
'								  Session("NombreUsuario") & ", " & iTipoParidad & ", " & cTarifaSugerida & ", " & _
'								  cTarifaCobrada & ", " & sBancoInterm & ", " & sCuentaInterm & ", " & sCiudadInterm & ", " & _
'								  sDireccionInterm & ", " & sGlosaVoucher & ", " & False & ",,, " & nUso & ", " & sSucursalXP & _
'									", " & Request.Form("txtFurtherCredit") & "**"

'	response.End 


		dim sfecha

		sfecha = formatdatetime(date,1)

		if instr(ucase(trim(sfecha)), "SÁBADO") > 0 then
			sfecha = date + 2
		elseif instr(ucase(trim(sfecha)), "DOMINGO") > 0 then
			sfecha = date + 1
		else
			sfecha = date
		end if		
		
		nTransfer = afxTransfer.Enviar(Session("afxCnxAFEXchange"), Session("afxCnxAFEXpress"), _
								  iBancoOrigen, sAFEXchange, _
								  cMontoTransfer,sfecha, nEstado, 0, 8, 0, , sCuentaOrigen,,  _
								  sBancoDestino, sCuentaDestino,  sNombreBeneficiario, _
								  cNumeroAba,, sNombreCiudadDestino, sDireccionDestino,_
								  sTiempoDespacho, sMensaje, _
								  sMoneda, "USD",spago,SformaP,cParidad, _
								  cMontoEquivalente, cComision, sCodigoAgente, _
								  Session("NombreUsuario"), iTipoParidad, cTarifaSugerida, _
								  cTarifaCobrada, sBancoInterm, sCuentaInterm, sCiudadInterm, _
								  sDireccionInterm, sGlosaVoucher, False,,, nUso, sSucursalXP, Request.Form("txtFurtherCredit"), _
								  Request.form("txtEjecutivoParidad"), Request.Form("txtRefernciaTransfer"), _
								  Request.Form("cbxPaisBeneficiario"), sDireccionBeneficiario)
		
		If Err.number <> 0 Then
			Set afxTransfer = Nothing
			
			'verifica el error para enviar un mail al oficial de cumplimiento
			If InStr(UCase(Err.Description), "NO SE PUEDEN ENVIAR") > 0 Then				
			   ' envia un mail
			   EnviarEMail "AFEX", "arturo.munoz@afex.cl", "", "Transferencia AFEXweb", _
						"El cliente " & Trim(sAFEXchange) & " está intentando enviar una Transferencia a " & _
                     Trim(Request.Form("cbxPaisBeneficiario")) & ". ", 1		   
			   
			End If
			
			MostrarErrorMS "Grabar Envio Transfer 2"
		End If
		If afxTransfer.ErrNumber <> 0 Then
			'verifica el error para enviar un mail al oficial de cumplimiento
			If InStr(UCase(afxTransfer.ErrDescription), "NO SE PUEDEN ENVIAR") > 0 Then				
			   ' envia un mail
			   EnviarEMail "AFEX", "jonathan.miranda@afex.cl", "", "Transferencia AFEXweb", _
						"El cliente " & Trim(sAFEXchange) & " está intentando enviar una Transferencia a " & _
                     Trim(Request.Form("cbxPaisBeneficiario")) & ". ", 1		   
			   
			   IF err.number <> 0 then
					MostrarErrorms "Grabar Envio Transfer 1"
			   end if
			   
			End If
		
			MostrarErrorAFEX afxTransfer, "Grabar Envio Transfer 3"
		End If		
		
		If nTransfer = 0 Then
			Set afxTransfer = Nothing				
			response.Redirect "../compartido/error.asp?Titulo=Grabar Envio Transfer 6&description=Se produjo un error desconocido y no se pudo enviar la transferencia"
		End If
		Set afxTransfer = Nothing			
		'Fin Transferencia para AFEXchange
	End Sub	


%>
<!--#INCLUDE VIRTUAL="/compartido/Rutinas.asp" -->