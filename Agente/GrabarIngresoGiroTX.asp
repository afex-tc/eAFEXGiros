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
<!--#INCLUDE virtual="/Compartido/Errores.asp" -->
<!--#INCLUDE virtual="/Compartido/Rutinas.asp" -->
<!--#INCLUDE virtual="/Compartido/Giros.asp" -->
<%	
	Dim afxGIMG, afxP, Giro
	Dim sAFEXchange, sAFEXpress
	Dim sCodigo, nTipoCliente, bExtranjero
	Dim invoi
	invoi="TMT"&Trim(Request.Form("txtInvoiceMG"))
	
	On Error	Resume Next
		Dim b
		b = ValidarBDGiros
		If Not b Then Response.Redirect "http:../compartido/informacion.asp?Tipo=1"
	
	sAFEXpress = Request.Form("txtExpress")
	sAFEXchange = Request.Form("txtExchange")
	
	If sAFEXpress = "" Then
		sAFEXpress = AgregarClienteXP
	End If
	'MostrarErrorMS "1"	
	If ValidarInvoice(Session("afxCnxAFEXpress"), invoi, Session("CodigoTXPago")) Then	
		response.Redirect "http:../compartido/informacion.asp?detalle=El Invoice se encuentra duplicado. No se puede agregar el giro."
	End If	
	
	' codigo nuevo para validar credito de agente 22-03-2010 PSS
	If Not ValidarCredito(Session("CodigoAgente"), Session("MonedaExtranjera"), Request.Form("txtMonto.value")) Then
	End If
	
	AgregarGiro
	Set afxGIMG = Nothing
	If Trim(Giro) = "" Then
		MostrarErrorMS "TX 5, Invoice Duplicado"
	Else
		Response.Redirect "AtencionClientes.asp"
	End If


	'Métodos	
	Sub AgregarGiro()
'		Set afxGIMG = Server.CreateObject("AfexWebXP.web")
		Set afxGIMG = Server.CreateObject("AfexGiro.Giro")
		If Err.number <> 0 Then
			Set afxGIMG = Nothing
			MostrarErrorMS "Grabar Ingreso Giro TX "
		End If

		'MostrarErrorMS Session("afxCnxAFEXpress") & ", " & Session("CodigoTXPago") & ", " & _
		'					Session("CodigoAgente") & ", " & cCur(0 & cDbl(Request.Form("txtMonto"))) & ", " & _
		'					cCur(0) & ", " & afxPrioridad.afxGiroUrgente & ", " & afxLugarPago.afxPagoSucursal & ", " & afxFormaPago.afxPagoEfectivo & ", " & Session("MonedaExtranjera") & ", " & _
		'					Session("MonedaExtranjera") & ", " & Request.Form("txtMensajeB") & ", " & _
		'					Request.Form("txtMsjPagador") & ", " & _
		'					Request.Form("txtRut") & ", " & Request.Form("txtPasaporte") & ", " & _
		'					Request.Form("txtPaisPasaporte") & ", " & _
		'					Request.Form("txtNombres") & ", " & Request.Form("txtApellidos") & ", " & _
		'					Trim(Request.Form("txtDireccion")) & ", " & Trim(Request.Form("cbxCiudad")) & ", " & _
		'					Request.Form("cbxComuna") & ", " & Request.Form("cbxPais") & ", " & _
		'					cInt(0 & Request.Form("txtPaisFono")) & ", " & cInt(0 & Request.Form("txtAreaFono")) & ", " & cCur(0 & Request.Form("txtFono")) & ", " & _
		'					", , , " & _
		'					Request.Form("txtNombreR") & ", " & Request.Form("txtApellidoR") & ", " & Request.Form("txtDireccionR") & ", " & _
		'					Request.Form("cbxCiudadR") & ", " & "" & ", " & Request.Form("cbxPaisR") & ", " & _
		'					cInt(0) & ", " & cInt(0) & ", " & cCur(0) & ", " & _
		'					Session("NombreUsuarioOperador") & ", " & sAFEXpress & ", " & "" & ", " & Request.Form("txtInvoiceMG") & ", " & _
		'					cCur(0 & Request.Form("txtBoleta")) & ", " & _
		'					Session("PaisCliente") & ", " & Session("CiudadCliente") & ", " & afxGuardarBoleta & ", " & _
		'					False & ", 0, 0, 0, " & ccur(0 & Request.Form("txtComisionGanada")) & ", " & 0 - ccur(0 & Request.Form("txtComisionGanada")) & ", 0, " & Session("Categoria") & ", " & false
		
		Giro = afxGIMG.IngresarPagoTX(Session("afxCnxAFEXpress"), Session("CodigoTXPago"), _
							Session("CodigoAgente"), cCur(0 & cDbl(Request.Form("txtMonto"))), _
							cCur(0), afxPrioridad.afxGiroUrgente, afxLugarPago.afxPagoSucursal, afxFormaPago.afxPagoEfectivo, Session("MonedaExtranjera"), _
							Session("MonedaExtranjera"), Request.Form("txtMensajeB"), _
							Request.Form("txtMsjPagador"), _
							Request.Form("txtRut"),  Request.Form("txtPasaporte"), _
							Request.Form("txtPaisPasaporte"), _
							Request.Form("txtNombres"), Request.Form("txtApellidos"), _
							Trim(Request.Form("txtDireccion")), Trim(Request.Form("cbxCiudad")), _
							Request.Form("cbxComuna"), Request.Form("cbxPais"),  _
							cInt(0 & Request.Form("txtPaisFono")), cInt(0 & Request.Form("txtAreaFono")), cCur(0 & Request.Form("txtFono")), _
							"", "", "", _
							Request.Form("txtNombreR"), Request.Form("txtApellidoR"), Request.Form("txtDireccionR"), _
							Request.Form("cbxCiudadR"), "", Request.Form("cbxPaisR"), _
							cInt(0), cInt(0), cCur(0), _
							
							Session("NombreUsuarioOperador"), sAFEXpress,  "",invoi,  _
							cCur(0 & Request.Form("txtBoleta")), _
							Session("PaisCliente"), Session("CiudadCliente"), afxGuardarBoleta, _
							True, 0, 0, 0, ccur(0 & Request.Form("txtComisionGanada")), 0 - ccur(0 & Request.Form("txtComisionGanada")), 0,  Session("Categoria"), false)
		
		If Err.number <> 0 Then
			Set afxGIMG = Nothing
			MostrarErrorMS "Grabar Ingreso Giro TX 2"
		End If						
		If afxGIMG.ErrNumber <> 0 Then
			MostrarErrorAFEX afxGIMG, "Grabar Ingreso Giro TX 3"
		End If
				
		If Giro = "" Then
			'afxGIMG.ErrNumber = clng(4)
			MostrarErrorMS "5"
			'afxGIMG.ErrSource = "AFEXGiro.Giro"
			'afxGIMG.ErrDescription = "Error desconocido, revise el numero de invoice"
			'MostrarErrorAFEX afxGIMG, "Grabar Ingreso Giro MG 4"
			'MostrarErrorMS "Grabar Ingreso Giro MG 4"
		End If
		'MostrarErrorMS Giro
		Set afxGIMG = Nothing
	End Sub	
 							 	
	Function EvaluarStr(ByVal Valor)
		Dim Devuelve
		
		If Valor="" Then 
			EvaluarStr = "Null"	
		Else
			EvaluarStr = "'" & Valor & "'"
		End If

	End Function

%>

