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
	If ValidarInvoice(Session("afxCnxAFEXpress"), Trim(Request.Form("txtInvoiceMG")), Session("CodigoMGPago")) Then	
		response.Redirect "http:../compartido/informacion.asp?detalle=El Invoice se encuentra duplicado. No se puede agregar el giro."
	End If
	
	AgregarGiro
	Set afxGIMG = Nothing
	If Trim(Giro) = "" Then
		MostrarErrorMS "MG 5, Invoice Duplicado"
	Else
		Response.Redirect "AtencionClientes.asp"
	End If

	


	'Métodos	
	Sub AgregarGiro()
		dim rsGiro
		dim sSQL

'		Set afxGIMG = Server.CreateObject("AfexWebXP.web")
'		Set afxGIMG = Server.CreateObject("AfexGiro.Giro")
'		If Err.number <> 0 Then
'			Set afxGIMG = Nothing
'			MostrarErrorMS "Grabar Ingreso Giro MG "
'		End If

		'MostrarErrorMS Session("afxCnxAFEXpress") & ", " & Session("CodigoMGPago") & ", " & _
		'		Session("CodigoAgente") & ", " & cCur(0 & cDbl(Request.Form("txtMonto"))) & ", " & _
		'		cCur(0) & ", " & afxPrioridad.afxGiroUrgente & ", " & afxLugarPago.afxPagoSucursal & ", " & afxFormaPago.afxPagoEfectivo & ", " & Session("MonedaExtranjera") & ", " & _
		'		Session("MonedaExtranjera") & ", " & Request.Form("txtMensajeB") & ", " & _
		'		Request.Form("txtMsjPagador") & ", " & _
		'		Request.Form("txtRut") & ", " &  Request.Form("txtPasaporte") & ", " & _
		'		Request.Form("txtPaisPasaporte") & ", " & _
		'		Request.Form("txtNombres") & ", " & Request.Form("txtApellidos") & ", " & _
		'		Trim(Request.Form("txtDireccion")) & ", " & Trim(Request.Form("cbxCiudad")) & ", " & _
		'		Request.Form("cbxComuna") & ", " & Request.Form("cbxPais") & ", " &  _
		'		cInt(0 & Request.Form("txtPaisFono")) & ", " & cInt(0 & Request.Form("txtAreaFono")) & ", " & cCur(0 & Request.Form("txtFono")) & ", " & _
		'		"" & ", " & "" & ", " & "" & ", " & _
		'		Request.Form("txtNombreR") & ", " & Request.Form("txtApellidoR") & ", " & Request.Form("txtDireccionR") & ", " & _
		'		Request.Form("cbxCiudadR") & ", " & "" & ", " & Request.Form("cbxPaisR") & ", " & _
		'		cInt(0) & ", " & cInt(0) & ", " & cCur(0) & ", " & _
		'		Session("NombreUsuarioOperador") & ", " & sAFEXpress & ", " &  "" & ", " & Request.Form("txtInvoiceMG") & ", " &  _
		'		cCur(0 & Request.Form("txtBoleta")) & ", " & _
		'		Session("PaisCliente") & ", " & Session("CiudadCliente") & ", " & afxGuardarBoleta & ", " & _
		'		False & ", " & 0 & ", " & 0 & ", " & 0 & ", " & ccur(0 & Request.Form("txtComisionGanada")) & ", " & 0 - ccur(0 & Request.Form("txtComisionGanada")) & ", " & 0 & ", " &  Session("Categoria") & ", " & false


		'Giro = afxGIMG.IngresarPagoMG(Session("afxCnxAFEXpress"), Session("CodigoMGPago"), _
		'					Session("CodigoAgente"), cCur(0 & cDbl(Request.Form("txtMonto"))), _
		'					cCur(0), afxPrioridad.afxGiroUrgente, afxLugarPago.afxPagoSucursal, 'afxFormaPago.afxPagoEfectivo, Session("MonedaExtranjera"), _
		'					Session("MonedaExtranjera"), Request.Form("txtMensajeB"), _
		'					Request.Form("txtMsjPagador"), _
		'					Request.Form("txtRut"),  Request.Form("txtPasaporte"), _
		'					Request.Form("txtPaisPasaporte"), _
		'					Request.Form("txtNombres"), Request.Form("txtApellidos"), _
		'					Trim(Request.Form("txtDireccion")), Trim(Request.Form("cbxCiudad")), _
		'					Request.Form("cbxComuna"), Request.Form("cbxPais"),  _
		'					cInt(0 & Request.Form("txtPaisFono")), cInt(0 & 'Request.Form("txtAreaFono")), cCur(0 & Request.Form("txtFono")), _
		'					"", "", "", _
		'					Request.Form("txtNombreR"), Request.Form("txtApellidoR"), 'Request.Form("txtDireccionR"), _
		'					Request.Form("cbxCiudadR"), "", Request.Form("cbxPaisR"), _
		'					cInt(0), cInt(0), cCur(0), _
		'					Session("NombreUsuarioOperador"), sAFEXpress,  "", 'Request.Form("txtInvoiceMG"),  _
		'					cCur(0 & Request.Form("txtBoleta")), _
		'					Session("PaisCliente"), Session("CiudadCliente"), afxGuardarBoleta, _
		'					False, 0, 0, 0, ccur(0 & Request.Form("txtComisionGanada")), 0 - ccur(0 & 'Request.Form("txtComisionGanada")), 0,  Session("Categoria"), false)
		
		' Jonathan Miranda G. 03-01-2007
		If trim(Request.Form("cbxMonedaPago"))=session("MonedaNacional")  Then
			sSQL="AgregarPagoMGPesos " 
		else
			sSQL="AgregarPagoMG "
		End If
		
		
		sSQL = sSQL & EvaluarSTR(Request.Form("txtInvoiceMG")) & ", " & _
		EvaluarSTR(Session("CodigoAgente")) & ", " & _
		FormatoNumeroSQL(cCur(0 & cDbl(Request.Form("txtMonto")))) & ", " & _
		formatonumerosql(ccur(0 & Request.Form("txtComisionGanada"))) & ", " & _
		formatonumerosql(ccur(0 & Request.Form("txtComisionGanada"))) & ", " & _
		EvaluarSTR(sAFEXpress) & ", " & EvaluarSTR(Request.Form("txtRut")) & ", " & _
		EvaluarSTR(Request.Form("txtPasaporte")) & ", " & _
		EvaluarSTR(Request.Form("cbxPaisPasaporte")) & ", " & _
		EvaluarSTR(Request.Form("txtNombres")) & ",	" & _
		EvaluarSTR(Request.Form("txtApellidos")) & ", " & _
		EvaluarSTR(Trim(Request.Form("txtDireccion"))) & ", " & _
		EvaluarSTR(Request.Form("cbxComuna")) & ", " & _
		EvaluarSTR(Trim(Request.Form("cbxCiudad"))) & ", " & _
		EvaluarSTR(Request.Form("cbxPais")) & ", " & _
		formatonumerosql(cInt(0 & Request.Form("txtPaisFono"))) & ", " & _
		formatonumerosql(cInt(0 & Request.Form("txtAreaFono"))) & ", " & _
		formatonumerosql(cCur(0 & Request.Form("txtFono"))) & ", " & _
		evaluarstr(Request.Form("txtNombreR")) & ", " & _
		evaluarstr(Request.Form("txtApellidoR")) & ", " & evaluarstr(Request.Form("cbxPaisR")) & _
		", " & evaluarstr(Session("NombreUsuarioOperador")) & ", " & _
		evaluarstr(Session("CodigoMGPago"))

		' JFMG 13-05-2009 este objeto tenía otro nombre en la página anterior
		'EvaluarSTR(Request.Form("txtPaisPasaporte")) & ", " & _
		' ******************** FIN ********************* 


		'response.write ssql
		'response.end

		set rsGiro = EjecutarSqlCliente(Session("afxCnxAFEXpress"), sSQL)		
		'------------------------- Fin -----------------------------------

		If Err.number <> 0 Then
'			Set afxGIMG = Nothing
			MostrarErrorMS "Grabar Ingreso Giro MG 2"
		End If						
'		If afxGIMG.ErrNumber <> 0 Then
'			MostrarErrorAFEX afxGIMG, "Grabar Ingreso Giro MG 3"
'		End If
		Giro = rsGiro("codigogiro")
		If Giro = "" Then
			'afxGIMG.ErrNumber = clng(4)
			MostrarErrorMS "5"
			'afxGIMG.ErrSource = "AFEXGiro.Giro"
			'afxGIMG.ErrDescription = "Error desconocido, revise el numero de invoice"
			'MostrarErrorAFEX afxGIMG, "Grabar Ingreso Giro MG 4"
			'MostrarErrorMS "Grabar Ingreso Giro MG 4"
		End If
		'MostrarErrorMS Giro
'		Set afxGIMG = Nothing
		rsGiro.close
		set rsGiro = nothing
	End Sub	
 							 	
%>
