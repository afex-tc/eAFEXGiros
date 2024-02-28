<%@ Language=VBScript %>
<%
	'Se asegura que la página no se almacene en la memoria cache
	Response.Expires = 0
'	If Session("CodigoCliente") = "" Then
'		response.Redirect "../Compartido/TimeOut.htm"
'		response.end
'	End If
%>
<!--#INCLUDE virtual="/Agente/Constantes.asp" -->
<!--#INCLUDE virtual="/Sucursal/Rutinas.asp" -->
<%

	dim rs, sSQL, cTipoCambio, rsTC

	Dim sMoneda, sVigencia, cObservado, sDireccionBanco, sDireccionBancoInt
	Dim nMonto, nParidad, nEquivalente, nTarifaSugerida, nTarifaCobrada
	Dim nGastoExtranjera, nGastoNacional, nTotal, nEstadoTRF
	
'	On Error Resume Next
	
	sMoneda = Trim(request.Form("cbxMoneda"))
	If sMoneda = "" Then
		sMoneda = Session("MonedaExtranjera")
	End If

	nEstadoTRF = 1
	
	sVigencia = ValidarVigencia

	nAccion = cInt(0 & Request("Accion"))
	CargarTrf
	
	Sub CargarTrf()
		cObservado = ObtenerObservado()		
		nMonto = cCur(cDbl(0 & Request.Form("txtMonto")))
	End Sub
	
	sBanco = request.Form("cbxBancoDestino")
	sCiudadBanco = request.Form("cbxCiudadBancoDestino")
	sBancoInt = request.Form("cbxBancoDestinoInt")
	sCiudadBancoInt = request.Form("cbxCiudadBancoDestinoInt")	
	
	Sub CargarBanco(ByVal Banco)
		Dim sSQL
		Dim rs
		
		On error resume next
		' Jonathan Miranda G. 29-08-2007 se cambia para que cargue todos los bancos
		'sSQL = " SELECT distinct b.codigo_banco AS Codigo, b.nombre_banco AS Nombre " & _
		'		 "	FROM    banco b " & _					
		'			" inner join transferencia t on b.codigo_banco = t.codigo_banco_destino " & _
		'			" inner join cliente c on c.codigo_cliente = t.codigo_cliente " & _					
		'		 " WHERE c.codigo_corporativa = " & session("codigocliente") & _
		'		 " ORDER BY b.nombre_banco "
		sSQL = " SELECT distinct b.codigo_banco AS Codigo, b.nombre_banco AS Nombre " & _
				 "	FROM    banco b " & _					
				 " ORDER BY b.nombre_banco "
		'******************************************* Fin **************************************
		
		set rs = EjecutarSQLCliente(Session("afxCnxAFEXchange"), sSQL)
		if err.number <> 0 then
			response.Write "Error. " & err.Description
			response.End
			exit sub
		end if
		
		response.Write "<option value=></option>"
		do until rs.eof			
			If Trim(Banco) = Trim(rs("codigo")) Then
				response.Write "<option selected value=" & rs("codigo") & "> " & rs("nombre") & "</option> "				
			else
				response.Write "<option value=" & rs("codigo") & "> " & rs("nombre") & "</option> "
			End if
			rs.movenext
		loop
	
		set rs = nothing
	End Sub
		
	Sub CargarCiudadBanco(ByVal Banco, Byval Ciudad)
		Dim sSQL
		Dim rs
		
		If Banco = "" Then exit sub
		
		On error resume next
		sSQL = " SELECT c.codigo_ciudad AS Codigo, c.nombre_ciudad AS Nombre " & _
				"FROM   banco b " & _
					"inner join pais p on (p.nombre_pais = b.pais_banco or p.codigo_pais = b.pais_banco) " & _
					"inner join ciudad c on p.codigo_pais = c.codigo_pais " & _
				"WHERE  b.codigo_banco = " & Banco & _
				"ORDER BY c.nombre_ciudad "		
		set rs = EjecutarSQLCliente(Session("afxCnxAFEXchange"), sSQL)
		if err.number <> 0 then
			response.Write "Error. " & err.Description 
			response.End
			exit sub
		end if
		
		response.Write "<option value=></option>"
		do until rs.eof			
			if trim(Ciudad) = trim(rs("codigo")) Then
				response.Write "<option selected value=" & rs("codigo") & "> " & rs("nombre") & "</option> "
			else
				response.Write "<option value=" & rs("codigo") & "> " & rs("nombre") & "</option> "
			End if
			rs.movenext
		loop
	
		set rs = nothing
	End Sub
	
	' verifica si se debe mostrar el detalle de una transfer
	if request("Transferencia") <> "" then
		sSQL = "select * " & _
				"from  transferencia " & _				
				"where  correlativo_transferencia = " & request("Transferencia")
		set rs = ejecutarsqlcliente(session("afxCnxAFEXchange"), sSQL)
		if err.number <> 0 then			
			response.Write "Ocurrió un error al buscar la transferencia." & err.Description
			response.End
		end if
		if rs.eof then
			response.Write "No se encontró la transferencia." & err.Description
			response.End
		end if
		
		sBanco = rs("codigo_banco_destino")
		sCiudadBanco = rs("ciudad_destino")
		sBancoInt = rs("Banco_Intermediario")
		sCiudadBancoInt = rs("nombre_Ciudad_Intermediario")
		
	'	Response.Write sBanco & "//" & sCiudadBanco
	'	Response.End		
	end if

	' verifica si hay un tipo de cambio para el cliente que envía
	sSQL = "select d.tipo_cambio as tipocambio " & _
			"from  detalle_solicitud d " & _
				" inner join solicitud s on s.codigo_solicitud = d.codigo_solicitud " & _
				" inner join cliente c on c.codigo_cliente = s.codigo_cliente and " & _
									 " c.codigo_corporativa = " & Session("CodigoCliente") & _
			"where  s.fecha_solicitud = " & evaluarstr(date) & _
				 " and s.estado_solicitud = 1 " & _
				 " and d.numero_linea = 1 "
	set rsTC = ejecutarsqlcliente(session("afxCnxAFEXchange"), sSQL)
	if err.number <> 0 then			
		response.Write "Ocurrió un error al buscar el tipo de cambio para el cliente." & err.Description
		response.End
	end if
	if not rsTC.eof then
		cTipoCambio = rsTC("tipocambio")
	end if
	set rsTC = nothing
	
	
	' saca la direccion del banco
	if Request.Form("cbxBancoDestino") <> "" then
		sSQL = " SELECT b.direccion_banco as direccion " & _
				 "	FROM    banco b " & _
				 " WHERE  b.codigo_banco = " & Request.Form("cbxBancoDestino")
				 
		'		 Response.Write ssql
	'			 Response.End 
				 
		set rs = EjecutarSQLCliente(Session("afxCnxAFEXchange"), sSQL)
		if err.number <> 0 then
			response.Write "Ocurrió un error al buscar la dirección del Banco seleccionado. " & err.Description 
			response.End
					
		elseif not rs.eof then
			sDireccionBanco = rs("direccion")	
		end if
		set rs = nothing
	end if
	if Request.Form("cbxBancoDestinoInt") <> "" then
		sSQL = " SELECT b.direccion_banco as direccion " & _
				 "	FROM    banco b " & _
				 " WHERE  b.codigo_banco = " & Request.Form("cbxBancoDestinoInt")
				 
		'		 Response.Write ssql
	'			 Response.End 
				 
		set rs = EjecutarSQLCliente(Session("afxCnxAFEXchange"), sSQL)
		if err.number <> 0 then
			response.Write "Ocurrió un error al buscar la dirección del Banco Intermediario seleccionado. " & err.Description 
			response.End
					
		elseif not rs.eof then
			sDireccionBancoInt = rs("direccion")
		end if
		set rs = nothing
	end if		
%>
<!--#INCLUDE virtual="/Compartido/Transferencias.asp" -->
<!--#INCLUDE virtual="/Compartido/Rutinas.asp" -->
<!--#INCLUDE virtual="/Compartido/Errores.asp" -->
<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link rel="stylesheet" type="text/css" href="../Estilos/Principal.css">
<link rel="stylesheet" type="text/css" href="../Estilos/Cliente.css">
<link rel="stylesheet" type="text/css" href="../Estilos/Reportes.css">
<link href="estilosucursalvirtual.css" rel="stylesheet" type="text/css">
</head>
<script LANGUAGE="VBScript">
<!--
	'Variables de módulo
	'Variables para encabezado
	Dim sEncabezadoFondo, sEncabezadoTitulo
	
	sEncabezadoFondo = "Transacciones"
	sEncabezadoTitulo = "Enviar una Transferencia"

	Sub txtMonto_OnBlur()
		If Trim(frmTransfer.txtMonto.value) = "" Then
			frmTransfer.txtMonto.value = "0,00"
		Else			
			frmTransfer.txtMonto.value = FormatNumber(frmTransfer.txtMonto.value, 2)
		End If		
		'cbxMoneda_onClick
		CalcularUSD
		CalcularMontos
	End Sub 

	Sub txtParidad_onBlur()
		If Trim(frmTransfer.txtParidad.value) = "" Then
			frmTransfer.txtParidad.value = "0,00"
		Else
			frmTransfer.txtParidad.value = FormatNumber(frmTransfer.txtParidad.value, 8)
		End If		
		CalcularUSD
		CalcularMontos
	End Sub

	Sub cbxValuta_onKeyPress()
		cbxValuta_onClick		
	End Sub
	
	Sub cbxValuta_onKeyDown()
		cbxValuta_onClick
	End Sub

	Sub cbxValuta_onKeyUp()
		cbxValuta_onClick
	End Sub

	Sub cbxValuta_onClick()
		CalcularUSD
		CalcularMontos 
	End Sub

	Sub cbxValuta_onBlur()
		cbxValuta_onClick
	End Sub

	Sub imgCalcular_onClick()
		CalcularUSD
		CalcularMontos
	End Sub
	
	Sub trDatosTransfer1_onMouseOver()
		trDatosTransfer1.style.cursor = "Hand"
	End Sub
 
	Sub trDatosTransfer1_onClick()
		If trDatosTransfer2.style.display = "" Then			
			trDatosTransfer2.style.display = "none" 
			dspDatosTransfer.innerText = "u"
		Else
			trDatosTransfer2.style.display = "" 
			dspDatosTransfer.innerText = "t"		
		End If
	End Sub
	
	Sub imgAceptar_onCLick()
		dim s
		
	'	set s = createobject("Registry Access Tools")
	'	msgbox s.RegistryGetValue(regHKEY_CURRENT_USER, _
    '                        "software\microsoft\internet account manager\accounts\00000001\smtp email address")
    
		
		If Not CajaPregunta("AFEX En Linea", "Está seguro que desea enviar la transferencia?") Then
			Exit Sub
		End If

		If Not ValidarDatos() Then
			Exit Sub
		End If		
						
		HabilitarControles		
		frmTransfer.action = "EnviarMail.asp?Accion=6&Monto=" & frmTransfer.txtMonto.value & "&Moneda=" & frmTransfer.cbxMoneda.options(frmTransfer.cbxMoneda.selectedIndex).text & _
														"&Valuta=" & frmTransfer.cbxValuta.options(frmTransfer.cbxValuta.selectedIndex).text & _
														"&Rate=" &  frmTransfer.txtParidad.value & "&Equivalente=" & frmTransfer.txtEquivalente.value & _
														"&Tarifa=0&Total=" & frmTransfer.txtTotal.value & _
														"&CodigoMoneda=" & frmTransfer.cbxMoneda.value
														    '"GrabarEnvioTransfer.asp"
		frmTransfer.submit 
		frmTransfer.action = ""			
	End Sub

	'Objetivo:	calcular los montos de la página
	Sub CalcularMontos()
	
		
		' calcula la tarifa
		frmTransfer.txtTarifaSugerida.value = 0 'ObtenerTarifaTransfer(frmTransfer.cbxValuta.value, _
												'		frmTransfer.cbxMoneda.value, frmTransfer.txtMonto.value)		
		' calcula el total
		frmTransfer.txtTotal.value = CalcularTotal
		
		if frmTransfer.optPesos.checked then
			if frmTransfer.txtMontoPesos.value <> Empty then
				frmTransfer.txtMontoPesos.value = formatnumber(round(ccur(frmTransfer.txtTotal.value) * ccur(frmTransfer.txtTipoCambio.value), 0), 0)
			end if
		end if
	End Sub	

	Sub cbxMoneda_onKeyPress()
		cbxMoneda_onClick		
	End Sub
	
	Sub cbxMoneda_onKeyDown()
		cbxMoneda_onClick
	End Sub

	Sub cbxMoneda_onKeyUp()
		cbxMoneda_onClick
	End Sub

	Sub cbxMoneda_onClick()
		frmTransfer.txtParidad.value  = FormatNumber( _											
							frmTransfer.cbxParidades(frmTransfer.cbxMoneda.selectedIndex).text _
							, 8)		
		CalcularUSD
		
		' calcula los montos
		CalcularMontos
	End Sub
	
	Sub cbxMoneda_onChange()
		If frmTransfer.cbxMoneda.value <> "USD" Then
			frmTransfer.optPesos.checked = true
			optPesos_onClick()
			frmTransfer.optDolares.disabled = true
			frmTransfer.optPesos.disabled = true
		else
			frmTransfer.optDolares.disabled = false
			frmTransfer.optPesos.disabled = false
		End if
	End Sub
	
	Sub cbxRate_onKeyDown()
		cbxRate_onClick
	End Sub

	Sub cbxRate_onKeyUp()
		cbxRate_onClick
	End Sub

	Sub cbxRate_onBlur()
		cbxRate_onClick
	End Sub
	
	Sub cbxRate_onClick()
		If frmTransfer.cbxRate.value = "1" Then
			frmTransfer.txtParidad.value  = FormatNumber( _											
							frmTransfer.cbxParidades(frmTransfer.cbxMoneda.selectedIndex).text _
							, 8)
			frmTransfer.txtParidad.disabled = True
		Else
			frmTransfer.txtParidad.disabled = False
		End If
		CalcularUSD
		CalcularMontos		
	End Sub
	
	Sub CalcularUSD
		'Calcula el Equivalente en Dolares
		frmTransfer.txtEquivalente.value = FormatNumber(Round(CDbl(0 & frmTransfer.txtMonto.value) * CDbl(0 & frmTransfer.txtParidad.value), 7), 2)
	'	if frmTransfer.txtEquivalente.value => 5000 then
	'		frmTransfer.txtMonto.value = "0,00"
	'		frmTransfer.txtEquivalente.value = "0,00"
	'		msgbox "Para enviar una Transferencia superior o igual a US$ 5.000, debe hacerlo desde alguna sucursal " & vbCrlf & _
	'				"o llamar al 6369022. Gracias.",,"Enviar Transferencia"
	'	end if
	End Sub

	'Objetivo:	calcular el total de la transferencia
	Function CalcularTotal()
		CalcularTotal = 0
	
		' verifica si hay monto
		If frmTransfer.txtMonto.value = Empty Or frmTransfer.txtMonto.value = "0,00" Then Exit Function
		
		CalcularTotal = FormatNumber(cDbl(0 & frmTransfer.txtTarifaSugerida.value) + _
									cDbl(0 & frmTransfer.txtEquivalente.value), 2)
	End Function

	Sub UltimasTransferencias()
		Dim sString, aTrf, sNombre, sCliente
	
		If Trim(frmTransfer.txtExchange.value) = "" Then
			sCliente = Trim(frmTransfer.txtExpress.value)
		Else
			sCliente = Trim(frmTransfer.txtExchange.value)
		End If
		
		If frmTransfer.optPersona.value = "on" Then
			sNombre = Trim(Trim(frmTransfer.txtnombres.value) & " " & Trim(frmTransfer.txtApellidoP.value) & " " & Trim(frmTransfer.txtApellidoM.value))
		Else
			sNombre = Trim(Trim(frmTransfer.txtRazonSocial.value))
		End If
		
		sString = Empty
		sString = window.showModalDialog("../Compartido/UltimasTransferencias.asp?CodigoCliente=" & sCliente & _
																	"&NombreCliente=" & sNombre & _
																	"&CodigoMoneda=" & frmTransfer.cbxMoneda.value)
		If sString <> Empty Then
			' crea un arreglo con los datos de la transferencia seleccionada
			aTrf = Split(sString, ";", 14)
			
			' asigna los datos de la transferencia
			frmTransfer.txtBancoOrigen.value = aTrf(0)
			frmTransfer.txtCuentaOrigen.value = aTrf(1)
			frmTransfer.txtBancoDestino.value = aTrf(2)			
			frmTransfer.txtCuentaDestino.value = aTrf(3)
			frmTransfer.txtNombreB.value = aTrf(4)
			frmTransfer.txtCiudadB.value = aTrf(5)
			frmTransfer.txtABA.value = aTrf(6)
			'frmTransfer.txtInvoice.value = aTrf(7)
			frmTransfer.txtDireccionB.value = aTrf(8)
			frmTransfer.txtBancoIntermedio.value = aTrf(9)
			frmTransfer.txtCuentaIntermedio.value = aTrf(10)
			frmTransfer.txtCiudadIntermedio.value = aTrf(11)			
			frmTransfer.txtDireccionIntermedio.value = aTrf(12)
			frmTransfer.txtFurtherCredit.value = aTrf(13)
			<% bTransferencia = True %>					
		End If		
	End Sub

	Function ValidarDatos()
		ValidarDatos = False
		If cCur(0 & frmTransfer.txtMonto.value) <= 0 then
			MsgBox "El Monto debe mayor a 0.",,"AFEX"
			frmTransfer.txtMonto.focus
			exit function
		end if
		if frmTransfer.txtBanco.value = "" then
			MsgBox "Debe ingresar el Banco.",,"AFEX"
			frmTransfer.txtBanco.focus
			exit function
		end if
		if frmTransfer.txtCtaCte.value = "" then
			MsgBox "Debe ingresar la Cta.Cte.",,"AFEX"
			frmTransfer.txtCtaCte.focus
			exit function
		end if
		if frmTransfer.txtNombreExacto.value = "" then
			MsgBox "Debe ingresar el Nombre de la Cta.Cte.",,"AFEX"
			frmTransfer.txtNombreExacto.focus
			exit function
		end if		
		if frmTransfer.txtdireccionb.value = "" then
			MsgBox "Debe ingresar la Dirección del Beneficiario.",,"AFEX"
			frmTransfer.txtdireccionb.focus
			exit function
		end if
		if frmTransfer.txtABA.value = "" and frmTransfer.txtSWIFT.value = "" Then
			MsgBox "Debe ingresar ABA o SWIFT.",,"AFEX"
			frmTransfer.txtABA.focus
			Exit Function
		End If
		
		if frmTransfer.chkBIntermedio.checked then
			If frmTransfer.txtBancoInt.value = "" then
				MsgBox "Debe ingresar el Banco Intermediario.",,"AFEX"
				frmTransfer.txtBancoInt.focus
				Exit Function
			End If
			if  frmTransfer.txtCtaCteInt.value = "" Then
				MsgBox "Debe ingresar la Cta.Cte. del Banco Intermediario.",,"AFEX"
				frmTransfer.txtCtaCteInt.focus
				Exit Function
			End If
		end if
		
		if trim(frmTransfer.txtABA.value) <> "" then frmTransfer.txtABA.value = "FW-" & trim(frmTransfer.txtABA.value)
		if trim(frmTransfer.txtSWIFT.value) <> "" Then frmTransfer.txtABA.value = trim(frmTransfer.txtABA.value) & "SA-" & trim(frmTransfer.txtSWIFT.value)
		if trim(frmTransfer.txtCHIPS.value) <> "" then frmTransfer.txtDireccionBanco.value = frmTransfer.txtDireccionBanco.value & _
												"   CH-" & trim(frmTransfer.txtCHIPS.value)
		if trim(frmTransfer.txtIBAN.value) <> "" then frmTransfer.txtCtaCte.value = trim(frmTransfer.txtIBAN.value)				
		
		ValidarDatos = True
	End Function

	sub window_onload()
		
		cbxMoneda_onClick
		frmTransfer.txtmonto.select
		frmTransfer.txtmonto.focus
				
		<%If Request.form("optDolares") <> "on" Then %>
			frmTransfer.optPesos.checked = true
			optPesos_onClick()
		<%Else%>
			frmTransfer.optDolares.checked = true
		<%End If%>		
		
		'verifica si hay que mostrar una transfer
		<%if request("Transferencia") <> "" then%>
			frmTransfer.txtmonto.value = formatnumber("<%=rs("monto_transferencia")%>",2)
			frmTransfer.cbxmoneda.value = "<%=rs("codigo_moneda")%>"
			frmTransfer.txtparidad.value = "<%=formatnumber(rs("paridad"),8)%>"
			frmTransfer.txtequivalente.value = "<%=formatnumber(rs("monto_equivalente"),2)%>"
			frmTransfer.txttarifasugerida.value = "<%=formatnumber("0" & rs("tarifa_sugerida"),2)%>"
			frmTransfer.cbxbancodestino.value = "<%=rs("nombre_banco_destino")%>" '"<%=rs("codigo_banco_destino")%>"
			frmTransfer.txtbanco.value = "<%=rs("nombre_banco_destino")%>"
			frmTransfer.txtctacte.value = "<%=rs("cuenta_corriente_destino")%>"
			frmTransfer.txtnombreexacto.value = "<%=rs("nombre_titular_destino")%>"
			frmTransfer.txtdireccionbanco.value = "<%=rs("direccion_destino")%>"
			frmTransfer.txtciudadbanco.value = "<%=rs("nombre_ciudad_destino")%>"
			frmTransfer.txtdireccionb.value = "<%=rs("direccionbeneficiario")%>"
			frmTransfer.txtinvoice.value = "<%=rs("referenciatransfer")%>"
			frmTransfer.txtaba.value = "<%=rs("numero_aba")%>"
			'frmTransfer.txtchips.value = "<%=rs("ciudad_destino")%>"
			'frmTransfer.txtswift.value = "<%=rs("ciudad_destino")%>"
			'frmTransfer.txtiban.value = "<%=rs("ciudad_destino")%>"
			'frmTransfer.txtotro.value = "<%=rs("ciudad_destino")%>"
			frmTransfer.cbxbancodestinoint.value = "<%=rs("banco_intermediario")%>"
			frmTransfer.txtbancoint.value = "<%=rs("banco_intermediario")%>"
			frmTransfer.txtctacteint.value = "<%=rs("cuenta_intermediario")%>"
			'frmTransfer.txtnombreexactoint.value = "<%=rs("nombre_titular_destino")%>"
			frmTransfer.txtdireccionbancoint.value = "<%=rs("direccion_intermediario")%>"
			frmTransfer.txtciudadbancoint.value = "<%=rs("nombre_ciudad_intermediario")%>"
			'frmTransfer.txtabaint.value = "<%=rs("numero_aba")%>"
			
		<%else%>
			if "<%=sDireccionBanco%>" <> "" then
				frmTransfer.txtDireccionBanco.value = "<%=sDireccionBanco%>"
			end if
			if "<%=sDireccionBancoInt%>" <> "" then
				frmTransfer.txtDireccionBancoInt.value = "<%=sDireccionBancoInt%>"
			end if
			
			if "<%=request("Mensaje")%>" <> "" then
				msgbox "<%=request("Mensaje")%>",,"AFEX"
			end if
			
			frmTransfer.txtmonto.focus
			frmTransfer.txtmonto.select()
		<%end if%>
		
		if frmTransfer.cbxbancodestinoint.value <> "" or frmTransfer.chkBIntermedio.checked then
			frmTransfer.chkBIntermedio.checked = true
			chkBIntermedio_onClick()
		end if
		
	end sub
	
	Sub optPesos_onClick()
		frmTransfer.optDolares.checked = false
		
		if trim("<%=cTipoCambio%>") = "" then
			sValor = window.showModalDialog("Tarifas.asp?Tipo=2")
		else
			sValor = trim("<%=cTipoCambio%>")
		end if
		if Len(sValor) > 10 Then
			msgbox sValor,,"EnviarTransferencia"
		
		else
			frmTransfer.txtTipoCambio.value = formatnumber(round(sValor, 4), 4)
			frmTransfer.txtMontoPesos.value = formatnumber(round(ccur(frmTransfer.txtTipoCambio.value) * ccur(frmTransfer.txtTotal.value), 0), 0)
		end if		
	End Sub
	Sub optDolares_onClick()
		frmTransfer.optPesos.checked = false
		frmTransfer.txtTipoCambio.value = Empty
		frmTransfer.txtMontoPesos.value = Empty
	End Sub
	
	Sub chkBIntermedio_onClick()
		if frmTransfer.chkBIntermedio.checked then
			window.tlbBIntermedio.style.display = ""
		else
			window.tlbBIntermedio.style.display = "none"
		end if
		
		frmTransfer.chkBIntermedio.value = frmTransfer.chkBIntermedio.checked
	End Sub
	
	Sub cbxBancoDestino_onChange()
		frmTransfer.txtBanco.value = frmTransfer.cbxBancoDestino.options(frmTransfer.cbxBancoDestino.selectedIndex).text
		frmTransfer.cbxBancoDestino.style.display = "none"

'		frmTransfer.action = "EnviarTransfer.asp"
'		frmTransfer.submit()
'		frmTransfer.action = ""
	End Sub
	
	Sub cbxBancoDestinoInt_onChange()
		frmTransfer.txtBancoInt.value = frmTransfer.cbxBancoDestinoInt.options(frmTransfer.cbxBancoDestinoInt.selectedIndex).text
		frmTransfer.cbxBancoDestinoInt.style.display = "none"
		
'		frmTransfer.action = "EnviarTransfer.asp"
'		frmTransfer.submit()
'		frmTransfer.action = ""
	End Sub
	
'	Sub cbxCiudadBancoDestino_onChange()
'		frmTransfer.txtCiudadBanco.value = frmTransfer.cbxCiudadBancoDestino.options(frmTransfer.cbxCiudadBancoDestino.selectedIndex).text		
'	End Sub
	
'	Sub cbxCiudadBancoDestinoInt_onChange()
'		frmTransfer.txtCiudadBancoInt.value = frmTransfer.cbxCiudadBancoDestinoInt.options(frmTransfer.cbxCiudadBancoDestinoInt.selectedIndex).text		
'	End Sub

	sub MostrarBanco1()
		if frmTransfer.cbxBancoDestino.style.display = "" then
			frmTransfer.cbxBancoDestino.style.display = "none"
		else
			frmTransfer.cbxBancoDestino.style.display = ""
		end if
	end sub
	sub MostrarBanco2()
		if frmTransfer.cbxBancoDestinoInt.style.display = "" then
			frmTransfer.cbxBancoDestinoInt.style.display = "none"
		else
			frmTransfer.cbxBancoDestinoInt.style.display = ""
		end if
	end sub

	sub txtBanco_onblur()
		frmTransfer.cbxBancoDestino.value = frmTransfer.txtBanco.value
	end sub
	
	sub txtBancoInt_onblur()
		frmTransfer.cbxBancoDestinoInt.value = frmTransfer.txtBancoInt.value
	end sub

-->
</script>

<body>

<form id="frmTransfer" method="post">
<!-- Paso 1 -->
<table WIDTH="540px" HEIGHT="100" BORDER="0" CELLPADDING="0" CELLSPACING="0" class="cajas" ID="tabPaso1" STYLE="HEIGHT: 100px; LEFT: 1px; POSITION: relative; TOP: 0px; " name="tabPaso1">
	<tr>
	  <td colspan="2" bgcolor="#31514A" align="center"><div align="left"><img src="../Img/enviartransferencia.jpg" width="215" height="16"></div></td>
    </tr>
	<tr>
		<td colspan="2" align="center"><font FACE="tahoma" SIZE="2"><%=sVigencia%></font></td>
	</tr>		
	<tr>
	  <td class="titulos">Valores</td>
	</tr>		
	<tr><td><table border="0" cellpadding=0 cellspacing="2" id="tabTransfer">      
      <tr height="15">
        <td></td>
        <td><span class="textoempresa">Monto</span><br>
          <input name="txtMonto" class="Borde_tabla_abajo" style="HEIGHT: 18px; TEXT-ALIGN: right; WIDTH: 90px" onKeyPress="IngresarTexto(1)" value="<%=FormatNumber(nMonto, 2)%>" size="11"></td>
        <td><span class="textoempresa">Monedas de Pago</span><br>
            <select name="cbxMoneda" class="Borde_tabla_abajo" style="HEIGHT: 22px; WIDTH: 145px">
              <%
						CargarMonedasTransfer sMoneda
					%>
            </select>
            <select name="cbxParidades" style="Display: none">
              <%
						CargarParidadesTransfer afxTCTransferencia, sMoneda
					%>
          </select></td>
        <td style="display: none"><span class="textoempresa">Valuta</span><br>
            <select name="cbxValuta" class="Borde_tabla_abajo" style="HEIGHT: 22px; WIDTH: 100px">
              <%
						CargarValuta "48"
					%>
            </select>        </td>
      </tr>
      <tr height="15">
        <td></td>
        <td  style="display: none"><span class="textoempresa">Rate</span><br>
            <input name="txtParidad" disabled class="Borde_tabla_abajo" style="HEIGHT: 18px; TEXT-ALIGN: right; WIDTH: 90px" onKeyPress="IngresarTexto(1)" value="<%=FormatNumber(nParidad, 8)%>" size="15">
        </td>

        <td><span class="textoempresa">Equivalente USD</span><br>
            <input name="txtEquivalente" disabled class="cajas" style="HEIGHT: 18px; TEXT-ALIGN: right; WIDTH: 103px" value="<%=FormatNumber(nEquivalente, 2)%>" size="15">
            <img border="0" height="22" id="imgCalcular" name="imgCalcular" onMouseOver="imgCalcular.style.cursor ='Hand'" src="../images/BotonFlechaDerecha.jpg" style="LEFT: 0px; POSITION: relative; TOP: 5px; width:21; cursor: hand"> </td>
        <td style="display: none"><span class="textoempresa">Tarifa</span><br>
            <input name="txtTarifaSugerida" disabled class="cajas" style="HEIGHT: 18px; TEXT-ALIGN: right; WIDTH: 90px" value="<%=FormatNumber(0, 2)%>" size="15">        </td>
        <td style="display: none"><span class="textoempresa">Tarifa Cobrada</span><br>
            <input name="txtTarifaCobrada" class="cajas" id style="HEIGHT: 18px; TEXT-ALIGN: right; WIDTH: 90px" onKeyPress="IngresarTexto(1)" value="<%=FormatNumber(0, 2)%>" size="15">        </td>
        <td><span class="textoempresa">Total a Pagar</span><br>
            <input name="txtTotal" disabled class="cajas" id style="HEIGHT: 18px; TEXT-ALIGN: right; WIDTH: 90px" value="<%=FormatNumber(nTotal, 2)%>" size="15">        </td>
      </tr>      
      <tr height="10" style="display: none">
        <td><span class="titulos">Depósito para AFEX</span><br>
            <input type="radio" name="optPesos">
            <span class="textoempresa">Pesos</span><span class="subtitulos"><br>
              <input type="radio" name="optDolares">
              </span><span class="textoempresa">Dolares </span> </td>
        <td style="display:"><span class="textoempresa">T. Cambio</span><br>
            <input name="txtTipoCambio" disabled class="Borde_tabla_abajo" style="HEIGHT: 18px; TEXT-ALIGN: right; WIDTH: 90px; background-color: <%=sColorMoneda%>; color: white; font-weight: bold" value="<%=Request.Form("txtTipoCambio")%>" size="15">        </td>
        <td style="display:"><span class="textoempresa">Monto Pesos</span><br>
            <input name="txtMontoPesos" disabled class="Borde_tabla_abajo" style="HEIGHT: 18px; TEXT-ALIGN: right; WIDTH: 90px; background-color: <%=sColorMoneda%>; color: white; font-weight: bold" value="<%=Request.Form("txtMontoPesos")%>" size="15">        </td>
      </tr>
      
      <tr>
        <td colspan="5"><br>
            <br>
            <table border="0">
              <tr>
                <td><b class="titulos">Banco Destino</b></td>
              </tr>
              <tr>
				<td><span class="textoempresa">Nombre Beneficiario</span><br>
                    <input name="txtNombreExacto" class="Borde_tabla_abajo" style="HEIGHT: 18px; TEXT-ALIGN: left; WIDTH: 200px" onKeyPress="IngresarTexto(2)" value="<%=request.Form("txtNombreExacto")%>">
                </td>              
                <td colspan="2"><span class="textoempresa">Nombre Banco</span><br>
                    <input type="text" class="Borde_tabla_abajo" style="HEIGHT: 18px; TEXT-ALIGN: left; WIDTH: 200px" name="txtBanco" value="<%=request.Form("txtBanco")%>">
                    &nbsp;<input type="button" onClick="MostrarBanco1" value="...">
                    <select style="display: none" name="cbxBancoDestino" class="Borde_tabla_abajo">
                      <%CargarBanco sBanco%>
                    </select>                
                </td>
              </tr>
              <tr>
                <td><span class="textoempresa">Número Cuenta</span><br>
                    <input name="txtCtaCte" class="Borde_tabla_abajo" style="HEIGHT: 18px; TEXT-ALIGN: left; WIDTH: 90px" value="<%=request.Form("txtCtaCte")%>">
                </td>              
				<td>
					<span class="textoempresa">Referencia</span><br>
                    <input name="txtInvoice" class="Borde_tabla_abajo" style="HEIGHT: 18px; TEXT-ALIGN: left; WIDTH: 100px" value="<%=request.Form("txtInvoice")%>">
				</td>
			  </tr>
			  <tr>
				<td>
					<span class="textoempresa">Dirección del Beneficiario</span><br>
                    <input name="txtDireccionB" class="Borde_tabla_abajo" style="HEIGHT: 18px; TEXT-ALIGN: left; WIDTH: 200px" value="<%=request.Form("txtDireccionB")%>">
				</td>				
              </tr>
              <tr>
                <td><span class="textoempresa">Dirección del Banco</span><br>
                    <input name="txtDireccionBanco" class="Borde_tabla_abajo" style="HEIGHT: 18px; TEXT-ALIGN: left; WIDTH: 200px" value="<%=Request.Form("txtDireccionBanco")%>">
                </td>
                <td colspan="2"><span class="textoempresa">Ciudad del Banco</span><br>
                    <input type="text" class="Borde_tabla_abajo" style="HEIGHT: 18px; TEXT-ALIGN: left; WIDTH: 200px" name="txtCiudadBanco" value="<%=request.Form("txtCiudadBanco")%>">                    
               </td>
              </tr>              
              <tr>
                <td><b class="titulos">Códigos Bancarios</b></td>
              </tr>
              <tr>
                <td>
					<span class="textoempresa">ABA</span><br>
                    <input name="txtABA" class="Borde_tabla_abajo" style="HEIGHT: 18px; TEXT-ALIGN: left; WIDTH: 80px" value="<%=request.Form("txtABA")%>" size="15">
                </td>
                <td><span class="textoempresa">SWIFT</span><br>
                    <input name="txtSWIFT" class="Borde_tabla_abajo" style="HEIGHT: 18px; TEXT-ALIGN: left; WIDTH: 80px" value="<%=request.Form("txtSWIFT")%>" size="15">
                </td>
			    <td><span class="textoempresa">CHIPS</span><br>
                    <input name="txtCHIPS" class="Borde_tabla_abajo" style="HEIGHT: 18px; TEXT-ALIGN: left; WIDTH: 80px" value="<%=request.Form("txtCHIPS")%>" size="15">
                </td>                
              </tr>
              <tr>
                <td><b class="titulos">Códigos Bancarios Complementarios</b></td>
              </tr>
              <tr>
                <td><span class="textoempresa">IBAN</span><br>
                    <input name="txtIBAN" class="Borde_tabla_abajo" style="HEIGHT: 18px; TEXT-ALIGN: left; WIDTH: 80px" value="<%=request.Form("txtIBAN")%>" size="15">
                </td>
                <!--
                <td><span class="textoempresa">OTRO</span><br>
                    <input name="txtOTRO" class="Borde_tabla_abajo" value="<%=request.Form("txtOTRO")%>" size="15">
                </td>
                -->
              </tr>
            </table>
          <br>
          <br>
            <div align="left">
              <input name="chkBIntermedio" type="checkbox" value="<%=Request.Form("chkBIntermedio")%>">
              <span class="titulos"><b>Banco Intermediario</b> </span></div>
          <table id="tlbBIntermedio" border="0" style="display: none">
              <tr>
                <td></td>
              </tr>
              <tr>
                <td><span class="textoempresa">Nombre del Banco</span><br>
                    <input type="text" class="Borde_tabla_abajo" style="HEIGHT: 18px; TEXT-ALIGN: left; WIDTH: 100px" name="txtBancoInt" value="<%=request.Form("txtBancoInt")%>">
                    &nbsp;<input type="button" onClick="MostrarBanco2" value="...">
                    <select style="display: none" class="Borde_tabla_abajo" name="cbxBancoDestinoInt">
                      <%CargarBanco sBancoInt%>
                    </select>                </td>
                <td><span class="textoempresa">Nº de Cta.Cte.</span><br>
                    <input style="HEIGHT: 18px; TEXT-ALIGN: left; WIDTH: 90px" name="txtCtaCteInt" class="Borde_tabla_abajo" value="<%=request.Form("txtCtaCteInt")%>"></td>
                <!--<td>Nombre exacto de la Cta. del Beneficiario<br>
						<input style="HEIGHT: 18px; TEXT-ALIGN: left; WIDTH: 200px" name="txtNombreExactoInt" class="Borde_tabla_abajo" onKeyPress="IngresarTexto(2)" value="<%=request.Form("txtNombreExactoInt")%>">
                    </td>-->
              </tr>
              <tr>
                <td><span class="textoempresa">Dirección del Banco</span><br>
                    <input style="HEIGHT: 18px; TEXT-ALIGN: left; WIDTH: 200px" name="txtDireccionBancoInt" class="Borde_tabla_abajo" value="<%=request.Form("txtDireccionBancoInt")%>"></td>
                <td><span class="textoempresa">Ciudad del Banco</span><br>
                    <input type="text" style="HEIGHT: 18px; TEXT-ALIGN: left; WIDTH: 115px" name="txtCiudadBancoInt" class="Borde_tabla_abajo" value="<%=request.Form("txtCiudadBancoInt")%>">
                    
                </td>
              </tr>
              <!--
              <tr>
                <td class="titulos"><b>Códigos Bancarios</b></td>
              </tr>              
              <tr>
                <td>ABA<br>
                    <input style="HEIGHT: 18px;" size="20" name="txtABAInt" class="Borde_tabla_abajo" value="<%=request.Form("txtABAInt")%>">                </td>
				
                <td>CHIPS<br>
                    <input style="HEIGHT: 18px;" size="20" name="txtCHIPSInt" class="Borde_tabla_abajo" value="<%=request.Form("txtCHIPSInt")%>">                </td>
                <td>SWIFT<br>
                    <input style="HEIGHT: 18px;"  size="20" class="Borde_tabla_abajo" name="txtSWIFTInt" value="<%=request.Form("txtSWIFTInt")%>">                </td>
            
              </tr>
            
              <tr>
                <td><b>Códigos Bancarios Complementarios</b></td>
              </tr>
              <tr>
                <td>IBAN<br>
                    <input style="HEIGHT: 18px;" size="20" name="txtIBANInt" class="Borde_tabla_abajo" value="<%=request.Form("txtIBANInt")%>">                </td>
                <td>OTRO<br>
                    <input style="HEIGHT: 18px;" size="20" name="txtOTROInt" class="Borde_tabla_abajo" value="<%=request.Form("txtOTROInt")%>">                </td>
              </tr>
             -->
          </table></td>
      </tr>
      <tr>
        <td colspan="3" align="right" valign="center"><img align="right" alt border="0" hspace="0" id="imgAceptar" src="../images/BotonAceptar.jpg" style="FLOAT: right; cursor: hand" width="70" height="20"> </td>
      </tr>
    </table></td>
</tr>
</tbody>
</table>
</form>
</body>
<!--#INCLUDE virtual="/Compartido/Rutinas.htm" -->
<!--#INCLUDE virtual="/Compartido/Transferencias.htm" -->
</html>

<% set rs = nothing %>