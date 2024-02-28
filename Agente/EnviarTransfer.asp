<%@ Language=VBScript %>
<%
	'Se asegura que la página no se almacene en la memoria cache
	Response.Expires = 0
	If Session("CodigoCliente") = "" Then
		response.Redirect "../Compartido/TimeOut.htm"
		response.end
	End If
%>
<!--#INCLUDE virtual="/agente/Constantes.asp" -->
<!--#INCLUDE virtual="/Sucursal/Rutinas.asp" -->
<%

	Dim sMoneda, sVigencia, cObservado
	Dim nMonto, nParidad, nEquivalente, nTarifaSugerida, nTarifaCobrada
	Dim nGastoExtranjera, nGastoNacional, nTotal, nEstadoTRF
	
	On Error Resume Next
	
	' JFMG 25-09-2009 se agrega para sacar la uf del día
	Dim cUFDia, cDolarObservado
	Dim sSQL, rsUFDia, rsDolarObservado
	
	cUFDia = 0	
	sSQL = " select dbo.mostrarufdia() as uf "
	set rsUFDia = ejecutarsqlcliente(Session("afxCnxCorporativa"), sSQL)
	if err.number <> 0 then
		mostrarerrorms "Consultar UF"
	end if
	if not rsUFDia.eof then
		cUFDia = rsUFDia("uf")
	end if
	
	set rsUFDia = nothing
	
	cDolarObservado = 0
	sSQL = " select giros.dbo.MostrarTipoCambioObservado() as dolarobservado "
	set rsDolarObservado = ejecutarsqlcliente(Session("afxCnxAFEXpress"), sSQL)
	if err.number <> 0 then
		mostrarerrorms "Consultar DolarObservado"
	end if
	if not rsDolarObservado.eof then
		cDolarObservado = rsDolarObservado("dolarobservado")
	end if
	
	set rsDolarObservado = nothing
	' ********* FIN JFMG 25-09-2009 ********************
	
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
		ValidarMonto		
		CalcularMontos
	End Sub 

	Sub txtTarifaCobrada_OnBlur()
		If Trim(frmTransfer.txtTarifaCobrada.value) = "" Then
			frmTransfer.txtTarifaCobrada.value = "0,00"
		Else
			frmTransfer.txtTarifaCobrada.value = FormatNumber(frmTransfer.txtTarifaCobrada.value, 2)
		End If
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
		Dim sEjecutivoParidad, sReferencia
		
		' JFMG 25-09-2009 se agrega para mostrar información de como realizar una transfer cuando esta supera los 450UF		
		if ccur("0" & ccur("0" & "<%=cUFDia%>")) * 450 <= ccur("0" & frmTransfer.txtTotal.value) * ccur("0" & "<%=cDolarObservado%>") then
			if CajaPreguntaSiNo("AFEX En Linea", "Recuerde que esta Transferencia debe estar INGRESADA en " & vbCrlf & _
													"AFEXchange con el mismo monto y forma de pago. " & vbCrlf & _
													"¿Desea ver la presentación en Línea de como realizar una Transferencia?" & vbcrlf & _
													"Si la quiere ver en este momento presione el botón ""SI"", de lo contrario puede seleccionar" & vbcrlf & _
													"en el menú de Opciones ""¿Como Vender y Enviar una Transferencia?""") Then
				linkCumplimiento
			end if
		end if
		' ************** FIN JFMG 25-09-2009 *******************
		
			
		If Not CajaPregunta("AFEX En Linea", "Está seguro que desea enviar la transferencia?") Then
			Exit Sub
		End If

		If Not ValidarDatos() Then
			Exit Sub
		End If		
		
		' verifica si es mayor a US$ 5.000
		if frmTransfer.txtEquivalente.value > 5000 and frmTransfer.cbxMoneda.value <> "USD" then
			sEjecutivoParidad = inputbox("Ingrese el nombre del ejecutivo con el que cerró la paridad para esta Transferencia que supera los US$ 5.000.","Transferencias")
			sReferencia = inputbox("Ingrese una Referencia para la Transferencia.","Transferencias")
			if sEjecutivoParidad = empty or sReferencia = empty then 
				frmTransfer.txtMonto.value = "0,00"
				exit sub
			end if
			frmTransfer.txtEjecutivoParidad.value = sEjecutivoParidad
			frmTransfer.txtRefernciaTransfer.value = sReferencia
		end if		
			
		HabilitarControles
		frmTransfer.action = "GrabarEnvioTransfer.asp"
		frmTransfer.submit 
		frmTransfer.action = ""			
	End Sub

	'Objetivo:	calcular los montos de la página
	Sub CalcularMontos()
	
		
		' si la tarifa cobrada es distinta de la sugerida no se calcula
		If frmTransfer.txtTarifaCobrada.value = frmTransfer.txtTarifaSugerida.value Then
			frmTransfer.txtTarifaCobrada.value = ObtenerTarifaTransfer(frmTransfer.cbxValuta.value, _
														   frmTransfer.cbxMoneda.value, frmTransfer.txtMonto.value)
		End If
		
		' calcula la tarifa
		frmTransfer.txtTarifaSugerida.value = ObtenerTarifaTransfer(frmTransfer.cbxValuta.value, _
														frmTransfer.cbxMoneda.value, frmTransfer.txtMonto.value)		
		' calcula la comisión
		frmTransfer.txtGastoExtranjera.value = ObtenerComisionTransfer(frmTransfer.cbxValuta.value, _
													    frmTransfer.cbxMoneda.value, frmTransfer.txtMonto.value)
		frmTransfer.txtGastoNacional.value = FormatNumber(cDbl(0 & frmTransfer.txtGastoExtranjera.value) * _
												   cCur(0 & "<%=cObservado%>"), 0)
		' calcula el total
		frmTransfer.txtTotal.value = CalcularTotal
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
		sCodigo = frmTransfer.cbxMoneda.value
		frmTransfer.cbxParidades.value = sCodigo		
		if sCodigo = empty or frmTransfer.cbxParidades.value = empty then
			frmTransfer.txtParidad.value = "0,00000000"		
		
		else
			If frmTransfer.cbxRate.value = "1" Or frmTransfer.cbxMoneda.value = "USD" Then
				frmTransfer.txtParidad.value  = FormatNumber( _											
								frmTransfer.cbxParidades(frmTransfer.cbxParidades.selectedIndex).text _
								, 8)
			End If
		end if				
		CalcularUSD
		ValidarMonto
		' calcula los montos
		CalcularMontos
	End Sub
	
	Sub cbxMoneda_onChange()
		dim bVigencia		
		If frmTransfer.cbxMoneda.value = "USD" Then
			if frmTransfer.cbxMoneda.value <> empty then			
				bVigencia = window.showModalDialog("../Compartido/ValidarParidad.asp?Moneda=" & frmTransfer.cbxMoneda.value)
				if not bVigencia then
					msgbox "En este momento no puede enviar transferencias por la moneda seleccionada ya que no tiene Paridad Vigente. ",, "Transferencias"
					frmTransfer.cbxMoneda.value = empty
				end if
			end if
		else
			msgbox "ATENCIÓN : Por disposición de la GERENCIA DE AFEX indefinidamente sólo se pueden enviar Transferencias en DÓLAR, " & _
				   " para otras monedas debe comunicarse con la Agencia de Valores: " & _
				   " Alejandro Morales (alejandro.morales@afexav.cl) o Constanza San Martín (constanza.sanmartin@afexav.cl) " & _
				   "para cerrar precio, Telefonos: 6369072 y 6369152 respectivamente. " ,,"AFEX EN LÍNEA."
			frmTransfer.cbxMoneda.value = empty 			
		end if	
		sCodigo = frmTransfer.cbxMoneda.value  
		frmTransfer.cbxParidades.value = sCodigo		
		If frmTransfer.cbxRate.value <> "1" And frmTransfer.cbxMoneda.value <> "USD" Then
			msgbox "Verifique que el Rate ingresado corresponda a la moneda Seleccionada. ",, "Transferencias"
		End If
		
		if frmTransfer.cbxMoneda.value = "USD" then
			frmTransfer.cbxRate.value = "1"
			frmTransfer.txtParidad.value = "1,00000000"
			cbxRate_onClick()
		end if
		
		cbxMoneda_onClick()
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
	sub cbxPago_onclick()
		dim pago,lugar		
			If frmTransfer.cbxpago.value = 2 then
				frmTransfer.formas.disabled  = false		
				frmTransfer.formas1.disabled  = false
			else
				frmTransfer.formas.disabled = true			
				frmTransfer.formas.checked = false	
				frmTransfer.formas1.disabled = true			
				frmTransfer.formas1.checked = false					
			end if		
		pago=frmTransfer.cbxpago.value
	end sub	
	sub formas_onclick()
		frmTransfer.formas1.checked = false	
	end sub
	sub formas1_onclick()
		frmTransfer.formas.checked = false
	end sub
	
	Sub cbxRate_onClick()
		sCodigo = frmTransfer.cbxMoneda.value
		frmTransfer.cbxParidades.value = sCodigo
		If frmTransfer.cbxRate.value = "1" Then
			frmTransfer.txtParidad.value  = FormatNumber( _											
							frmTransfer.cbxParidades(frmTransfer.cbxParidades.selectedIndex).text _
							, 8)
			frmTransfer.txtParidad.disabled = True
			ValidarMonto
		Else
			frmTransfer.txtParidad.disabled = False
		End If
		CalcularUSD		
		CalcularMontos		
	End Sub
	
	Sub CalcularUSD
		' Jonathan Miranda G. 31-10-2006
		' si la moneda es dolar el rate es automatico
		if frmTransfer.cbxMoneda.value = "USD" and frmTransfer.cbxRate.value = 2 then
			frmTransfer.cbxRate.value = 1
			cbxRate_onClick
		end if
		'---------------- Fin ------------------
	
		'Calcula el Equivalente en Dolares
		frmTransfer.txtEquivalente.value = FormatNumber( _
				Round(CDbl(0 & frmTransfer.txtMonto.value) * CDbl(0 & frmTransfer.txtParidad.value), 7), 2)		
	End Sub
	
	Sub ValidarMonto
		' verifica si el monto es mayor a USD 5000, para que el rate sea manual e ingrese un ejecutivo para cerrar el
		' el valor del rate
		'if frmTransfer.cbxMoneda.value <> empty then
		'	cMonto = frmTransfer.txtMonto.value * frmTransfer.txtParidad.value
			if frmTransfer.txtEquivalente.value > 5000 and frmTransfer.cbxmoneda.value <> "USD" then
				frmTransfer.cbxRate.value = 2
				cbxRate_onClick()
			end if
		'end if
	End Sub

	'Objetivo:	calcular el total de la transferencia
	Function CalcularTotal()
		CalcularTotal = 0
	
		' verifica si hay monto
		If frmTransfer.txtMonto.value = Empty Or frmTransfer.txtMonto.value = "0,00" Then Exit Function
		
		CalcularTotal = FormatNumber(cDbl(0 & frmTransfer.txtTarifaCobrada.value) + _
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
			aTrf = Split(sString, ";", 15)
			
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
			frmTransfer.txtDireccionBeneficiario.value = aTrf(14)
			<% bTransferencia = True %>					
		End If		
	End Sub

	Function ValidarDatos()
		ValidarDatos = False		
		If frmTransfer.cbxMoneda.value = empty Then
			MsgBox "Debe seleccionar una moneda para la transferencia",,"AFEX"			
			frmTransfer.cbxMoneda.focus
			Exit Function
		End If
		If cCur(0 & frmTransfer.txtParidad.value) <= 0 Then
			MsgBox "Debe ingresar un rate válido para la transferencia",,"AFEX"
			'frmTransfer.txtMonto.select
			'frmTransfer.txtMonto.focus
			Exit Function
		End If
		If cCur(0 & frmTransfer.txtMonto.value) <= 0 Then
			MsgBox "Debe ingresar un monto válido para la transferencia",,"AFEX"
			frmTransfer.txtMonto.select
			frmTransfer.txtMonto.focus
			Exit Function
		End If
		if frmTransfer.cbxpago.value = 2 and (frmTransfer.formas.checked = false) and (frmTransfer.formas1.checked = false) then
			MsgBox "Debe Marcar una de las opciones en Lugar del Cheque",,"AFEX"
						frmTransfer.formas.focus 
			Exit Function
		end if
		if frmTransfer.cbxpago.value = 0 then
			MsgBox "Debe ingresar Forma de Pago",,"AFEX"		
			frmTransfer.cbxpago.focus
			exit function
		end if
		If frmTransfer.cbxPaisBeneficiario.value = "" Then
			MsgBox "Debe seleccionar el país del Beneficiario",,"AFEX"			
			frmTransfer.cbxPaisBeneficiario.focus
			Exit Function
		End If
		ValidarDatos = True
	End Function

	Sub CargarMenu()
		Dim sId

		frmTransfer.objmenu.bgColor = document.bgColor 
		frmTransfer.objmenu.stylesheet = "../Estilos/Cliente.css"
		sId = frmTransfer.objmenu.addparent("Opciones")

		frmTransfer.objMenu.addchild sId, "Imprimir Transferencia", "Imprimir", "Principal"
		' JFMG 25-09-2009 se agrega para ver el link de como vender una transfer
		frmTransfer.objMenu.addchild sId, "¿Como Vender y Enviar una Transferencia?", "LinkCumplimiento", "Principal"
		' *********** FIN JFMG 25-09-2009 ****************************
		
		frmTransfer.objMenu.addchild sId, "", "", ""
	End Sub

	sub window_onload()
		' JFMG 25-09-2009 se agrega para mostrar un mensaje a los usuarios
		tbMensajeTransferencia.style.display = ""
		window.setInterval "IniciarVentana", 30000, "vbscript"		
		' ************** FIN JFMG 25-09-2009 ******************		
		
		
		CargarMenu
		cbxMoneda_onClick
		frmTransfer.txtmonto.select
		frmTransfer.txtmonto.focus
		frmTransfer.formas.disabled = true
		frmTransfer.formas1.disabled = true			
		
		frmTransfer.cbxMoneda.value = Trim("<%=sMoneda%>")
		cbxMoneda_onClick()
		
	end sub

	sub cbxparidades_onchange()
		msgbox frmTransfer.cbxParidades.value 
	end sub

	
	' JFMG 25-09-2009 se agrega para mostrar un mensaje a los usuarios	
	sub IniciarVentana()		
		
		tbMensajeTransferencia.style.display = "none"
		tbDatosTransferencia.style.display = ""
		frmTransfer.objmenu.style.display = ""
		
		window.setInterval ""
	end sub
	
	sub SalirMensaje()
		IniciarVentana
	end sub
	' ************** FIN JFMG 25-09-2009 ******************

-->
</script>

<body>
<!-- #INCLUDE virtual="/Compartido/Encabezado.htm" -->
<form id="frmTransfer" method="post">
<input type="hidden" name="txtExchange" value="<%=Request.Form("txtExchange")%>">
<input type="hidden" name="txtExpress" value="<%=Request.Form("txtExpress")%>">
<input type="hidden" name="optPersona" value="<%=Request.Form("optPersona")%>">
<input type="hidden" name="optEmpresa" value="<%=Request.Form("optEmpresa")%>">
<input type="hidden" name="txtDireccion" value="<%=Request.Form("txtDireccion")%>">
<input type="hidden" name="cbxComuna" value="<%=Request.Form("cbxComuna")%>">
<input type="hidden" name="cbxCiudad" value="<%=Request.Form("cbxCiudad")%>">
<input type="hidden" name="cbxPais" value="<%=Request.Form("cbxPais")%>">
<input type="hidden" name="txtPaisFono" value="<%=Request.Form("txtPaisFono")%>">
<input type="hidden" name="txtAreaFono" value="<%=Request.Form("txtAreaFono")%>">
<input type="hidden" name="txtFono" value="<%=Request.Form("txtFono")%>">
<input type="hidden" name="txtRut" value="<%=Request.Form("txtRut")%>">
<input type="hidden" name="txtPasaporte" value="<%=Request.Form("txtPasaporte")%>">
<input type="hidden" name="cbxPaisPasaporte" value="<%=Request.Form("cbxPaisPasaporte")%>">
<input type="hidden" name="txtPaisFax" value="0&gt;">
<input type="hidden" name="txtAreaFax" value="0">
<input type="hidden" name="txtFax" value="0">
<input type="hidden" name="txtMensaje" value>
<input type="hidden" name="txtIdentificacion" value>
<input type="hidden" name="txtEjecutivoParidad" value>
<input type="hidden" name="txtRefernciaTransfer" value>

<!-- Paso 1 -->
<table id="tbMensajeTransferencia" style="display: none;" border="0">
	<tr>
		<td colspan="2">&nbsp;		
			<table class="borde">
				<tr>		
					<td width="80%" align="justify">
						<b>Sres. Usuarios:</b><br><br>		
						Se les comunica que luego de algunas revisiones en las operaciones de 
						Venta y Envío de Transferencias se han encontrado diferencias en el ingreso 
						que cada Sucursal realiza en el Sistema AFEXchange y la página WEB de AFEX en Línea, 
						lo cual provoca inconvenientes a los Departamentos de Informática, Cumplimiento, 
						Gestión y Contabilidad, por lo cual se ha creado una presentación Power Point 
						para que todos Uds. puedan revisar como se realiza el correcto ingreso de los datos 
						en los distintos Sistemas involucrados. Si desea ver la presentación en este momento 
						haga click en botón "Ver presentación", de lo contrario puede seleccionar 
						"¿Como Vender y Enviar una Transferencia?" en el menú 
						"Opciones" de la página. 
						<br>
						<br>						
					</td>
				</tr>
				<tr>
					<td align="center">
						<input type="button" onclick="LinkCumplimiento" value="Ver Presentación" id=button1 name=button1>
						&nbsp;
						<input type="button" onclick="SalirMensaje" value="Salir" id=button2 name=button2>
					</td>
				</tr>
			</table>
		</td>
	</tr>
</table>
<table id="tbDatosTransferencia" border="0" style="display: none;">
	<tr>
		<td width="60%">
			<table class="Borde" ID="tabPaso1" CELLSPACING="0" CELLPADDING="0" BORDER="0" HEIGHT="100" WIDTH="560" STYLE="HEIGHT: 100px; LEFT: 4px; POSITION: relative; TOP: 0px; WIDTH: 566px">
			<tr><td colspan="2" align="center" style="display: none"><font FACE="Verdana" SIZE="2"><%=sVigencia%></font></td></tr>
			<tr><td class="titulo">Datos del Remitente</td></tr>
			<tr><td>
				<table ID="tabRemitente" width="100%">
				<tr HEIGHT="15">
							
				<% 
					Dim sDisplayEmpresa, sDisplayPersona
					
					
					If Request.Form("optPersona") = "on" Then
						sDisplayEmpresa = "none"
						sDisplayPersona = ""
					Else
						sDisplayEmpresa = ""
						sDisplayPersona = "none"
					End If
				%>
					<td></td>
					<td style="display: <%=sDisplayPersona%>">Nombres<br><input NAME="txtNombres" SIZE="25" style="HEIGHT: 22px; WIDTH: 240px" disabled value="<%=request.Form("txtNombres")%>"></td>
					<td style="display: <%=sDisplayPersona%>">Apellido Paterno<br>
						<input NAME="txtApellidoP" SIZE="20" style="HEIGHT: 22px; WIDTH: 110px" disabled value="<%=request.Form("txtApellidoP")%>">
					</td>
					<td style="display: <%=sDisplayPersona%>">Apellido Materno<br>
						<input NAME="txtApellidoM" SIZE="20" style="HEIGHT: 22px; WIDTH: 110px" disabled value="<%=request.Form("txtApellidoM")%>">
					</td>
					<td style="display: <%=sDisplayEmpresa%>">Nombres<br>
						<input NAME="txtRazonSocial" SIZE="25" style="HEIGHT: 22px; WIDTH: 460px" disabled value="<%=request.Form("txtRazonSocial")%>">
					</td>
					<td><img border="0" id="imgTransferencia" src="../images/Transferencia.jpg" style="cursor: hand" WIDTH="19" HEIGHT="22" onClick="UltimasTransferencias"></td>
					<td style="cursor: hand" onClick="UltimasTransferencias">Ultimas Transferencias</td>
				</tr>
				</table>		
				<table>
				<tr>
				</tr>
				</table>
			</td></tr>
			<tr><td class="titulo">Valores</td></tr>		
			<tr><td>
				<table ID="tabTransfer" border="0" cellspacing="2" cellpadding=0>
				<tr>
					<td></td>
					<!--<td colspan="2">Banco<br>
						<select name="cbxBanco" style="width: 240px">
							<%	CargarBanco nCodigoBanco%>
						</select>
					</td>-->
					<td colspan="1" >Estado Transferencia<br>
						<select name="cbxEstado" style="width: 115px">
							<%	CargarEstadoTRF nEstadoTRF%>
						</select>
					</td>
					<td colspan="1">&nbsp; Forma de Pago<br>
						&nbsp;<select name="cbxpago" style="width: 115px">	
						<option value="0">--------</option>	 
					    <option value="1">Efectivo</option>
					    <option value="2">Cheque</option>                         
						</select>
					</td>
					<td colspan="1">Lugar del Cheque<br>
						<input type="radio"  name="formas" value="1"  disabled>Sobre Plaza <br>
						<input type="radio"  name="formas1" value="2" disabled>Fuera Plaza 	
					</td>
					<td>País Beneficiario<br>
						<select NAME="cbxPaisBeneficiario" style="HEIGHT: 22px; WIDTH: 166px">
							<%					
								CargarPaisTransfer ""
							%>
						</select>
					</td>
					
				</tr>
				<tr HEIGHT="15">
				<td></td>
				<td>Monto<br>
					<input STYLE="HEIGHT: 22px; TEXT-ALIGN: right; WIDTH: 115px" NAME="txtMonto" SIZE="15" onkeypress="IngresarTexto(1)" value="<%=FormatNumber(nMonto, 2)%>">
				</td>				
				<td>Monedas de Pago<br>
					<select NAME="cbxMoneda" style="HEIGHT: 22px; WIDTH: 166px">
					<%
						'CargarMonedasTransfer sMoneda
						CargarMonedasMantenedor
					%>
					</select>							
					<select name="cbxParidades" STYLE="Display: none">	
					<%
						'CargarParidadesTransfer afxTCTransferencia, sMoneda
						CargarParidadesMantenedor
					%>
					</select></td>
					<td>Rate<br>
					<select NAME="cbxRate" style="HEIGHT: 22px; WIDTH: 106px">
						<option SELECTED VALUE="1">Automática</option>
						<option VALUE="2">Manual</option>					
					</select></td>			
				<td>Valuta<br>
					<select NAME="cbxValuta" style="HEIGHT: 22px; WIDTH: 101px">
						<%
							CargarValuta "48"
						%>
			<!--			<option SELECTED VALUE="48">48 hrs</option>			<option VALUE="24">24 hrs</option>			<option VALUE="0">Express</option>-->
					</select>
				</td>			
				</tr>
				<tr HEIGHT="15">
				<td></td>
				<td>Rate<br>
					<input STYLE="HEIGHT: 22px; TEXT-ALIGN: right; WIDTH: 93px" NAME="txtParidad" SIZE="15" disabled onkeypress="IngresarTexto(1)" value="<%=FormatNumber(nParidad, 8)%>">
				</td>
				<td>Equivalente USD<br>
					<input STYLE="HEIGHT: 22px; TEXT-ALIGN: right; WIDTH: 103px" NAME="txtEquivalente" SIZE="15" disabled value="<%=FormatNumber(nEquivalente, 2)%>">
					<img border="0" height="22" id="imgCalcular" name="imgCalcular" onmouseover="imgCalcular.style.cursor ='Hand'" src="../images/BotonFlechaDerecha.jpg" style="LEFT: 0px; POSITION: relative; TOP: 5px; width:21; cursor: hand">
				</td>
				<td>Tarifa Sugerida<br>
					<input STYLE="HEIGHT: 22px; TEXT-ALIGN: right; WIDTH: 100px" NAME="txtTarifaSugerida" SIZE="15" disabled value="<%=FormatNumber(nTarifaSugerida, 2)%>">
				</td>
				<td>Tarifa Cobrada<br>
					<input STYLE="HEIGHT: 22px; TEXT-ALIGN: right; WIDTH: 100px" ID NAME="txtTarifaCobrada" SIZE="15" onkeypress="IngresarTexto(1)" value="<%=FormatNumber(nTarifaCobrada, 2)%>">
				</td>
				</tr>
				<tr HEIGHT="15">
				<td></td>
				<td>Gasto USD<br>
					<input STYLE="HEIGHT: 22px; TEXT-ALIGN: right; WIDTH: 93px" ID NAME="txtGastoExtranjera" SIZE="15" disabled value="<%=FormatNumber(nGastoExtranjera, 2)%>">
				</td>
				<td>Gasto Pesos<br>
					<input STYLE="HEIGHT: 22px; TEXT-ALIGN: right; WIDTH: 100px" ID NAME="txtGastoNacional" SIZE="15" disabled value="<%=FormatNumber(nGastoNacional, 2)%>">
				</td>
				<td>Total a Pagar<br>
					<input STYLE="HEIGHT: 22px; TEXT-ALIGN: right; WIDTH: 115px" ID NAME="txtTotal" SIZE="15" disabled value="<%=FormatNumber(nTotal, 2)%>">
				</td>
				<td align="right" valign="center">
					<img align="right" alt border="0" hspace="0" id="imgAceptar" src="../images/BotonAceptar.jpg" style="FLOAT: right; cursor: hand" WIDTH="70" HEIGHT="20">
				</td>
				</tr>
				</table>
			</td></tr>
			<tr ID="trDatosTransfer1"><td class="titulo">Datos de la Transferencia &nbsp;
				<font ID="dspDatosTransfer" FACE="Marlett" SIZE="3">u</font> 
			</td></tr>
			<tr ID="trDatosTransfer2" style="DISPLAY: none"><td>
				<table ID="tabBeneficiario" width="100%" border="0">
			   <tbody>
				<tr><td>
					<table width="100%">
					<tr><td colspan="3" class="titulo" style="BACKGROUND-COLOR: lightgrey">Origen</td></tr>
					<tr HEIGHT="15">
					<td></td>
					<td>Banco <br><input NAME="txtBancoOrigen" SIZE="25" style="HEIGHT: 22px; WIDTH: 301px" disabled></td>
					<td>Cuenta Corriente<br><input NAME="txtCuentaOrigen" SIZE="25" style="HEIGHT: 22px; WIDTH: 179px" disabled></td>
					</tr>
					</table>
				</td></tr>
				<tr><td>
					<table width="100%" border="0">		
					<tr><td colspan="3" class="titulo" style="BACKGROUND-COLOR: lightgrey">Destino</td></tr>
					<tr HEIGHT="15">
					<td></td>
					<td>Banco <br><input NAME="txtBancoDestino" SIZE="25" style="HEIGHT: 22px; WIDTH: 299px" disabled></td>
					<td>Cuenta Corriente<br><input name="txtCuentaDestino" SIZE="25" style="HEIGHT: 22px; WIDTH: 179px" disabled></td>
					</tr>
					<tr HEIGHT="15">
					<td></td>
					<td colspan="2">Nombre Beneficiario<br><input name="txtNombreB" SIZE="25" style="HEIGHT: 22px; WIDTH: 299px" disabled></td>
					</tr>
					<tr HEIGHT="15">
					<td></td>
					<td>Dirección<br><input NAME="txtDireccionB" SIZE="50" style="HEIGHT: 22px; WIDTH: 321px" disabled></td>
					<td>Ciudad<br><input NAME="txtCiudadB" style="HEIGHT: 22px; WIDTH: 153px" disabled></td>
					<td>Dirección Beneficiario<br><input NAME="txtDireccionBeneficiario" SIZE="50" style="HEIGHT: 22px; WIDTH: 321px" disabled></td>
					</tr>
					<tr HEIGHT="15">
					<td></td>
					<td>ABA<br><input NAME="txtAba" SIZE="25" style="HEIGHT: 22px; WIDTH: 126px" disabled></td>
					<td>Invoice<br><input ID="txtInvoice" NAME="txt" SIZE="25" style="HEIGHT: 22px; WIDTH: 179px" disabled></td>
					<td>Crédito a<br><input NAME="txtFurtherCredit" SIZE="25" style="HEIGHT: 22px; WIDTH: 179px" disabled></td>
					</tr>
					</table>
				</td></tr>
				<tr><td width="60%">
					<table width="100%" border="0">		
					<tr><td colspan="3" class="titulo" style="BACKGROUND-COLOR: lightgrey">Intermediario</td></tr>
					<tr HEIGHT="15">
					<td></td>
					<td>Banco<br><input NAME="txtBancoIntermedio" SIZE="25" style="HEIGHT: 22px; WIDTH: 303px" disabled></td>
					<td>Cuenta Corriente<br><input NAME="txtCuentaIntermedio" SIZE="25" style="HEIGHT: 22px; WIDTH: 179px" disabled></td>
					</tr>
					<tr HEIGHT="15">
					<td></td>
					<td>Dirección<br><input NAME="txtDireccionIntermedio" SIZE="50" style="HEIGHT: 22px; WIDTH: 321px" disabled></td>
					<td>Ciudad<br><input NAME="txtCiudadIntermedio" style="HEIGHT: 22px; WIDTH: 121px" disabled></td>
					</tr>
					</table>
				</td>
				</td></tr>
				</tbody>
				</table>
			</td></tr>
			</tbody>
			</table>
		</td>		
	</tr>
	
	<table BORDER="0" cellspacing="0" cellpadding="0" STYLE="LEFT: 437px; POSITION: absolute; TOP: 50px">	
		<tr><td>
		    <object align="left" id="objMenu" style="HEIGHT: 60px; LEFT: 0px; POSITION: relative; TOP: 0px; WIDTH: 280px; display:none" type="text/x-scriptlet" width="170" VIEWASTEXT border="0" valign="top"><param NAME="Scrollbar" VALUE="0"><param NAME="URL" VALUE="http:../Scriptlets/Menu.htm"></object>
		</td></tr>
	</table>

</table>

</form>
</body>
<script LANGUAGE="VBScript">
<!--
	Sub objMenu_OnScriptletEvent(strEventName, varEventData)
		Select Case strEventName
			Case "linkClick"
				If Right(varEventData, 8) = "Imprimir" Then
					ImprimirTransferencia
				End If
				
				If Right(varEventData, 16) = "LinkCumplimiento" Then
					linkCumplimiento
				End If
		End Select
	End Sub
	
	Sub ImprimirTransferencia()
		Dim sTo, sFonoCliente, sFaxCliente, sNombreContacto
		Dim sParamatros, sReleaseD, sValueD, sNombreCliente
		Dim sDireccionB
		
		sNombreCliente = trim(trim(frmTransfer.txtNombres.value) & " " & _
						 trim(frmTransfer.txtApellidoP.value) & " " & _
						 trim(frmTransfer.txtApellidoM.value) & _
						 trim(frmTransfer.txtRazonSocial.value))
		
		sDireccionB = ""
		
		If trim(frmTransfer.txtDireccionB.value) <> Empty Then
			sDireccionB = trim(trim(frmTransfer.txtDireccionB.value) & ", " & _
						  trim(frmTransfer.txtCiudadB.value))
		End If					  
		
		sFonoCliente = ""
		sFaxCliente = ""
		
		If frmTransfer.txtFono.value > 0 Then
			sFonoCliente = "(" + frmTransfer.txtPaisFono.value & " " & _
								 frmTransfer.txtAreaFono.value & ")" & " " & _
								 frmTransfer.txtFono.value
		End If
		If frmTransfer.txtFax.value > 0 Then
			sFaxCliente = "(" + frmTransfer.txtPaisFax.value & " " & _
 								frmTransfer.txtAreaFax.value & ")" & " " & _
								frmTransfer.txtFax.value
		End If
		
		sReleaseD = date
		If frmTransfer.cbxValuta.value = "24" Then
			sValueD = date + 1
		ElseIf frmTransfer.cbxValuta.value = "48" Then
			sValueD = date + 2
		Else
			sValueD = date
		End If
		
		sTo = "<%=Session("NombreCliente")%>"
		
		sParamatros = "ImprimirTransfer.asp?" & _
	 				  "pr=AFEX TRANSFERENCIAS" & _
	 				  "&fc=<%=Date%>" & _
	 				  "&mn=" & frmTransfer.cbxMoneda.value & _
					  "&bb=" & replace(frmTransfer.txtBancoDestino.value, "&", " @ ") & _
					  "&ct=" & frmTransfer.txtCuentaDestino.value & _
					  "&bf=" & frmTransfer.txtNombreB.value & _
					  "&fa=" & frmTransfer.txtABA.value & _
					  "&ab=" & sDireccionB & _
					  "&prompt9=" & _
					  "&mt=" & frmTransfer.txtMonto.value & _
					  "&prompt11= &prompt12= " & _
					  "&ms=" & frmTransfer.txtMensaje.value & _
					  "&prompt14= " & _
					  "&iv=" & frmTransfer.txtInvoice.value & _
					  "&ft=" & frmTransfer.txtFurtherCredit.value & _
					  "&rf=" & frmTransfer.txtIdentificacion.value & _
					  "&rd=" & sReleaseD & _
					  "&vd=" & sValueD & _
					  "&ib=" & frmTransfer.txtBancoIntermedio.value & _
					  "&ca=" & frmTransfer.txtCuentaIntermedio.value & _
					  "&dd=" & sTo &  _
					  "&prompt22=AFEX TRANSFERENCIAS" & _
					  "&nb=" & sNombreCliente & _
					  "&rt=" & FormatoRut(frmTransfer.txtRut.value) & _
					  "&tl=" & sFonoCliente & _
					  "&fx=" & sFaxCliente & _
					  "&co=" & sNombreContacto & _
					  "&prompt28=" & "<%=session("NombreCliente")%>" & _
					  "&db=" & frmTransfer.txtDireccionBeneficiario.value
					  
		sParamatros = replace(sParamatros, "ñ", "n")
		sParamatros = replace(sParamatros, "Ñ", "N")
		
		window.open sParamatros, _
				    "", "dialogHeight= 800pxl; dialogWidth= 800pxl; " & _
					"dialogTop= 0; dialogLeft= 0; resizable=no; " & _
					"status=no; scrollbars=yes"

	End Sub
	
	
	sub LinkCumplimiento()		
		window.open "<%=session("URLAyudaTransferencia") %>" , "", "width=1100,top=0,height=1000"
	end sub
	
-->
</script>
<!--#INCLUDE virtual="/Compartido/Rutinas.htm" -->
<!--#INCLUDE virtual="/Compartido/Transferencias.htm" -->
</html>