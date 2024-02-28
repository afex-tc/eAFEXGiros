<%@ Language=VBScript %>
<!--#INCLUDE virtual="/Compartido/TimeOut.asp" -->
<!--#INCLUDE virtual="/Compartido/Constantes.asp" -->
<%

	Dim nAccion, sPais, sCiudad, sPagador, sMoneda, nMonto
	Dim nTarifaSugerida, nTarifaCobrada, bCliente, bExtranjero
	Dim nTarifa, nGastoTransfer, nComisionCaptador, nComisionPagador
	Dim nComisionMatriz, nAfectoIva, nDDIpais, nDDICiudad
	Dim sNombreB, sApellidoB, sDireccionB, sFonoB, bGiroAnterior, sDisplay
	Dim bGiroNacional, nDec, sColorMoneda
	Dim sMensaje, rs, sNota, sMaximo, sRecomendacion, sRequerimiento
	Dim bCargarMensaje
	
	bCargarMensaje = False
	nAccion = cInt(0 & Request("Accion"))
	nTarifa = cCur(0)
	nTarifaCobrada = cCur(0)
	nMonto = cCur(0)
	nGastoTransfer = cCur(0)
	nComisionCaptador = cCur(0)
	nComisionPagador = cCur(0)
	nComisionMatriz = cCur(0)
	nAfectoIva = cCur(0)
	If Session("Categoria") = 4 Then		
		bExtranjero = True
		sDisplay = "none"
	Else
		bExtranjero = False
		sDisplay = ""
	End If
	If Request.Form("txtExchange") <> "" Or Request.Form("txtExpress") <> "" Then
		bCliente = True
	Else
		bCliente = False
	End If
	bGiroAnterior = False

	Select Case nAccion
	Case afxAccionPais, afxAccionCiudad, afxAccionPagador
		If nAccion = afxAccionPagador Then
			Set rs = BuscarInformacionPagador(Session("afxCnxAFEXpress"), Request.Form("cbxPagador"), _
											  Request.Form("cbxPaisB"), Request.Form("cbxCiudadB"))
			'window.showModalDialog "InformacionPagador.asp?PaisB=" & frmGiro.cbxPaisB.value & "&CiudadB=" & frmGiro.cbxCiudadB.value,,"dialogWidth:20; dialogHeight:10"
			If Err.Number <> 0 Then
				Set rs = Nothing
				
				Response.Redirect "../Compartido/Error.asp?description=" & err.Description
			End If
		
			sMensaje = Empty
		
			If Not rs.EOF Then
				bCargarMensaje = True
				sNota = rs("nota")
				If IsNull(rs("maximo")) Then
					sMaximo = ""
				Else
					sMaximo = "MONTO MAXIMO DIARIO U$" & rs("maximo")
				End If
				sRecomendacion = rs("recomendacion")
				sRequerimiento = rs("requerimiento")
			End If
			Set rs = Nothing
			
			'MsgBox sMensaje
		End If
		
		Cargar
		If Not bExtranjero Then	CalcularTarifa			
			
	Case afxAccionMonedaPago
			Cargar
			CalcularTarifaUltimoGiro
			'CalcularTarifa
			
	Case afxAccionMonto		
			Cargar
			If Not bExtranjero Then	CalcularTarifa			

	Case afxAccionTarifa
			Cargar
			CalcularTarifaCobrada
						
	Case Else
			If Not CargarUltimoEnvio() Then
				If bExtranjero Then
					sPais = Session("PaisMatriz")
					sCiudad = Session("CiudadMatriz")
					sPagador = Session("CodigoMatriz")
					nDDIPais = ObtenerDDI(1, sPais)
					nDDICiudad = ObtenerDDI(2, sCiudad)
				End If
				sMoneda = Session("MonedaExtranjera")
			Else
				bGiroAnterior = True
			End If
				
	End Select
	bGiroNacional = (TRIM(sPais) = Session("PaisMatriz") And TRIM(Request.Form("cbxPais")) = Session("PaisMatriz"))		
	If sPais = Empty Then sPais = Request("Pais")
	If sPais = "CL" Then
		bGiroNacional = True
		sPagador = "AF"
		sMoneda = "CLP"
	End If
	
	If Trim(sMoneda) = "" Then 
		If bGiroNacional Then
			sMoneda = Session("MonedaNacional")
		Else
			sMoneda = Session("MonedaExtranjera")
		End If
	End If
	If bExtranjero Then sPagador = Session("CodigoMatriz")
	If bGiroNacional Then sPagador = Session("CodigoMatriz")
	If sMoneda = Session("MonedaNacional") Then
		nDec = 0
		sColorMoneda = "DodgerBlue"
	Else
		nDec = 2
		sColorMoneda = "#4dc087"
	End If

	
	Function CargarUltimoEnvio()
		Dim rsUG
		
		CargarUltimoEnvio = False		
		If Trim(Request.Form("txtExpress")) = "" Then Exit Function
		Set rsUG = ObtenerUltimosGiros(afxGirosEnviados, Request.Form("txtExpress"), "", 1)
		If rsUG Is Nothing Then Exit Function
		If Not rsUG.EOF Then
			sNombreB = MayMin(rsUG("nombre_beneficiario"))
			sApellidoB = MayMin(rsUG("apellido_beneficiario"))
			sDireccionB = MayMin(rsUG("direccion_beneficiario"))
			sFonoB = rsUG("fono_beneficiario")
			'nMonto = cCur(0 & rsUG("monto_giro"))
			sPais = TRIM(rsUG("pais_beneficiario"))
			sCiudad = rsUG("ciudad_beneficiario")
			sPagador = rsUG("agente_pagador")
			sMoneda = rsUG("codigo_moneda")
			nDDIPais = ObtenerDDI(1, sPais)
			nDDICiudad = ObtenerDDI(2, sCiudad)
			'CalcularTarifaUltimoGiro
			CargarUltimoEnvio = True
		End If
		Set rsUG = Nothing
	End Function
		
	Sub Cargar
		sPais = Request("PaisB")
		sCiudad = Request("CiudadB")
		if sPais = "" then
			sPais = Request.Form("cbxPaisB")
			sCiudad = Request.Form("cbxCiudadB")
		end if
		sPagador = Trim(Request("APagador"))
		If sPagador = "" then
			If Request.Form("optMG") <> "on" Then
				sPagador = Request.Form("cbxPagador")
			Else
				sPagador = Session("CodigoMGEnvio")
			End If
		End if
		'sMoneda = Request("MonedaPago")
		sMoneda = Request("MonedaGiro")
		if sMoneda = "" then
			'sMoneda = Request.Form("cbxMonedaPago")
			sMoneda = Request.Form("cbxMonedaGiro")
		end if
		
		If sPais = "CL" Then
			bGiroNacional = True
			sPagador = "AF"
			sMoneda = "CLP"
		End If
		
		nMonto = cDbl(0 & Request.Form("txtMonto"))		
		nDDIPais = ObtenerDDI(1, sPais)
		nDDICiudad = ObtenerDDI(2, sCiudad)
		sApellidos = Request.Form("txtApellidos")

		sNombreB = Request.Form("txtNombreB")
		sApellidoB = Request.Form("txtApellidoB")
		sDireccionB = Request.Form("txtDireccionB")
		sFonoB = Request.Form("txtFonoB")
		'CalcularTarifa
	End Sub

	Sub CalcularTarifa
		' giro brasil
		If sPais = "BR" And (sPagador <> "ME" And sPagador <> Empty) Then sPagador = "AF"
		
		If sPais <> "" And sCiudad <> "" And sPagador <> "" And sMoneda <> "" And nMonto <> 0 Then
			'Response.redirect "../compartido/error.asp?description=" & Session("CodigoAgente") & ", " & sPagador & ", " & sPais & ", " & sCiudad & ", " & "USD" & ", " & sMoneda  & ", " &  _
			'								nMonto & ", " & nTarifa
			ObtenerTarifaGiros Session("CodigoAgente"), sPagador, sPais, sCiudad, sMoneda, sMoneda, _
											nMonto, nTarifa, nGastoTransfer, nComisionCaptador, nComisionPagador, _
											nComisionMatriz, nAfectoIva			
			nTarifaCobrada = nTarifa
		End If
				
	End Sub

	Sub CalcularTarifaCobrada
		If sPais <> "" And sCiudad <> "" And sPagador <> "" And sMoneda <> "" And nMonto <> 0 Then
			'Response.redirect "../compartido/error.asp?description=" & Session("CodigoAgente") & ", " & sPagador & ", " & sPais & ", " & sCiudad & ", " & "USD" & ", " & sMoneda  & ", " &  _
			'								nMonto & ", " & nTarifa
			ObtenerTarifaGiros Session("CodigoAgente"), sPagador, sPais, sCiudad, sMoneda, sMoneda, _
											nMonto, nTarifa, nGastoTransfer, nComisionCaptador, nComisionPagador, _
											nComisionMatriz, nAfectoIva			
		End If
		nTarifaCobrada = CDbl(0 & Request.Form("txtTarifaCobrada"))
	End Sub

	Sub CalcularTarifaUltimoGiro
		If sPais <> "" And sCiudad <> "" And sPagador <> "" And sMoneda <> "" And nMonto <> 0 Then								
			'Response.redirect "../compartido/error.asp?description=" & Session("CodigoAgente") & ", " & sPagador & ", " & sPais & ", " & sCiudad & ", " & "USD" & ", " & sMoneda  & ", " &  _
			'								nMonto & ", " & nTarifa
			ObtenerTarifaUltimoGiro Session("CodigoAgente"), sPagador, sPais, sCiudad, "USD", sMoneda, _
											nMonto, nTarifa, nGastoTransfer, nComisionCaptador, nComisionPagador, _
											nComisionMatriz, nAfectoIva			
			nTarifaCobrada = nTarifa
		End If
				
	End Sub	
%>
<!--#INCLUDE virtual="/Compartido/Rutinas.asp" -->
<!--#INCLUDE virtual="/Compartido/Errores.asp" -->
<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link rel="stylesheet" type="text/css" href="../Estilos/Cliente.css">
<link href="../CSS/Links_T1T2T3.css" rel="stylesheet" type="text/css">
<link href="../CSS/Linksnuevos.css" rel="stylesheet" type="text/css">
</head>
<script LANGUAGE="VBScript">
<!--

	Sub window_onLoad()
		Dim sCadena	
		
		If Not CalcularTotal Then Exit Sub
		ActivarFoco		
		
		If "<%=sPais%>" <> "CL" Then
			<%If sPagador = Session("CodigoMGEnvio") Then %>				
				frmGiro.optMG.value = "on"				
				frmGiro.optMG.checked = true
				frmGiro.optOtroAgt.value = ""
				frmGiro.optOtroAgt.checked = false
				frmGiro.cbxPagador.disabled = true
				'frmGiro.txtInvoiceMG.disabled = false
			
				frmGiro.imgRedMoneyGram.style.display  = ""
				frmGiro.imgAgtPagador.style.display  = "none"
				frmGiro.cbxPagador.disabled = True
			<%Else%>
				frmGiro.optMG.value = ""
				frmGiro.optMG.checked = false
				frmGiro.optOtroAgt.value = "on"
				frmGiro.optOtroAgt.checked = true
				'frmGiro.cbxPagador.disabled = false
				'frmGiro.txtInvoiceMG.disabled = true
				frmGiro.imgAgtPagador.style.display = ""
				frmGiro.imgRedMoneyGram.style.display  = "none"
			<% End If %>
			<%If Request.form("optDolares") <> "on" Then %>
				frmGiro.optPesos.checked = true
				optPesos_onClick()
			<%Else%>
				frmGiro.optDolares.checked = true
			<%End If%>
			
		Else
			frmGiro.cbxPaisB.disabled = True
		End If
		'msgbox frmGiro.optMG.value & ", " & frmGiro.optOtroAgt.value
		If "<%=bCargarMensaje%>" Then
			If frmGiro.cbxPagador.value <> Empty Then
				'sCadena = Split("<%=sMensaje%>", ";")
				'MsgBox vbCrLf & sCadena(0) & vbCrLf & sCadena(1), vbInformation, "Recordatorio"
				MsgBox "<%=sNota%>" & vbCrLf & "<%=sMaximo%>" & vbCrLf & _
					   "<%=sRecomendacion%>" & vbCrLf & "<%=sRequerimiento%>", vbInformation, "Recordatorio"
			End If
		End If		
		
		frmGiro.txtPaisB.value = frmGiro.cbxPaisB.options(frmGiro.cbxPaisB.selectedIndex).text
		If frmGiro.cbxCiudadB.selectedIndex > -1 Then frmGiro.txtCiudadB.value = frmGiro.cbxCiudadB.options(frmGiro.cbxCiudadB.selectedIndex).text
		
		If "<%=sPais%>" <> "CL" Then
			if frmGiro.optMG.checked Then
				frmGiro.txtPagador.value = "MONEYGRAM"
			else
				if frmGiro.cbxPagador.selectedIndex > - 1 then frmGiro.txtPagador.value = frmGiro.cbxPagador.options(frmGiro.cbxPagador.selectedIndex).text
			end if
		Else
			frmGiro.txtPagador.value = "AFEX"
		End If
	End Sub
	
	Sub ActivarFoco()
		<% If bGiroAnterior Then %>
				frmgiro.txtMonto.focus
				frmGiro.txtMonto.select
				Exit Sub
		<% End If %>
		
		<% Select Case nAccion %>
		<%	Case afxAccionPais %>
				frmGiro.cbxCiudadB.focus 
		<%	Case afxAccionCiudad %>
				frmGiro.txtfonoB.select
		<%	Case afxAccionPagador %>
				frmGiro.txtMonto.select
		<%	Case afxAccionMonto %>
				
		<%	Case afxAccionTarifa %>
				
		<% Case Else %>
				<% If Session("Categoria") = 4 Then %>
						frmGiro.txtNombres.select
				<% Else %>
						frmGiro.txtNombreB.select
				<% End If %>
		<% End Select %>		
	End Sub 

	Sub txtMonto_onBlur()
		Dim nPos
		
		If cCur(0 & Trim(frmGiro.txtMonto.value)) = 0 Then
			frmGiro.txtMonto.value = "<%=FormatNumber(0, nDec)%>"
			frmGiro.txtTarifaSugerida.value = "<%=FormatNumber(0, nDec)%>"
'			frmGiro.txtTarifaCobrada.value = "<%=FormatNumber(0, nDec)%>"
			frmGiro.txtTotal.value = "<%=FormatNumber(0, nDec)%>"
			frmGiro.txtMontoPesos.value = Empty
			Exit Sub
		Else
			nPos = Instr(frmGiro.txtMonto.value, ",")
			If nPos > 0 Then
				If cCur(0 & Mid(frmGiro.txtMonto.value, nPos)) > 0 Then
					msgbox "El monto del giro no puede incluir decimales"
					frmGiro.txtMonto.select
					Exit Sub
				End If
			End If
		End If
		
		'frmGiro.txtMonto.value = FormatNumber(frmGiro.txtMonto.value, 2)
		'If cCur(frmGiro.txtMonto.value) <> cCur(0 & "<%=nMonto%>") Then
		if "<%=sPais%>" <> "CL" Then
			if frmGiro.optPesos.checked Then
				frmGiro.txtMontoPesos.value = formatnumber(round(ccur(frmGiro.txtTotal.value) * ccur(frmGiro.txtTipoCambio.value), 0), 0)
			End If
		end if
		<% If Not bExtranjero Then %>
			HabilitarControles
			frmGiro.action = "EnviarGiro.asp?Accion=<%=afxAccionMonto%>"
			frmGiro.submit 
			frmGiro.action = ""
		<% End IF %>
	End Sub		

	Sub imgCalcular_onClick()
		CalcularTotal
	End Sub

	Function CalcularTotal()
		Dim nCobrada, nMonto
		
		CalcularTotal=False
		frmGiro.txtTarifaSugerida.value = "<%=FormatNumber(nTarifa, nDec)%>"
		nMonto = cDbl(0 & frmGiro.txtMonto.value)
		'nCobrada = cDbl(0 & frmGiro.txtTarifaCobrada.value)
		frmGiro.txtTotal.value = "<%=FormatNumber(nMonto + nTarifaCobrada, nDec)%>"
	'	<%	If Not bGiroNacional Then
	'			If nTarifaCobrada = 0 And nMonto > 0 Then %>
	'				MsgBox "La tarifa cobrada debe ser mayor que cero", ,"AFEX"
	'				Exit Function
	'		<%	ElseIf nTarifaCobrada < nGastoTransfer Then %>
	'				MsgBox "La tarifa cobrada <%=nTarifaCobrada%> no debe ser menor que los gastos de transferencia <%=nGastoTransfer%>", ,"AFEX"
	'				Exit Function
	'	<%		End If 
	'		End If	%>
		CalcularTotal = True
	End Function
 
	Sub cbxPaisB_onblur()
		Dim sCiudad
		
		If frmGiro.cbxPaisB.value = "" Then Exit Sub
		If frmGiro.cbxPaisB.value = "<%=sPais%>" Then Exit Sub
		HabilitarControles
		frmGiro.txtMonto.value = "<%=FormatNumber(0, nDec)%>"
		frmGiro.cbxCiudadB.value = ""
		frmGiro.cbxPagador.value = ""
		frmGiro.cbxMonedaPago.value = ""
		frmGiro.action = "EnviarGiro.asp?Accion=<%=afxAccionPais%>"
		frmGiro.submit 
		frmGiro.action = ""
	End Sub
	
	Sub cbxCiudadB_onblur()
		Dim sComuna
		
		If frmGiro.cbxCiudadB.value = "" Then Exit Sub
		If frmGiro.cbxCiudadB.value = "<%=sCiudad%>" Then Exit Sub		
		HabilitarControles
		frmGiro.txtMonto.value = "<%=FormatNumber(0, nDec)%>"
		
		If "<%=sPais%>" <> "CL" Then
			frmGiro.cbxPagador.value = ""
			frmGiro.cbxMonedaPago.value = ""
		End If
		
		frmGiro.action = "EnviarGiro.asp?Accion=<%=afxAccionCiudad%>"
		frmGiro.submit 
		frmGiro.action = ""
	End Sub

	Sub cbxPagador_onblur()
		Dim sComuna, sCadena
				
		If frmGiro.cbxPagador.value = "" Then Exit Sub
		If frmGiro.cbxPagador.value = "<%=sPagador%>" Then Exit Sub
		
		HabilitarControles
		frmGiro.cbxMonedaPago.value = ""
		frmGiro.action = "EnviarGiro.asp?Accion=<%=afxAccionPagador%>"
		frmGiro.submit 
		frmGiro.action = ""
	End Sub

 
	Sub cbxMonedaPago_onblur()
		Dim sComuna
		
		If frmGiro.cbxMonedaPago.value = "" Then Exit Sub
		If frmGiro.cbxMonedaPago.value = "<%=sMoneda%>" Then Exit Sub		
		HabilitarControles
		frmGiro.action = "EnviarGiro.asp?Accion=<%=afxAccionMonedaPago%>"
		frmGiro.submit 
		frmGiro.action = ""
	End Sub
	
	Sub imgAceptar_OnClick()		
		
		If Not CajaPregunta("AFEX En Linea", "Está seguro que desea enviar el giro?") Then
			Exit Sub
		End If
		
		If Not CalcularTotal Then
			Exit Sub
		End If
		
		'Validaciones 
		If Not ValidarDatos Then
			Exit Sub
		Else
			' si el giro es para brasil solicita otros datos
			'If trim(frmGiro.cbxPaisB.value) = "BR" Then
			'	frmGiro.txtGiroBrasil.value = window.showModalDialog("datosgirobrasil.asp?Dolares=" & frmGiro.txtMonto.value)				
			'	If frmGiro.txtGiroBrasil.value = Empty Then
			'		msgbox "Debe ingresar los datos antes solicitados."					
			'	End If				
			'End IF
				
		
			HabilitarControles
		
			frmGiro.action = "EnviarMail.asp?Accion=3"		
			frmGiro.submit 
			frmGiro.action = ""
		End If	
		
	End Sub

	Function ValidarDatos()
		
		ValidarDatos = False		
		If Trim(frmGiro.txtNombreB.value) = "" Then
			MsgBox "Debe ingresar el nombre del beneficiario",,"AFEX"
			Exit Function
		End If
		If Trim(frmGiro.txtApellidoB.value) = "" Then
			MsgBox "Debe ingresar apellidos del beneficiario",,"AFEX"
			Exit Function
		End If
		If Trim(frmGiro.cbxPaisB.value) = "" Then
			MsgBox "Debe ingresar el pais del beneficiario",,"AFEX"
			Exit Function
		End If
		If Trim(frmGiro.cbxCiudadB.value) = "" Then
			MsgBox "Debe ingresar la ciudad del beneficiario",,"AFEX"
			Exit Function
		End If		
		
		If frmGiro.cbxPaisB.value = "CL" Then
			If Trim(frmGiro.txtFonoB.value) = "" Then
				MsgBox "Debe ingresar el teléfono del beneficiario",,"AFEX"
				Exit Function
			End If
			If frmGiro.cbxCiudadB.value = "SCL" Then
				If Len(Trim(frmGiro.txtfonoB.value)) <> 7 Then
					Msgbox "El número de teléfono para Santiago de Chile debe ser de siete dígitos",, "AFEX"
					Exit Function
				End If
			Else
				If Len(Trim(frmGiro.txtfonoB.value)) <> 6 Then
					Msgbox "El número de teléfono para Regiones de Chile debe ser de seis dígitos",, "AFEX"
					Exit Function
				End If			
			End If
		else
			if not frmGiro.optMG.checked then
				If Trim(frmGiro.txtFonoB.value) = "" Then
					MsgBox "Debe ingresar el teléfono del beneficiario",,"AFEX"
					Exit Function
				End If
			end if
		End If
		
		If "<%=sPais%>" <> "CL" Then							
			If frmGiro.optMG.value <> "on" Then
				If Trim(frmGiro.cbxPagador.value) = "" Then
					MsgBox "Debe seleccionar el agente pagador",,"AFEX"
					Exit Function
				End If
			End If		
			If Trim(frmGiro.cbxMonedaPago.value) = "" Then
				MsgBox "Debe seleccionar la moneda de pago",,"AFEX"
				Exit Function
			End If
		End If
		
		If cCur(0 & Trim(frmGiro.txtMonto.value)) = 0 Then
			MsgBox "Debe ingresar el monto del giro",,"AFEX"
			Exit Function
		End If
		<% If Not bGiroNacional Then %>
			'If cCur(0 & frmGiro.txtTarifaCobrada.value)  < cCur(0 & "<%=nGastoTransfer%>") Then
			'	MsgBox "La tarifa cobrada no debe ser menor que " & FormatNumber("<%=nGastoTransfer%>", 2), ,"AFEX"
			'	Exit Function
			'End If
		<% End If %>
		
		ValidarDatos = True
	End Function
	
	Sub cbxMonedaGiro_onBlur()
		Dim sColor
		
		If frmGiro.cbxMonedaGiro.value = "<%=Session("MonedaNacional")%>" Then
			sColor = "DodgerBlue"
		Else		
			sColor = "#4dc087"
		End If
		frmGiro.cbxMonedaGiro.Style.backgroundColor = sColor
		frmGiro.txtMonto.Style.backgroundColor = sColor
		frmGiro.txtTarifaSugerida.Style.backgroundColor = sColor
	'	frmGiro.txtTarifaCobrada.Style.backgroundColor = sColor
		frmGiro.txtTotal.Style.backgroundColor = sColor		
	End Sub

	Sub UltimosGiros()
		Dim sString, aGiro, sNombre, sCliente
	
		If Trim(frmGiro.txtExpress.value) = "" Then
			sCliente = Trim(frmGiro.txtExchange.value)
		Else
			sCliente = Trim(frmGiro.txtExpress.value)
		End If
		
'		If frmGiro.optPersona.value = "on" Then
			sNombre = "<%=Session("NombreCliente")%>" 'Trim(Trim(frmGiro.txtnombres.value) & " " & Trim(frmGiro.txtApellidoP.value) & " " & Trim(frmGiro.txtApellidoM.value))
'		Else
'			sNombre = Trim(Trim(frmGiro.txtRazonSocial.value))
'		End If
		
		sString = Empty
		sString = window.showModalDialog("../Compartido/UltimosGiros.asp?CodigoCliente=" & sCliente & "&NombreCliente=" & sNombre & "&CodigoMoneda=" & "<%=sMoneda%>" & "&TipoGiro=<%=afxListaGirosEnviados%>")
		If sString <> Empty Then
			' crea un arreglo con los datos de la transferencia seleccionada
			aGiro = Split(sString, ";", 13)
			
			' asigna los datos al envio
			window.frmGiro.txtNombreB.value = aGiro(0)
			window.frmGiro.txtApellidoB.value = aGiro(1)
			window.frmGiro.txtDireccionB.value = aGiro(2)
			window.frmGiro.txtPaisFonoB.value = aGiro(5)
			window.frmGiro.txtAreaFonoB.value = aGiro(6)
			window.frmGiro.txtFonoB.value = aGiro(7)
			window.frmGiro.txtMonto.value = aGiro(8)
			
			If "<%=sPais%>" <> "CL" Then
				frmGiro.cbxMonedaPago.value = aGiro(9)
				frmGiro.cbxMonedaGiro.value = aGiro(9)
				frmGiro.cbxPagador.value = ""
			End If
			
			HabilitarControles
			frmGiro.cbxCiudadB.value = ""
			
			frmGiro.action = "EnviarGiro.asp?Accion=<%=afxAccionMonedaPago%>&PaisB=" & aGiro(3) & "&CiudadB=" & aGiro(4) & _
																		   "&APagador=" & aGiro(9) & "&MonedaPago=" & aGiro(10)
			frmGiro.submit 
			frmGiro.action = ""
		End If		
	End Sub
	
	Sub optMG_onClick()
		'frmGiro.optMG.checked = true
		frmGiro.optOtroAgt.checked = false
		frmGiro.cbxPagador.disabled = true		
		HabilitarControles
		frmGiro.cbxMonedaPago.value = ""
		frmGiro.action = "EnviarGiro.asp?Accion=<%=afxAccionPagador%>&APagador=<%=Session("CodigoMGEnvio")%>"
		frmGiro.submit 
		frmGiro.action = ""		
	End Sub
	
	Sub optOtroAgt_onClick()
		frmGiro.optMG.checked = false
		'frmGiro.optOtroAgt.checked = true
		frmGiro.cbxPagador.disabled = false		
		HabilitarControles
		frmGiro.cbxMonedaPago.value = ""
		frmGiro.action = "EnviarGiro.asp?Accion=<%=afxAccionPagador%>"
		frmGiro.submit 
		frmGiro.action = ""		
	End Sub
	
	Sub optPesos_onClick()
		frmGiro.optDolares.checked = false
		sValor = window.showModalDialog("Tarifas.asp?Tipo=2")
		if Len(sValor) > 10 Then
			msgbox sValor,,"EnviarGiro"
		
		else
			frmGiro.txtTipoCambio.value = formatnumber(round(sValor, 0), 0)
			frmGiro.txtMontoPesos.value = formatnumber(round(ccur(frmGiro.txtTipoCambio.value) * ccur(frmGiro.txtTotal.value), 0), 0)
		end if		
	End Sub
	Sub optDolares_onClick()
		frmGiro.optPesos.checked = false
		frmGiro.txtTipoCambio.value = Empty
		frmGiro.txtMontoPesos.value = Empty
	End Sub	
	
-->
</script>

<body>
<!--<marquee STYLE="HEIGHT: 400; LEFT: 4px; POSITION: absolute; TOP: 16px; WIDTH: 573px" BEHAVIOR="slide" DIRECTION="up" SCROLLAMOUNT="2000" SCROLLDELAY="1">-->
<form id="frmGiro" method="post">
	<input type="hidden" name="txtExchange" value="<%=Request.Form("txtExchange")%>">
	<input type="hidden" name="txtExpress" value="<%=Session("AFEXpress")%>">
	<input type="hidden" name="optPersona" value="<%=Request.Form("optPersona")%>">
	<input type="hidden" name="optEmpresa" value="<%=Request.Form("optEmpresa")%>">
	<input type="hidden" name="txtApellidoP" value="<%=Request.Form("txtApellidoP")%>">
	<input type="hidden" name="txtApellidoM" value="<%=Request.Form("txtApellidoM")%>">
	<input type="hidden" name="txtDireccion" value="<%=Request.Form("txtDireccion")%>">
	<input type="hidden" name="cbxComuna" value="<%=Request.Form("cbxComuna")%>">
	<input type="hidden" name="cbxCiudad" value="<%=Request.Form("cbxCiudad")%>">
	<input type="hidden" name="cbxPais" value="<%=Trim(Request.Form("cbxPais"))%>">
	<input type="hidden" name="txtPaisFono" value="<%=Request.Form("txtPaisFono")%>">
	<input type="hidden" name="txtAreaFono" value="<%=Request.Form("txtAreaFono")%>">
	<input type="hidden" name="txtFono" value="<%=Request.Form("txtFono")%>">
	<input type="hidden" name="txtRut" value="<%=Request.Form("txtRut")%>">
	<input type="hidden" name="txtPasaporte" value="<%=Request.Form("txtPasaporte")%>">
	<input type="hidden" name="cbxPaisPasaporte" value="<%=Request.Form("cbxPaisPasaporte")%>">
	<input type="hidden" name="txtGasto" value="<%=nGastoTransfer%>">
	<input type="hidden" name="txtComisionCaptador" value="<%=nComisionCaptador%>">
	<input type="hidden" name="txtComisionPagador" value="<%=nComisionPagador%>">
	<input type="hidden" name="txtComisionMatriz" value="<%=nComisionMatriz%>">
	<input type="hidden" name="txtAfectoIva" value="<%=nAfectoIva%>">
	<input type="hidden" name="txtPaisB" value="<%=Request.Form("txtPaisB")%>">
	<input type="hidden" name="txtCiudadB" value="<%=Request.Form("txtCiudadB")%>">
	<input type="hidden" name="txtPagador" value="<%=Request.Form("txtPagador")%>">


	<!-- Paso 1 -->
	<table class="borde" border="1" ID="tabPaso1" cellspacing="0" cellpadding="1" style="position: relative; top: 10; left: 6;">		
		<tr>
		  <td bgcolor="#31514A" style="cursor: hand" onClick="UltimosGiros"><img src="../Img/titulos_virtual_Giros.jpg" width="130" height="16"></td>
	  </tr>
		<tr>
			<td style="cursor: hand" onClick="UltimosGiros"><img border="0" id="imgTransferencia" src="../images/Transferencia.jpg" style="cursor: hand" WIDTH="19" HEIGHT="22" onClick="UltimosGiros"><span class="textoempresa">Ultimos
		    Giros</span></td>			
		</tr>
		
		<tr HEIGHT="10"><td colspan="5" bgcolor="#CCCCCC" class="textoempresa"><strong>Datos del Beneficiario</strong></td>
		</tr>
		<tr><td colspan="5">
		<table cellspacing="0" cellpadding="1">
		<tr HEIGHT="15">
			<td></td>
			<td colspan="2"><span class="textoempresa">Nombres</span><br>
		    <input NAME="txtNombreB" class="Borde_tabla_abajo" style="HEIGHT: 22px; WIDTH: 300px" onKeyPress="IngresarTexto(2)" value="<%=sNombreB%>" SIZE="25" onBlurs="window.frmGiro.txtNombreB.value=MayMin(Trim(window.frmGiro.txtNombreB.value))"></td>
			<td colspan="2"><span class="textoempresa">Apellidos</span><br>
		    <input NAME="txtApellidoB" class="Borde_tabla_abajo" style="HEIGHT: 22px; WIDTH: 200px" onKeyPress="IngresarTexto(2)" value="<%=sApellidoB%>" SIZE="25" onBlurs="window.frmGiro.txtApellidoB.value=MayMin(Trim(window.frmGiro.txtApellidoB.value))"></td>
		</tr>
		</table>
		</td></tr>
		<tr><td colspan="5">
		<table border="0" cellspacing="0" cellpadding="1">
		<tr HEIGHT="15">
			<td></td>
			<td COLSPAN="2"><span class="textoempresa">Dirección</span><br>
			<input NAME="txtDireccionB" class="Borde_tabla_abajo" style="HEIGHT: 22px; WIDTH: 280px" onKeyPress="IngresarTexto(3)" value="<%=sDireccionB%>" SIZE="50" onBlurs="window.frmGiro.txtDireccionB.value=MayMin(Trim(window.frmGiro.txtDireccionB.value))"></td>
			<td><span class="textoempresa">Pa&iacute;s</span><br>
				<select name="cbxPaisB" class="textoempresa" style="width: 120px">
					<%	
						CargarUbicacion 1, "", sPais 	
					%>
				</select>			</td>
			<td colspan="1"><span class="textoempresa">Ciudad</span><br>
				<select name="cbxCiudadB" class="textoempresa" style="width: 160px">
					<%	
						CargarCiudadesPais sPais, sCiudad 
					%>
				</select>			</td>
		</tr>
		</table>
		</td></tr>
		<tr><td colspan="5">
		<table border="0" cellspacing="0" cellpadding="1">
		<tr HEIGHT="15">
			<td width="8"></td>
			<td width="218" colspan="1"><span class="textoempresa">Teléfono</span><br>
				<input name="txtPaisFonoB" disabled class="Borde_tabla_abajo" style="width: 40px" value="<%=nDDIPais%>">
				<input name="txtAreaFonoB" disabled class="Borde_tabla_abajo" style="width: 40px" value="<%=nDDICiudad%>">
				<input name="txtFonoB" class="Borde_tabla_abajo" style="width: 90px" onKeyPress="IngresarTexto(1)" value="<%=sFonoB%>" size="10" maxlength="10">			</td>
			<td width="383"><span class="textoempresa">Mensaje al Beneficiario</span><br>
		    <input name="txtMensajeB" class="Borde_tabla_abajo" style="font-family: verdana; font-size: 9pt; HEIGHT: 21px; WIDTH: 380px" onKeyPress="IngresarTexto(3)" value="<%=Request.Form("txtMensajeB")%>" SIZE="255" onBlurs="window.frmGiro.txtMensajeB.value=MayMin(Trim(window.frmGiro.txtMensajeB.value))">			</td>
		</tr>
		</table>
		</td></tr>
		<tr HEIGHT="15">
			<td colspan="5" bgcolor="#CCCCCC"><span class="textoempresa"><strong>Datos del Agente Pagador</strong></span></td>
		</tr>
		<tr HEIGHT="15">
		  <td colspan="5"><%If sPais = "CL" Then%>		    <img border="0" src="../images/Transferencia.jpg" style="cursor: hand; display: " WIDTH="19" HEIGHT="22" onClick="window.open 'http:../Sucursales.asp?Region=13'">
	        <%End If%></td>
	  </tr>
		<tr><td colspan="5">
		<table cellspacing="1" cellpadding="1" id="tbPagador">
			<tr HEIGHT="15" style="display: <%=sDisplay%>">
			<%If sPais <> "CL" Then%>
				<td></td>
				<td colspan=2>
				<%	
					Dim sDisabled
				
					If bGiroNacional Then sDisabled = "disabled"								
				%>				
					<input type="radio" name="optOtroAgt" checked  style="border: 0" <%=sDisabled%>>
					<span class="textoempresa">Red AFEX
					</input>
					</span><img border="0" name="imgAgtPagador" src="../images/Transferencia.jpg" style="cursor: hand; display: " WIDTH="19" HEIGHT="22" onClick="window.open 'http:../agente/ListaAgtPagador.asp?pa=<%=sPais%>&ci=<%=sCiudad%>'">
					&nbsp;&nbsp;&nbsp;
					<input type="radio" name="optMG" style="border: 0" <%=sDisabled%>>
					<span class="textoempresa">Red MoneyGram
					</input>
					</span>					<img border="0" name="imgRedMoneyGram" src="../images/Transferencia.jpg" style="cursor: hand; display: " WIDTH="19" HEIGHT="22" onClick="window.open 'http://www.moneygram.com/servlet/DefaultApplyXSL?xslURL=/Display/temgweb.xsl&URL=/Apps/AgentLocator/AgentSearch.jsp'">
					<br>
				
					<select name="cbxPagador" class="Borde_tabla_abajo" style="width: 250px">				
				
					<%	If sCiudad <> "" Then
							If sPagador <> Session("CodigoMGEnvio") Then
								CargarAgentePagador sPais, sCiudad, sPagador, 1
							Else
								CargarAgentePagador sPais, sCiudad, "", 1
							End If
						End If
					%>
					</select>				</td>
			
				<td valign="bottom"><span class="textoempresa">Moneda</span><br>
				<% If bGiroNacional Then %>
						<select name="cbxMonedaGiro" class="Borde_tabla_abajo" style="width: 130px; background-color: <%=sColorMoneda%>; color: white; font-weight: bold">
				<% Else %>
						<select name="cbxMonedaGiro" style="width: 130px; background-color: <%=sColorMoneda%>; color: white; font-weight: bold" disabled>
				<% End If %>
				<%	
					'response.Write sPagador & ", " &  sPais & ", " & sCiudad & ", " & sMoneda
					If sPagador <> "" Then
						CargarMonedaGiro sPagador, sPais, sCiudad, sMoneda
					End If
				%>
				  </select>				</td>			
				<td valign="bottom" style="display: none"><span class="textoempresa">Moneda de Pago</span><br>
					<select name="cbxMonedaPago" disabled class="Borde_tabla_abajo" style="width: 130px">
					<%	
					'response.Write sPagador & ", " &  sPais & ", " & sCiudad & ", " & sMoneda
					If sPagador <> "" Then
						CargarMonedaGiro sPagador, sPais, sCiudad, sMoneda
					End If
					%>
				  </select>				</td>			
			<%End If%>
			</tr>
		<tr HEIGHT="10">
			<td></td>
			<td><span class="textoempresa">Monto</span><br>
				<input NAME="txtMonto" class="Borde_tabla_abajo" STYLE="HEIGHT: 22px; TEXT-ALIGN: right; WIDTH: 120px; background-color: <%=sColorMoneda%>; color: white; font-weight: bold" onKeyPress="IngresarTexto(1)" value="<%=FormatNumber(nMonto, nDec)%>" SIZE="15">
				<img style="display: <%=sDisplay%>" border="0" height="22" id="imgCalcular" name="imgCalcular" onmouseover="imgCalcular.style.cursor ='Hand'" src="../images/BotonFlechaDerecha.jpg" style="LEFT: 0px; POSITION: relative; TOP: 5px" width="21">			</td>
			<td style="display:"><span class="textoempresa">Tarifa</span><br>
				<input NAME="txtTarifaSugerida" disabled class="Borde_tabla_abajo" STYLE="HEIGHT: 22px; TEXT-ALIGN: right; WIDTH: 120px; background-color: <%=sColorMoneda%>; color: white; font-weight: bold" SIZE="15">
&nbsp;&nbsp;&nbsp;&nbsp;			</td>			
			<td style="display: <%=sDisplay%>"><span class="textoempresa">Total</span><br>
				<input NAME="txtTotal" disabled class="Borde_tabla_abajo" STYLE="HEIGHT: 22px; TEXT-ALIGN: right; WIDTH: 120px; background-color: <%=sColorMoneda%>; color: white; font-weight: bold" SIZE="15">			</td>
		</tr>
		<%If sPais <> "CL" Then%>
		<tr HEIGHT="10">
			<td></td>
			
			<td><span class="textoempresa"><strong>Depósito para AFEX</strong></span><br>
				<input type="radio" name="optPesos">
				<span class="textoempresa">Pesos</span>
				</input><input type="radio" name="optDolares">
				<span class="textoempresa">D&oacute;lares</span>				</input>			</td>
			<td style="display:"><span class="textoempresa">T. Cambio</span><br>
				<input NAME="txtTipoCambio" disabled class="Borde_tabla_abajo" STYLE="HEIGHT: 22px; TEXT-ALIGN: right; WIDTH: 120px; background-color: <%=sColorMoneda%>; color: white; font-weight: bold" value="<%=Request.Form("txtTipoCambio")%>" SIZE="15">			</td>
			<td style="display:"><span class="textoempresa">Monto Pesos</span><br>
				<input NAME="txtMontoPesos" disabled class="Borde_tabla_abajo" STYLE="HEIGHT: 22px; TEXT-ALIGN: right; WIDTH: 120px; background-color: <%=sColorMoneda%>; color: white; font-weight: bold" value="<%=Request.Form("txtMontoPesos")%>" SIZE="15">			</td>			
		</tr>
		<%End If%>
		</table>
		</td></tr>
		
		<tr HEIGHT="0">
		<td><img border="0" id="imgAceptar" onclick src="../images/BotonAceptar.jpg" style="LEFT: 461px; POSITION: relative; TOP: 0px; cursor: hand" WIDTH="70" HEIGHT="20"></td>
		</tr>
	</table>	
	
	<input type="hidden" name="txtGiroBrasil" value="">
	
	
</form>
</body>
<!--#INCLUDE virtual="/Compartido/Rutinas.htm" -->
</html>
