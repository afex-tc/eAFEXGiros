<%@ Language=VBScript %>

<%
	'Se asegura que la página no se almacene en la memoria cache
	Response.Expires = 0
	If Session("CodigoCliente") = "" Then
		response.Redirect "http:../Compartido/TimeOut.htm"
		response.end
	End If
%>
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
	Dim bCargarMensaje, sww
	
	sww=Request.QueryString("sw")
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
		nMonto = cDbl(0 & Request.Form("txtMonto"))		
		nDDIPais = ObtenerDDI(1, sPais)
		nDDICiudad = ObtenerDDI(2, sCiudad)		

		'CalcularTarifa
	End Sub

	Sub CalcularTarifa
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
</head>
<script LANGUAGE="VBScript">
<!--
	dim sEncabezadoFondo, sEncabezadoTitulo
	Dim rsOficina, afxWeb, sColor1, sColor2
	sColor1 = "#F0F0F0"
	sColor2 = "#F6F6F6"
	
	sEncabezadoFondo = "Consultas"
	sEncabezadoTitulo = "Oficinas de Pago"
	
	Sub window_onLoad()
		Dim sCadena
		
		If Not CalcularTotal Then Exit Sub
		ActivarFoco	
		'msgbox frmLugarPago.optMG.value & ", " & frmLugarPago.optOtroAgt.value
		If "<%=bCargarMensaje%>" Then
			If frmLugarPago.cbxPagador.value <> Empty Then
				'sCadena = Split("<%=sMensaje%>", ";")
				'MsgBox vbCrLf & sCadena(0) & vbCrLf & sCadena(1), vbInformation, "Recordatorio"
				MsgBox "<%=sNota%>" & vbCrLf & "<%=sMaximo%>" & vbCrLf & _
					   "<%=sRecomendacion%>" & vbCrLf & "<%=sRequerimiento%>", vbInformation, "Recordatorio"
			End If
		End If
		frmLugarPago.cbxCiudadB.disabled =true
		frmLugarPago.cbxPagador.disabled =true
		'frmLugarPago.cbxPaisB.focus		
	End Sub
	
	Sub ActivarFoco()		
		
		<% Select Case nAccion %>
		<%	Case afxAccionPais %>
				frmLugarPago.cbxCiudadB.focus 
		<%	Case afxAccionCiudad %>
				frmLugarPago.txtfonoB.select
		<%	Case afxAccionPagador %>
				frmLugarPago.txtMonto.select		
		<%	Case afxAccionTarifa %>
				<% If bExtranjero Or request.Form("optMG") = "on" Then %>
						frmLugarPago.txtInvoiceMG.select				
				<% Else %>
						frmLugarPago.txtBoleta.select
				<% End If %>
		<% Case Else %>
				<% If Session("Categoria") = 4 Then %>
						frmLugarPago.txtNombres.select
				<% Else %>
						frmLugarPago.txtNombreB.select
				<% End If %>
		<% End Select %>		
	End Sub 
	
	Sub cbxPaisB_onblur()
		Dim sCiudad		
		If frmLugarPago.cbxPaisB.value = "" Then Exit Sub
		If frmLugarPago.cbxPaisB.value = "<%=sPais%>" Then Exit Sub
		HabilitarControles		
		frmLugarPago.cbxCiudadB.disabled =false
		frmLugarPago.action = "LugaresPago.asp?Accion=<%=afxAccionPais%>"
		frmLugarPago.submit 
		frmLugarPago.action = ""
	End Sub	

	Sub cbxCiudadB_onblur()
		Dim sComuna
		
		If frmLugarPago.cbxCiudadB.value = "" Then Exit Sub
		If frmLugarPago.cbxCiudadB.value = "<%=sCiudad%>" Then Exit Sub		
		HabilitarControles	
		frmLugarPago.cbxPagador.disabled=false	
		frmLugarPago.action = "LugaresPago.asp?Accion=<%=afxAccionCiudad%>"
		frmLugarPago.submit 
		frmLugarPago.action = ""
	End Sub
	Sub cbxCiudadB_onclick()
		Dim sComuna
		
		If frmLugarPago.cbxCiudadB.value = "" Then Exit Sub
		If frmLugarPago.cbxCiudadB.value = "<%=sCiudad%>" Then Exit Sub		
		HabilitarControles	
		frmLugarPago.cbxPagador.disabled=false	
		frmLugarPago.action = "LugaresPago.asp?Accion=<%=afxAccionCiudad%>"
		frmLugarPago.submit 
		frmLugarPago.action = ""
	End Sub	
	Sub cbxPagador_onblur()
		Dim sComuna, sCadena
		frmLugarPago.imgAgtPagador.focus		
		If frmLugarPago.cbxPagador.value = "" Then Exit Sub
		If frmLugarPago.cbxPagador.value = "<%=sPagador%>" Then Exit Sub
			
		HabilitarControles	
		frmLugarPago.cbxPagador.disabled=false	
		frmLugarPago.action = "LugaresPago.asp?Accion=<%=afxAccionPagador%>&sw=1"
		frmLugarPago.submit 
		frmLugarPago.action = ""
	End Sub
	Sub cbxPagador_onclick()
		Dim sComuna, sCadena
		frmLugarPago.imgAgtPagador.focus		
		If frmLugarPago.cbxPagador.value = "" Then Exit Sub
		If frmLugarPago.cbxPagador.value = "<%=sPagador%>" Then Exit Sub
			
		HabilitarControles	
		frmLugarPago.cbxPagador.disabled=false	
		frmLugarPago.action = "LugaresPago.asp?Accion=<%=afxAccionPagador%>&sw=1"
		frmLugarPago.submit 
		frmLugarPago.action = ""
	End Sub
	
	Function ValidarDatos()
		
	ValidarDatos = False		
					
		If frmLugarPago.optMG.value = "on" Then			
			If Len(Trim(frmLugarPago.txtInvoiceMG.value)) <> 8 Then
				MsgBox "El invoice debe ser de 8 dígitos", ,"AFEX"
				Exit Function
			End If				
		End If
		ValidarDatos = True
	End Function

	
-->
</script>

<!--#INCLUDE virtual="/Compartido/Rutinas.asp" -->
<!--#INCLUDE virtual="/Compartido/Errores.asp" -->
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link rel="stylesheet" type="text/css" href="../Estilos/Cliente.css">
<link rel="stylesheet" type="text/css" href="../Estilos/Reportes.css"></HEAD>
<BODY>
<!-- #INCLUDE virtual="/Compartido/Encabezado.htm" -->
<form id="frmLugarPago" method="post">
	<input type="hidden" name="cbxCiudad1" value="<%=Request.Form("cbxCiudadB")%>">
	<input type="hidden" name="cbxPais1" value="<%=Trim(Request.Form("cbxPaisB"))%>">
	<input type="hidden" name="cbxpagador1" value="<%=spagador%>">
	
	<table border="0" align="center" cellspacing="0" cellpadding="1">
		<tr HEIGHT="15">
			<td></td>			
			<td>Pais<br>
				<select name="cbxPaisB" style="width: 120px">
					<%	
						CargarUbicacion 1, "", sPais 	
					%>
				</select>
			</td>
			<td colspan="1">Ciudad<br>
				<select name="cbxCiudadB" style="width: 160px" >				
					<%	
						CargarCiudadesPais sPais, sCiudad 
					%>
				</select>
			</td>
		</tr>
		</table>
		<table align="center" cellspacing="1" cellpadding="1" id="tbPagador" border="0" > 
		<tr HEIGHT="15" style="display: <%=sDisplay%>">
			<td></td>
			<td colspan=2>Agente Pagador<br>
			<%	
				Dim sDisabled 
				If bGiroNacional Then sDisabled = "disabled"
			%>
			<select name="cbxPagador" style="width: 250px" >
				
				<%	If sCiudad <> "" Then
						If sPagador <> Session("CodigoMGEnvio") Then
							CargarAgentePagador sPais, sCiudad, sPagador, 0
						Else
							CargarAgentePagador sPais, sCiudad, "", 0
						End If
					End If
					
				%>
				</select>	
			</td>			
			</tr>
		</table>
		<table width="500" align="center">
			<tr>
			<td align="center" width="500" height="60">
				<img border="0" id="imgAgtPagador" align="center"  src="../images/pago.ico" style="cursor: hand" onclick="window.open'http:OficinasPago.asp?codpag=<%=sPagador%>&paisb=<%=sPais%>&ciudadb=<%=sCiudad%>','_self'"><br>
			</td>		
			</tr>
			
		</table>	
</form>
</BODY>
<!--#INCLUDE virtual="/Compartido/Rutinas.htm" -->
</HTML>
