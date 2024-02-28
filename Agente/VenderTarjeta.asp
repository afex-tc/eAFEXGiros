<%@ Language=VBScript %>

<%
	'Se asegura que la página no se almacene en la memoria cache
	Response.Expires = 0
	If Session("CodigoAgente") = "" Then
		response.Redirect "../Compartido/TimeOut.htm"
		response.end
	End If
%>
<!--#INCLUDE virtual="/agente/Constantes.asp" -->
<!--#INCLUDE virtual="/Compartido/Rutinas.htm" -->
<%
	Dim sMoneda, sVigencia, cObservado
	Dim nMonto, nParidad, nEquivalente, nTarifaSugerida, nTarifaCobrada
	Dim nGastoExtranjera, nGastoNacional, nTotal, sPais
	Dim sHabilitado, nBoletaServicio, nPIN
	
	
	On Error Resume Next

	sw =Request.QueryString("sw")
    nmn=request.QueryString("nmn")
	sHabilitado = Request("hb")
	sMoneda = Request("mn")
	'response.Write mto1
	nMonto = cCur(0 & Request("mto"))
	nBoletaServicio = Request("bs")
	cObservado = cCur(0 & Request("tc"))
	nTotal = 0	
	'response.Write sw
	if sw="" then	
		sw=-1
	elseif sw=3 then
		nMonto= Request.QueryString("mto")
	elseif sw = 1 then				
		response.Redirect "http:existe.asp?sw=1"	
	elseif sw = 11 then				
		response.Redirect "http:stock_tarjetas.asp?sw=11"		
	elseif sw=0 then	   	
		nMonto= Request.QueryString("mto")
		nmn=request.QueryString("nmn")		
		 response.Redirect "http:NoExiste.asp?nmn=" & nmn & "&mn=" & Smoneda & "&mto=" & nMonto & "&bs=" & nBoletaServicio & "&sw=0"
	end if
	
	If sMoneda = "" Then
		sMoneda = Session("MonedaNacional")
	End If

	If Trim() <> Empty Then
		nPIN = Request("np")
	End If
	
	If cObservado <> 0 Then
		nTotal = FormatNumber(cDbl(0 & nMonto) * cObservado, 0)
	Else
		nTotal = nMonto
	End If
	
	
	Sub CargarMoneda()
		Dim sSelect
		
		'If frmTarjeta.cbxMoneda.value = "" Then
			If sMoneda = "USD" Then
			'	Response.write "<option SELECTED value=USD>Dolar Americano</option>"
				Response.write "<option value=CLP>Peso Chileno</option>"
			Else
			'	Response.write "<option value=USD>Dolar Americano</option>"
				Response.write "<option SELECTED value=CLP>Peso Chileno</option>"
			End If
		'End If
	End Sub

	Sub CargarMonto()
		If Trim(sMoneda) = "USD" Then
			If Trim(nMonto) = 5 Then
				'Response.write "<option SELECTED value=5>5</option>"
				'Response.write "<option value=10>10</option>"
			ElseIf Trim(nMonto) = 10 Then
				'Response.write "<option value=5>5</option>"
				'Response.write "<option SELECTED value=10>10</option>"
			Else
			'	Response.write "<option value=5>5</option>"
			'	Response.write "<option value=10>10</option>"
			End If
		Else
			If Trim(nMonto) = 1000 Then
				Response.Write "<option SELECTED value=1000>" & FormatNumber(1000, 0) & "</option>"
				Response.write "<option value=2000>" & FormatNumber(2000, 0) & "</option>"
				Response.write "<option value=5000>" & FormatNumber(5000, 0) & "</option>"
			ElseIf Trim(nMonto) = 2000 Then
				Response.Write "<option value=1000>" & FormatNumber(1000, 0) & "</option>"
				Response.write "<option SELECTED value=2000>" & FormatNumber(2000, 0) & "</option>"
				Response.Write "<option value=5000>" & FormatNumber(5000, 0) & "</option>"
			ElseIf Trim(nMonto) = 5000 Then
				Response.Write "<option value=1000>" & FormatNumber(1000, 0) & "</option>"
				Response.write "<option value=2000>" & FormatNumber(2000, 0) & "</option>"
				Response.Write "<option SELECTED value=5000>" & FormatNumber(5000, 0) & "</option>"
			Else
				Response.Write "<option value=1000>" & FormatNumber(1000, 0) & "</option>"
				Response.write "<option value=2000>" & FormatNumber(2000, 0) & "</option>"
				Response.write "<option value=5000>" & FormatNumber(5000, 0) & "</option>"
			End If
		End If
	End Sub

%>
<!--#INCLUDE virtual="/Compartido/Transferencias.asp" -->
<!--#INCLUDE virtual="/Compartido/Rutinas.asp" -->
<!--#INCLUDE virtual="/Compartido/Errores.asp" -->
<html>
<head>
<META name=VI60_defaultClientScript content=VBScript>
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
	sEncabezadoTitulo = "Venta Tarjeta Telefónica"

	'Sub txtTarifaCobrada_OnBlur()
	'	If Trim(frmCheque.txtTarifaCobrada.value) = "" Then
	'		frmCheque.txtTarifaCobrada.value = "0,00"
	'	Else
	'		frmCheque.txtTarifaCobrada.value = FormatNumber(frmCheque.txtTarifaCobrada.value, 2)
	'	End If
	'	CalcularMontos
	'End Sub 
	
	'Sub txtParidad_onBlur()
	'	If Trim(frmCheque.txtParidad.value) = "" Then
	'		frmCheque.txtParidad.value = "0,00"
	'	Else
	'		frmCheque.txtParidad.value = FormatNumber(frmCheque.txtParidad.value, 8)
	'	End If		
	'	CalcularUSD
	'	CalcularMontos
	'End Sub

	'Sub imgCalcular_onClick()
	'	CalcularUSD
	'	CalcularMontos
	'End Sub
	
	
	Sub imgAceptar_onCLick()
	Dim cont
	Dim j
	Dim rs1
	Dim numero 
	Dim sSQL3
	Dim rs3
	Dim Cnn
	
		
	numero =frmTarjeta.txtNumeroBoleta.value		
	
	if  (frmTarjeta.txtNumeroBoleta.value) = "" then
		MsgBox "Debe ingresar  N° Boleta", ,"AFEX"
		frmTarjeta.txtNumeroBoleta.focus
		exit sub
	else
		cont = (Len(Trim(frmTarjeta.txtNumeroBoleta.value)))
			j=0
				for i=1 to cont
					if (isnumeric(Mid(Trim(frmTarjeta.txtNumeroBoleta.value), i,i))= false ) or ((asc(Mid(Trim(frmTarjeta.txtNumeroBoleta.value), i,i))=45) or (asc(Mid(Trim(frmTarjeta.txtNumeroBoleta.value), i,i))=43))then  
					   j=j+1 					
 					end if 	 							
 				next		
		
 			    if j>0 then 
 				  MsgBox "Debe ingresar solo Números en N° Boleta", ,"AFEX"
 				  frmTarjeta.txtNumeroBoleta.focus
 				  frmTarjeta.txtNumeroBoleta.value="" 				  			  
 				  Exit sub
 				end if  
 	end if	
 
		'if "<%=sw%>" <> empty then
			'frmTarjeta.action = "http:"grabarVentaTarjeta.asp"?nmn= " & frmTarjeta.cbxMoneda(frmTarjeta.cbxMoneda.selectedIndex).Text 
			'frmTarjeta.submit 
		   ' frmTarjeta.action = ""	
 		'else 			
			frmTarjeta.action = "http:Buscar_Boleta.asp?nmn=" & frmTarjeta.cbxMoneda(frmTarjeta.cbxMoneda.selectedIndex).Text & " &mn=" & frmTarjeta.cbxMoneda.value & " &mto=" & frmTarjeta.cbxMonto.value & " &bs=" & frmTarjeta.txtnumeroboleta.value 
		    frmTarjeta.submit 
		    frmTarjeta.action = ""
		'end if
	End Sub

	'Objetivo:	calcular los montos de la página
	'Sub CalcularMontos()
	'
	'	
	'	' si la tarifa cobrada es distinta de la sugerida no se calcula
	'	If frmCheque.txtTarifaCobrada.value = frmCheque.txtTarifaSugerida.value Then
	'		frmCheque.txtTarifaCobrada.value = "5,00"
	'	End If
	'	
	'	' calcula la tarifa
	'	frmCheque.txtTarifaSugerida.value = "5,00"
	'	' calcula la comisión
	'	frmCheque.txtGastoExtranjera.value = "3,00"
	'	
	'	frmCheque.txtGastoNacional.value = FormatNumber(cDbl(0 & frmCheque.txtGastoExtranjera.value) * _
	'											   cCur(0 & "<%=cObservado%>"), 0)
	'	' calcula el total
	'	frmCheque.txtTotal.value = CalcularTotal
	'End Sub	

	'Sub cbxMoneda_onKeyPress()
	'	cbxMoneda_onClick		
	'End Sub
	
	'Sub cbxMoneda_onKeyDown()
	'	cbxMoneda_onClick
	'End Sub

	'Sub cbxMoneda_onKeyUp()
	'	cbxMoneda_onClick
	'End Sub

	'Sub cbxMoneda_onClick()
	'	If frmCheque.cbxRate.value = "1" Then
	'		frmCheque.txtParidad.value  = FormatNumber( _											
	'						frmCheque.cbxParidades(frmCheque.cbxMoneda.selectedIndex).text _
	'						, 8)
	'	End If
	'	CalcularUSD
	'	
	'	' calcula los montos
	'	CalcularMontos
	'End Sub

	'Sub cbxRate_onKeyDown()
	'	cbxRate_onClick
	'End Sub

	'Sub cbxRate_onKeyUp()
	'	cbxRate_onClick
	'End Sub


	'Sub cbxRate_onClick()
	'	If frmCheque.cbxRate.value = "1" Then
	'		frmCheque.txtParidad.value  = FormatNumber( _											
	'						frmCheque.cbxParidades(frmCheque.cbxMoneda.selectedIndex).text _
	'						, 8)
	'		frmCheque.txtParidad.disabled = True
	'	Else
	'		frmCheque.txtParidad.disabled = False
	'	End If
	'	CalcularUSD
	'	CalcularMontos		
	'End Sub
	
	'Sub CalcularUSD
	'	'Calcula el Equivalente en Dolares
	'	frmCheque.txtEquivalente.value = FormatNumber( _
	'			Round(CDbl(0 & frmCheque.txtMonto.value) * CDbl(0 & frmCheque.txtParidad.value), 7), 2)
	'End Sub

	'Objetivo:	calcular el total de la transferencia
	'Function CalcularTotal()
	'	CalcularTotal = 0
	'
	'	' verifica si hay monto
	'	If frmCheque.txtMonto.value = Empty Or frmCheque.txtMonto.value = "0,00" Then Exit Function
	'	
	'	CalcularTotal = FormatNumber(cDbl(0 & frmCheque.txtTarifaCobrada.value) + _
	'								cDbl(0 & frmCheque.txtEquivalente.value), 2)
	'End Function	


	'Function ValidarDatos()
	'	ValidarDatos = False
	'	If cCur(0 & frmCheque.txtMonto.value) <= 0 Then
	'		MsgBox "Debe ingresar un monto válido para el cheque",,"AFEX En Linea"
	'		frmCheque.txtMonto.select
	'		frmCheque.txtMonto.focus
	'		Exit Function
	'	End If
	'	ValidarDatos = True
	'End Function

	Sub cbxMoneda_onblur()
		If frmTarjeta.cbxMoneda.value = "" Then Exit Sub
		If frmTarjeta.cbxMoneda.value = "<%=sMoneda%>" Then Exit Sub
		
		frmTarjeta.action = "http:VenderTarjeta.asp?mn=" & frmTarjeta.cbxMoneda.value
		
		frmTarjeta.submit 
		frmTarjeta.action = ""
	End Sub

	'Sub cbxMonto_onblur()
	'	If frmTarjeta.cbxMonto.value = "" Then Exit Sub
	'	If frmTarjeta.cbxMonto.value = "<%=nMonto%>" Then Exit Sub
	'End Sub

	Sub CargarMenu()
		Dim sId

		frmTarjeta.objmenu.bgColor = document.bgColor 
		frmTarjeta.objmenu.stylesheet = "../Estilos/Cliente.css"
		sId = frmTarjeta.objmenu.addparent("Opciones")

		If "<%=nBoletaServicio%>" <> Empty Then
			frmTarjeta.objMenu.addchild sId, "Boleta de Servicios", "Servicios", "Principal"
			'frmTarjeta.objMenu.addchild sId, "Ver Boleta", "BS", "Principal"
			frmTarjeta.objMenu.addchild sId, "", "", ""
			frmTarjeta.objMenu.addchild sId, "", "", ""
		Else
			frmTarjeta.objMenu.addchild sId, "", "", ""
			frmTarjeta.objMenu.addchild sId, "", "", ""
			frmTarjeta.objMenu.addchild sId, "", "", ""
		End If
	End Sub

	sub window_onload()
		CargarMenu
	'	cbxMoneda_onClick
	'	frmCheque.txtmonto.select
	'	frmCheque.txtmonto.focus
	end sub

-->
</script>

<body>
<!-- #INCLUDE virtual="/Compartido/Encabezado.htm" -->
<form id="frmTarjeta" method=post>
<!-- Paso 1 -->
<table ID="tabPaso1" CELLSPACING="0" CELLPADDING="0" BORDER="0" HEIGHT="100" WIDTH="560" STYLE="BORDER-BOTTOM: steelblue 2px solid; BORDER-LEFT: steelblue 2px solid; BORDER-RIGHT: steelblue 2px solid; BORDER-TOP: steelblue 2px solid;  HEIGHT: 80px; LEFT: 4px; POSITION: relative; TOP: 0px; WIDTH: 466px">
	<tr><td HEIGHT="1"></td></tr>
	<tr><td class="titulo">Valores</td></tr>
	<tr><td>
		<table ID="tabInformacion">
			<tr>
				<td width="135">Moneda<br>
					<!--<input NAME="txtMoneda" SIZE="20" style="HEIGHT: 22px; WIDTH: 110px">-->
					<select id="cbxMoneda" name="cbxMoneda" style="width: 130px" <%=sHabilitado%>>
						<% CargarMoneda %>						
					</select>
				</td>
				<td>Monto<br>
					<!--<input NAME="txtMonto" SIZE="25"  style="HEIGHT: 22px; WIDTH: 140px">-->
					<select id="cbxMonto" name="cbxMonto" style="width: 70px" <%=sHabilitado%>>
						<% CargarMonto %>
					</select>
				</td>
			</tr>
		</table>		
		<table>
			<tr>
			</tr>
		</table>
	</td></tr>
	<tr><td class="titulo">Información</td></tr>
	<tr><td>
		<table ID="tabInformacion">
			<tr>
				<td width="54">Nº Boleta<br>
					<input id="txtNumeroBoleta" NAME="txtNumeroBoleta"  style="HEIGHT: 22px; WIDTH: 130px" value="<%=nBoletaServicio%>" <%=sHabilitado%> maxlength="7">
					
				</td>
				<td>Nº PIN<br>
					<input NAME="txtNumeroPIN" style="HEIGHT: 22px; WIDTH: 100px" value="<%=nPIN%>" disabled>
				</td>
				<td width="160">
				</td>
			</tr>
			<tr>
				<td colspan="3" width="420">
					
					<% If Trim(sHabilitado) = Empty Then %>
						<img align="right" alt border="0" hspace="0" id="imgAceptar" src="../images/BotonAceptar.jpg" style="cursor: hand;" WIDTH="70" HEIGHT="20">
					<% End If %>
				</td>
			</tr>
			<tr>
			</tr>
		</table>		
		<table>
			<tr>
			</tr>
		</table>
	</td></tr>
</table>
<table BORDER="0" cellspacing="0" cellpadding="0" STYLE="LEFT: 337px; POSITION: absolute; TOP: 50px">	
	<tr><td>
	    <object align="left" id="objMenu" style="HEIGHT: 60px; LEFT: 0px; POSITION: relative; TOP: 0px; WIDTH: 160px" type="text/x-scriptlet" width="170" VIEWASTEXT border="0" valign="top"><param NAME="Scrollbar" VALUE="0"><param NAME="URL" VALUE="http:../Scriptlets/Menu.htm"></object>
	</td></tr>
</table>
</form>
</body>

<OBJECT id=Printer1 style="LEFT: 0px; TOP: 0px" codebase="AfexPrinter.CAB#version=1,0,0,0"
	classid="clsid:210A8E07-6FF9-4C4E-A664-5844036C0E33" VIEWASTEXT>
	<PARAM NAME="_ExtentX" VALUE="1">
	<PARAM NAME="_ExtentY" VALUE="1">
</OBJECT>

<SCRIPT LANGUAGE=vbscript>
<!--
	Sub objMenu_OnScriptletEvent(strEventName, varEventData)
	   Select Case strEventName
			Case "linkClick"
				If Right(varEventData, 9) = "Servicios" Then
					If Not CajaPregunta("AFEX En Linea", "Coloque la boleta en la impresora TM y haga click en Aceptar") Then
						Exit Sub
					End If
					<% If Session("ModoPrueba") Then %>
							ImprimirBoletaServicios
					<% Else %>
							ImprimirBoletaServicios
					<% End If %> 
					window.navigate "AtencionClientes.asp"
					window.close 

				ElseIf Right(varEventData, 2) = "BS" Then
					'MostrarBS
				End If		
		End Select		
	End Sub
	
	Sub ImprimirBoletaServicios()
		Dim afxTM
		
		Set afxTM = CreateObject("AfexPrinter.TM295")
		
		On Error Resume Next
		
		afxTM.TMInicializar 
		If afxTM.ErrNumber <> 0 Then
			MostrarErrorAFEX afxTM, "Imprimir Boleta 1"	   
		End If
		
		afxTM.TMHabilitar2
	    If afxTM.ErrNumber <> 0 Then
			MostrarErrorAFEX afxTM, "Imprimir Boleta 2"	   
		End If
		
		afxTM.TMEncabezadoBoletaBS "<%=nBoletaServicio%>", Date, Empty, Time 
		
		If afxTM.ErrNumber <> 0 Then
			MostrarErrorAFEX afxTM, "Imprimir Boleta 3"	   
		End If
		
		afxTM.TMDetalleBoletaBS "<%=sMoneda%>", _
								cCur(0 & "<%=nMonto%>"),cCur(0 & "<%=cObservado%>"), _
								cCur(0 & "<%=nTotal%>"), 4, "<%=nPIN%>"

		If afxTM.ErrNumber <> 0 Then
			MostrarErrorAFEX afxTM, "Imprimir Boleta 4"	   
		End If

		'If cCur(0 & "<%=request.Form("txtSubTotal1")%>") <> 0 Then
		'	afxTM.DetalleBoleta MSComm1, "<%=request.Form("cbxMoneda1")%>", _
		'			cCur(0 & "<%=request.Form("txtMonto1")%>"), cCur(0 & "<%=request.Form("txtTipoCambio1")%>"), cCur(0 & "<%=request.Form("txtSubTotal1")%>")
		'	If afxTM.ErrNumber <> 0 Then
		'		MostrarErrorAFEX afxTM, "Imprimir Boleta 5"	   
		'	End If
		'End If
		
		'If cCur(0 & "<%=request.Form("txtSubTotal2")%>") <> 0 Then
		'	afxTM.DetalleBoleta MSComm1, "<%=request.Form("cbxMoneda2")%>", _
		'			cCur(0 & "<%=request.Form("txtMonto2")%>"), cCur(0 & "<%=request.Form("txtTipoCambio2")%>"), cCur(0 & "<%=request.Form("txtSubTotal2")%>")
		'	If afxTM.ErrNumber <> 0 Then
		'		MostrarErrorAFEX afxTM, "Imprimir Boleta 6"	   
		'	End If
'hasta aki		'End If

		afxTM.TMPieBoletaBS cCur(0 & "<%=nTotal%>")
		
		If afxTM.ErrNumber <> 0 Then
			MostrarErrorAFEX afxTM, "Imprimir Boleta 5"	   
		End If
		afxTM.TMDeshabilitar2
				
		Set afxTM = Nothing
		
	End Sub

	Sub MostrarBS()
		'Dim sSolucion, aSolucion, sDetalle, nDec

		'txtDireccionR.value = Replace(txtDireccionR.value, "#", "")
		'txtDireccion.value = Replace(txtDireccion.value, "#", "")
		'txtMensaje.value = Replace(txtMensaje.value, "#", "")
		'If frmGiro.txtMonedaPago.value = "<%=Session("MonedaNacional")%>" Then nDec = 0 Else nDec = 2
		'sDetalle = 	"?Nombres=" & txtNombresR.value & _
		'				"&Apellidos=" & txtApellidosR.value & _
		'				"&Direccion=" & txtDireccionR.value & _
		'				"&AreaFono=" & txtAreaFonoR.value & _
		'				"&PaisFono=" & txtPaisFonoR.value & _
		'				"&Fono=" & txtFonoR.value & _
		'				"&Ciudad=" & Trim(txtCiudadR.value) & _
		'				"&Rut=" & txtRutR.value & _
		'				"&Codigo=" & txtGiro.value & _
		'				"&sPIN=" & "<%=sNumeroPIN%>" & _
		'				"&CiudadB=" & txtCiudad.value & _
		'				"&PaisB=" & txtPais.value & _
		'				"&NombresB=" & txtNombres.value & _
		'				"&ApellidosB=" & txtApellidos.value & _
		'				"&DireccionB=" & txtDireccion.value & _
		'				"&FonoB=(" & txtPaisFono.value & txtAreaFono.value &  ") " & txtFono.value & _
		'				"&Mensaje=" & txtMensaje.value & _
		'				"&Monto=" & frmGiro.txtMonedaPago.value & "  " & FormatNumber(cCur(0 & txtMonto.value), nDec) & _
		'				"&Gastos=" & frmGiro.txtMonedaPago.value & "  " & FormatNumber(cCur(0 & trim(txtGastos.value)), nDec) & _
		'				"&Comision=" & frmGiro.txtMonedaPago.value & "  " & FormatNumber(cCur(0 & trim(txtIva.value)), nDec) & _
		'				"&Total=" & frmGiro.txtMonedaPago.value & "  " & Formatnumber(cCur(0 & trim(txtMonto.value)) +  cCur(0 & trim(txtGastos.value)), nDec) & _
		'				"&TotalNacional=" & "<%=Session("MonedaNacional")%> "
		'If frmGiro.txtMonedaPago.value = "<%=Session("MonedaNacional")%>" Then
		'	sDetalle = sDetalle & Formatnumber(ROUND(cCur(0 & txtIva.value), 0), nDec)
		'Else
		'	sDetalle = sDetalle & Formatnumber(ROUND(cCur(0 & txtIva.value) * cCur(0 & txtTipoCambio.value), 0), nDec)
		'	sDetalle = sDetalle & "&TipoCambio=" & frmGiro.txtMonedaPago.value & "  " & FormatNumber(cCur(0 & trim(txtTipoCambio.value)), nDec)
		'End If
		'sSolucion = window.showModalDialog("BoletaServicios.asp" & sDetalle, , "center=yes" ) 				
	End Sub

-->
</SCRIPT>
</html>
