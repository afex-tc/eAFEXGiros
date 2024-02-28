<%@ Language=VBScript %>
<%
	'option explicit	
	'Se asegura que la página no se almacene en la memoria cache
	Response.Expires = 0
	If Session("CodigoCliente") = "" Then
		response.Redirect "../Compartido/TimeOut.htm"
		response.end
	End If
%>
<!--#INCLUDE virtual="/cliente/Constantes.asp" -->
<%
	Dim sMoneda, sVigencia, cObservado
	Dim nMonto, nParidad, nEquivalente, nTarifaSugerida, nTarifaCobrada
	Dim nGastoExtranjera, nGastoNacional, nTotal, sValuta, sFormaPago
	Dim nTarifa, sTarifa, sTotal
	
	On Error Resume Next

	sMoneda = Trim(request.Form("cbxMoneda"))
	sValuta = request.Form("cbxValuta")
	sFormaPago = request.Form("cbxFormaPago")
	If sFormaPago = "" Then
		sFormaPago = afxEfectivoCLP
	End If
	If sMoneda = "" Then
		sMoneda = Session("MonedaExtranjera")
	End If

	sVigencia = ValidarVigencia

	nAccion = cInt(0 & Request("Accion"))
	nTarifa = 0
	CargarTrf
	
	Sub CargarTrf()
		'cObservado = ObtenerObservado()		
		nMonto = cCur(cDbl(0 & Request.Form("txtMonto")))
		If nMonto <> 0 Then
			nTarifa = ObtenerTarifaTransfer(cDbl(Request.Form("txtMonto")), Request.Form("cbxPais"), "***", Request.Form("cbxMoneda"))
			sTarifa = Formatnumber(nTarifa, 2)
			sTotal = Formatnumber(cDbl(Request.Form("txtMonto")) + cDbl(nTarifa), 2)
		End If
	End Sub

%>
<!--#INCLUDE virtual="/Compartido/Transferencias.asp" -->
<!--#INCLUDE virtual="/Compartido/Rutinas.asp" -->
<!--#INCLUDE virtual="/Compartido/Errores.asp" -->
<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link rel="stylesheet" type="text/css" href="../Estilos/Cliente.css">
</head>
<body>
<script LANGUAGE="VBScript">
<!--
	Const sEncabezadoFondo = "Transacciones"
	Const sEncabezadoTitulo = "Enviar una Transferencia"
	Const sClass = "TituloPrincipal"

	Sub CambiarCursor(Byval sControl)

		document.all.item(sControl).style.cursor = "Hand"
	
	End Sub
	
	Sub MostrarPaso(ByVal sPaso)
		window.tabPaso1.style.display = "none"
		window.tabPaso2.style.display = "none"
		window.tabPaso3.style.display = "none"
		If sPaso="tabPaso1" Then
			window.tabPaso1.style.display = ""
		ElseIf sPaso="tabPaso2" Then
			window.tabPaso2.style.display = ""
		Else
			window.tabPaso3.style.display = ""
		End If
	End Sub
	
	Sub MostrarPaso3(ByVal sPaso)
		
		Select Case uCase(sPaso)
		Case "TABPASO2"
			If Not ValidarPaso1() Then Exit Sub
		Case "TABPASO3"
			If Not ValidarPaso2() Then Exit Sub
		End Select
		
		MostrarPaso sPaso
		lblNombre.innerText = trim(frmTransfer.txtNombre.value)
		lblMonto.innerText = trim(frmTransfer.txtMonto.value)
		lblTarifa.innerText = trim(frmTransfer.txtTarifa.value)
		lblTotal.innerText = trim(frmTransfer.txtTotal.value)
		lblMoneda.innerText = trim(frmTransfer.cbxMoneda(frmTransfer.cbxMoneda.selectedIndex).Text)
		lblMoneda2.innerText = trim(frmTransfer.cbxMoneda(frmTransfer.cbxMoneda.selectedIndex).Text)
		lblMoneda3.innerText = trim(frmTransfer.cbxMoneda(frmTransfer.cbxMoneda.selectedIndex).Text)
		lblDireccion.innerText = trim(frmTransfer.txtDireccion.value)
		lblCiudad.innerText = trim(frmTransfer.txtCiudad.value)
		lblBancoDestino.innerText = trim(frmTransfer.txtBancoDestino.value)
		lblCuentaDestino.innerText = trim(frmTransfer.txtCuentaDestino.value)
	End Sub

		
	Sub txtMonto_OnBlur()
		If Trim(frmTransfer.txtMonto.value) = "" Then
			frmTransfer.txtMonto.value = "0,00"
		Else
			frmTransfer.txtMonto.value = FormatNumber(frmTransfer.txtMonto.value, 2)
		End If		
		CalcularMontos
	End Sub 

	Sub tdCalcular_OnClick()
		CalcularMontos
	End Sub
	
	'Objetivo:	calcular los montos de la página
	Sub CalcularMontos()			
		' calcula la tarifa
		frmTransfer.txtTarifa.value = ObtenerTarifaTransfer(frmTransfer.cbxValuta.value, _
														frmTransfer.cbxMoneda.value, frmTransfer.txtMonto.value)																		
		CalcularUSD
		frmTransfer.txtTotal.value = CalcularTotal
		CalcularNacional
		
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
		If frmTransfer.cbxMoneda.value <> "<%=Session("MonedaExtranjera")%>" Then
			tdEquivalenteUSD.style.display = ""
		Else
			tdEquivalenteUSD.style.display = "none"
		End If
		frmTransfer.txtParidad.value  = FormatNumber( _											
						frmTransfer.cbxParidades(frmTransfer.cbxMoneda.selectedIndex).text _
						, 8)
		frmTransfer.txtTipoCambio.value  = FormatNumber( _											
						frmTransfer.cbxTipoCambio(frmTransfer.cbxMoneda.selectedIndex).text _
						, 4)		
		' calcula los montos
		CalcularMontos
	End Sub

	Sub CalcularUSD
		'Calcula el Equivalente en Dolares
		frmTransfer.txtEquivalente.value = FormatNumber( _
				Round(CDbl(0 & frmTransfer.txtMonto.value) * CDbl(0 & frmTransfer.txtParidad.value), 7), 2)
	End Sub

	'Objetivo:	calcular el total de la transferencia
	Function CalcularTotal()
		CalcularTotal = 0
	
		' verifica si hay monto
		If frmTransfer.txtMonto.value = Empty Or frmTransfer.txtMonto.value = "0,00" Then Exit Function
				
		CalcularTotal = FormatNumber(cDbl(0 & frmTransfer.txtTarifa.value) + _
									cDbl(0 & frmTransfer.txtEquivalente.value), 2)
	End Function	

	'Objetivo:	calcular el total de la venta
	Function CalcularNacional()
		
	
		' verifica si hay monto
		If frmTransfer.txtMonto.value = Empty Or frmTransfer.txtMonto.value = "0,00" Then Exit Function
				
		frmTransfer.txtMontoNacional.value  = FormatNumber(Round(cDbl(0 & frmTransfer.txtTotal.value) * _
									cDbl(0 & frmTransfer.txtTipoCambio.value), 0), 0)
	End Function	

	Sub cbxValuta_onClick()
		CalcularMontos 
	End Sub
	
	Sub window_onload()
		cbxMoneda_onClick 
		ActivarFoco
	End Sub	

	Sub tdAceptar_onClick()
		If Not ValidarDatos() Then
			Exit Sub
		End If
		HabilitarControles
		frmTransfer.action = "GrabarEnvioTransfer.asp"
		frmTransfer.submit()
		frmTransfer.action = ""
	End Sub 

	Sub ActivarFoco()
		<% Select Case nAccion %>
		<%	Case afxAccionPais %>
				frmTransfer.cbxCiudad.focus 
		<%	Case afxAccionCiudad, afxAccionMonedaPago %>
				frmTransfer.txtMonto.focus
				frmTransfer.txtMonto.select
		<% End Select %>		
	End Sub 

	Function ValidarDatos()
		ValidarDatos = False
		If Not ValidarPaso1() Then
			Exit Function
		End If
		If Not ValidarPaso2() Then
			Exit Function
		End If
		ValidarDatos = True		
	End Function

	Function ValidarPaso1()
		ValidarPaso1 = False
		If Trim(frmTransfer.cbxMoneda.value) = "" Then
			MsgBox "Debe seleccionar la moneda de envío",,"AFEX"
			Exit Function
		End If
		If cCur(0 & Trim(frmTransfer.txtMonto.value)) = 0 Then
			MsgBox "Debe ingresar el monto de la transferencia",,"AFEX"
			Exit Function
		End If	
		ValidarPaso1 = True
	End Function
	
	Function ValidarPaso2()
		ValidarPaso2 = False
		If Trim(frmTransfer.txtBancoDestino.value) = "" Then
			MsgBox "Debe ingresar el banco donde desea enviar la transferencia",,"AFEX"
			Exit Function
		End If
		If Trim(frmTransfer.txtCuentaDestino.value) = "" Then
			MsgBox "Debe ingresar la cuenta donde desea enviar la transferencia",,"AFEX"
			Exit Function
		End If
		If Trim(frmTransfer.txtNombre.value) = "" Then
			MsgBox "Debe ingresar el nombre de quien recibirá el dinero",,"AFEX"
			Exit Function
		End If
		If Trim(frmTransfer.txtDireccion.value) = "" Then
			MsgBox "Debe ingresar la dirección del banco donde desea enviar la transferencia",,"AFEX"
			Exit Function
		End If
		If Trim(frmTransfer.txtCiudad.value) = "" Then
			MsgBox "Debe ingresar la ciudad del banco donde desea enviar la transferencia",,"AFEX"
			Exit Function
		End If
		If Trim(frmTransfer.txtAba.value) = "" Then
			MsgBox "Debe ingresar el número de ABA de la transferencia",,"AFEX"
			Exit Function
		End If
		ValidarPaso2 = True
	End Function

	Sub UltimasTransferencias()
		Dim sString, aTrf
		Dim sNombre, sCliente
		
		sCliente = Trim("<%=Session("AFEXchange")%>")
		sNombre = "<%=Session("NombreCliente")%>"
		sString = Empty
		'msgbox sCliente & ", " & frmTransfer.cbxMoneda.value
		sString = window.showModalDialog("../Compartido/UltimasTransferencias.asp?CodigoCliente=" & sCliente & _
																	"&NombreCliente=" & sNombre & _
																	"&CodigoMoneda=" & frmTransfer.cbxMoneda.value)
		If sString <> Empty Then
			' crea un arreglo con los datos de la transferencia seleccionada
			aTrf = Split(sString, ";", 13)
			
			' asigna los datos de la transferencia
			'frmTransfer.txtBancoOrigen.value = aTrf(0)
			'frmTransfer.txtCuentaOrigen.value = aTrf(1)
			frmTransfer.txtBancoDestino.value = aTrf(2)			
			frmTransfer.txtCuentaDestino.value = aTrf(3)
			frmTransfer.txtNombre.value = aTrf(4)
			frmTransfer.txtCiudad.value = aTrf(5)
			frmTransfer.txtABA.value = aTrf(6)
			'frmTransfer.txtInvoice.value = aTrf(7)
			frmTransfer.txtDireccion.value = aTrf(8)
			frmTransfer.txtBancoIntermedio.value = aTrf(9)
			frmTransfer.txtCuentaIntermedio.value = aTrf(10)
			frmTransfer.txtCiudadIntermedio.value = aTrf(11)			
			frmTransfer.txtDireccionIntermedio.value = aTrf(12)
		End If		
	End Sub

	Sub cbxFormaPago_onKeyPress()
		cbxFormaPago_onClick		
	End Sub
	
	Sub cbxFormaPago_onKeyDown()
		cbxFormaPago_onClick
	End Sub

	Sub cbxFormaPago_onKeyUp()
		cbxFormaPago_onClick
	End Sub

	Sub cbxFormaPago_onClick()
		'msgbox frmTransfer.cbxFormaPago.value & ", " & afxEfectivoUSD & afxDepositoUSD
		Select Case frmTransfer.cbxFormaPago.value
		Case <%=afxEfectivoUSD%>, <%=afxDepositoUSD%>, <%=afxCustodiaUSD%>
			tdTipoCambio.style.display = "none"
			tdMontoNacional.style.display = "none"

		Case <%=afxEfectivoCLP%>, <%=afxDepositoCLP%>
			tdTipoCambio.style.display = ""
			tdMontoNacional.style.display = ""
		End Select
		
	End Sub

	Sub VerDemo_OnClick()
		window.open "DemoEnviarTransfer.htm", null, _
				"title=no, top=0, left=0, height=405, width=570, status=no,toolbar=no,menubar=no,location=no"
	End Sub

	Sub Condiciones_OnClick()
		window.open "CondicionesTransfer.htm", null, _
				"title=no, top=0, left=0, height=340, width=540, status=no,toolbar=no,menubar=no,location=no"
	End Sub
		
	Sub optEnviarBoleta_onClick()
		window.frmTransfer.optGuardarBoleta.checked=False		
	End Sub

	Sub optGuardarBoleta_onClick()
		window.frmTransfer.optEnviarBoleta.checked=False		
	End Sub

-->
</script>
<!--#INCLUDE virtual="/Compartido/Encabezado.htm" -->
<table id="tabHagaseCliente" cellspacing="0" cellpadding="0" border="0" style="LEFT: 5px; POSITION: relative">
<tr height="20"><td>
      <object height="20" id="objTab" style="HEIGHT: 20px; WIDTH: 465px" type="text/x-scriptlet" VIEWASTEXT><param NAME="Scrollbar" VALUE="0"><param NAME="URL" VALUE="../Scriptlets/Tab.htm"></object>
</td></tr>
<tr ID="tabPaso1"><td>
<form id="frmTransfer" method="post">
<table cellspacing="0" class="borde" ID="tabPaso11" border="0" style="HEIGHT: 280px; LEFT: 4px; POSITION: relative; TOP: 0px; WIDTH: 530px">
	<tr class="descripcion">
		<td colspan="5"><table><tr>
			<td WIDTH="5"></td>
			<td HEIGHT="50" COLSPAN="2">
  				<br>Seleccione la moneda y el tiempo en que desea el envío, luego ingrese el monto y haga click en la forma de pago o presione la tecla «Tab» para ver el cálculo. Luego haga click en «siguiente» para ir al siguiente paso.<br><br>
			</td>	
			<td colspan="2" id="VerDemo"><img id="VerDemo" src="../images/BotonVerDemo.jpg" style="cursor: hand" WIDTH="80" HEIGHT="20"></td>
		</tr></table>
		</td>
	</tr>
	<tr HEIGHT="15">
		<td colspan="5" class="titulo">Datos del Envío</td>
	</tr>
	<tr HEIGHT="60">
		<td></td>
		<td>En qué moneda desea enviar el dinero?<br>
			<select NAME="cbxMoneda" style="HEIGHT: 22px; WIDTH: 180px" disabled>
			<%
				CargarMonedasTransfer sMoneda
			%>
			</select>							
			<select name="cbxParidades" STYLE="Display: none">	
			<%
				CargarParidadesTransfer afxTCTransferencia, sMoneda
			%>
			</select>
			<select name="cbxTipoCambio" STYLE="Display: none">	
			<%
				CargarParidadesTransfer afxTCVenta, sMoneda
			%>
			</select>
		</td>
		<td>En cuánto tiempo desea el envío?<br>
		<select NAME="cbxValuta" style="HEIGHT: 22px; WIDTH: 106px">
			<%
				CargarValuta sValuta
			%>
		</select>
		</td>			
		<td>Cuánto dinero desea enviar?<br>
			<input STYLE="TEXT-ALIGN: right" ID="txtMonto" NAME="txtMonto" SIZE="15" onKeyPress="IngresarTexto(1)" value="0,00">
		</td>
		<td></td>
	</tr>
	<tr HEIGHT="2">
		<td></td>
	</tr>
	<tr HEIGHT="15">	
		<td></td>
		<td style="display: none">Paridad en <%=Session("MonedaExtranjera")%><br>
			<input STYLE="HEIGHT: 22px; TEXT-ALIGN: right; WIDTH: 93px" NAME="txtParidad" SIZE="15" onKeyPress="IngresarTexto(1)" disabled value="0,00">
		</td>
		<td id="tdEquivalenteUSD" style="display:">Este es el equivalente en <%=Session("MonedaExtranjera")%><br>
			<input STYLE="HEIGHT: 22px; TEXT-ALIGN: right; WIDTH: 103px" NAME="txtEquivalente" SIZE="15" onKeyPress="IngresarTexto(1)" disabled value="0,00">
		</td>
		<td>Esta es la tarifa que debe<br>pagar en <%=Session("MonedaExtranjera")%><br><strong><input NAME="txtTarifa" STYLE="FONT-SIZE: 10pt; HEIGHT: 20px; TEXT-ALIGN: right; WIDTH: 123px" value="<%=sTarifa%>" disabled></strong></td>
		<td>Este es el total que debe<br>pagar en <%=Session("MonedaExtranjera")%><br><strong><input NAME="txtTotal" STYLE="FONT-SIZE: 10pt; HEIGHT: 20px; TEXT-ALIGN: right; WIDTH: 123px" value="<%=sTotal%>" disabled></strong></td>
		<td id="tdCalcular" >
			<img align="absmiddle" src="../images/BotonCalcular.jpg" style="cursor: hand" WIDTH="70" HEIGHT="20">
		<td>
	</tr>	
	<tr HEIGHT="2">
		<td></td>
	</tr>
	<tr HEIGHT="60" ID="tdTarifa" NAME="tdTarifa">
		<td></td>
		<td>En qué forma va a pagar?<br>
		<select NAME="cbxFormaPago" style="HEIGHT: 22px; WIDTH: 120px">
			<%
				CargarFormaPago sFormaPago
			%>
		</select>
		</td>			
		<td id="tdTipoCambio" style="display:">Tipo Cambio<br>
			<input STYLE="TEXT-ALIGN: right" NAME="txtTipoCambio" SIZE="15" onKeyPress="IngresarTexto(1)" value="0,0000" disabled>
		</td>
		<td id="tdMontoNacional" style="display:">Total en Pesos<br>
			<input STYLE="HEIGHT: 22px; TEXT-ALIGN: right; WIDTH: 93px" NAME="txtMontoNacional" SIZE="15" disabled value="0">
		</td>
		<td></td>
	</tr>
	<tr HEIGHT="15">	
		<td></td>
		<td colspan="3"></td>
	</tr>
	<tr>
		<td></td>
		<td COLSPAN="4"></td>
	</tr>
	<tr>
		<td ></td>
		<td colspan="2" style="font-size: 7pt">*Nota: Los valores incluyen IVA</td>
		<td align="right">
		<table cellspacing="5" cellpadding="2">
		<tr>
			<td id="tdsiguiente" onclick="MostrarPaso3('tabPaso2')">
				<img align="right" src="../images/Botonsiguiente.jpg" id="imgsiguiente" style="cursor: hand" WIDTH="80" HEIGHT="20">
			</td>
		</tr>
		</table>
		</td>
	</tr>
</table></td></tr>

<tr ID="tabPaso2" style="DISPLAY: none"><td>
	<!-- Paso 2 -->
<table cellspacing="0" class="borde" ID="tabPaso22" border="0" style="HEIGHT: 280px; LEFT: 4px; POSITION: relative; TOP: 0px; WIDTH: 530px;">
	<tr class="descripcion">
		<td WIDTH="5"></td>
		<td colspan="4">
		<table></tr>
		<td HEIGHT="50" width="450px">
		  <br>Ingrese los datos de la cuenta que recibirá la transferencia. Si desea copiar los datos de una transferencia ya enviada, haga click en Ultimas Transferencias.<br>Luego haga click en «siguiente» para ir al siguiente paso.<br><br>
		</td>	
		<td style="cursor: hand" onClick="UltimasTransferencias">
			<img border="0" id="imgTransferencia" src="../images/Transferencia.jpg" WIDTH="19" HEIGHT="22">
		</td>
		<td style="cursor: hand" onClick="UltimasTransferencias">
			Ver Ultimas Transferencias</td>
		<td colspan="2"></td>
		</tr></table>
		</td>
	</tr>
	<tr ID="trDatosTransfer1"><td colspan="3" class="titulo">Datos de la Transferencia</td></tr>
	<tr ID="trDatosTransfer2" style="DISPLAY: nones">
		<td WIDTH="5"></td>
		<td>
		<table ID="tabBeneficiario" width="100%" border="0">
	   <!--		<tr><td>			<table width="100%">			<tr><td colspan="3" class="titulo" style="BACKGROUND-COLOR: lightgrey">Origen</td></tr>			<tr HEIGHT="15">			<td></td>			<td>Banco <br><input NAME="txtBancoOrigen" SIZE="25" style="HEIGHT: 22px; WIDTH: 301px" disabled></td>			<td>Cuenta Corriente<br><input NAME="txtCuentaOrigen" SIZE="25" style="HEIGHT: 22px; WIDTH: 179px" disabled></td>			</tr>			</table>		</td></tr>		-->
		<tr><td>
			<table width="100%" border="0">		
			<tr HEIGHT="15">
			<td></td>
			<td>Nombre del banco donde desea enviar el dinero?<br><input NAME="txtBancoDestino" SIZE="25" style="HEIGHT: 22px; WIDTH: 299px" onBlur="frmTransfer.txtBancoDestino.value=MayMin(frmTransfer.txtBancoDestino.value)"></td>
			<td>Cuenta Corriente donde desea enviar el dinero?<br><input name="txtCuentaDestino" SIZE="25" style="HEIGHT: 22px; WIDTH: 179px" onBlur="frmTransfer.txtCuentaDestino.value=MayMin(frmTransfer.txtCuentaDestino.value)"></td>
			</tr>
			<tr HEIGHT="15">
			<td></td>
			<td colspan="2">Nombre de quien recibirá el dinero?<br><input name="txtNombre" SIZE="25" style="HEIGHT: 22px; WIDTH: 299px" onBlur="frmTransfer.txtNombre.value=MayMin(frmTransfer.txtNombre.value)"></td>
			</tr>
			<tr HEIGHT="15">
			<td></td>
			<td>Dirección del banco?<br><input NAME="txtDireccion" SIZE="50" style="HEIGHT: 22px; WIDTH: 321px" onBlur="frmTransfer.txtDireccion.value=MayMin(frmTransfer.txtDireccion.value)"></td>
			<td>Ciudad del banco?<br><input NAME="txtCiudad" style="HEIGHT: 22px; WIDTH: 153px" onBlur="frmTransfer.txtCiudad.value=MayMin(frmTransfer.txtCiudad.value)"></td>
			</tr>
			<tr HEIGHT="15">
			<td></td>
			<td>Número de ABA?<br><input NAME="txtAba" SIZE="25" style="HEIGHT: 22px; WIDTH: 126px"></td>
			<td style="display: none">Invoice<br><input ID="txtInvoice" NAME="txt" SIZE="25" style="HEIGHT: 22px; WIDTH: 179px"></td>
			</tr>
			</table>
		</td></tr>
		</table>
	</td></tr>
	<tr>
		<td></td>
		<td colspan="3">
		<table cellspacing="5" cellpadding="2">
		<tr>
			<td id="tdAnterior2" onclick="MostrarPaso3('tabPaso1')">
				<img align="absMiddle" src="../images/BotonAnterior.jpg" style="cursor: hand" WIDTH="80" HEIGHT="20"></td>		
			<td width="100%"></td>
			<td id="tdsiguiente2" onclick="MostrarPaso3('tabPaso3')">
				<img align="absmiddle" src="../images/Botonsiguiente.jpg" id="imgsiguiente" style="cursor: hand" WIDTH="80" HEIGHT="20">
			</td>
		</tr>
		</table>
		</td>
	</tr>	
	<tr ID="trDatosTransfer1" style="display: none; cursor: hand;"><td colspan="3" class="titulo">Intermediario &nbsp;
		<font ID="dspDatosTransfer" FACE="Marlett" SIZE="3">u</font> 
	</td></tr>
	<tr id="trIntermediario" style="display: none">
		<td WIDTH="5"></td>
		<td width="60%">
		<table width="100%" border="0">		
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
	</td></tr>
</table></td></tr>

<tr ID="tabPaso3" style="DISPLAY: none"><td>
<!-- Paso 3 -->	
<table cellspacing="0" class="borde" ID="tabPaso33" border="0" style="HEIGHT: 280px; LEFT: 4px; POSITION: relative; TOP: 0px; WIDTH: 530px">
	<tr class="descripcion">
		<td WIDTH="5"></td>
		<td HEIGHT="50">
			  <p align="justify"><br>Lea atentamente el resumen y las condiciones de su operación. Si está seguro de la información y desea realizar la operación haga click en «aceptar». Si tiene dudas, o algún dato debe ser corregido, vuelva a los pasos anteriores para modificarlo.<br><br></p>
		</td>	
		<td><img id="Condiciones" src="../images/BotonCondiciones.jpg" style="cursor: hand" WIDTH="80" HEIGHT="20"></td>
		</tr>
	<tr HEIGHT="15">
		<td colspan="4" class="titulo">Resumen</td>
	</tr>
	<tr HEIGHT="20" style="font-size: 10pt">
		<td></td>
		<td COLSPAN="2"><p align="justify">
			<br>
			Usted desea enviar una transferencia por <strong ID="lblMonto" NAME="lblMonto"></strong>&nbsp;
			<strong ID="lblMoneda"></strong>&nbsp;
			a <strong ID="lblNombre"></strong>,&nbsp;
			en la cuenta <strong ID="lblCuentaDestino"></strong>,&nbsp;
			del banco <strong ID="lblBancoDestino"></strong>,&nbsp;
			ubicado en <strong ID="lblDireccion"></strong>,&nbsp; 
			<strong ID="lblCiudad" NAME="lblCiudad"></strong>.
			<br>
			La tarifa cobrada es de <strong ID="lblTarifa" NAME="lblTarifa"></strong>&nbsp;
			<strong ID="lblMoneda2"></strong>,&nbsp; 
			y debe cancelar a AFEX la suma total de 
			<strong ID="lblTotal"></strong>&nbsp;
			<strong ID="lblMoneda3"></strong>.<br><br>
		</p></td>
	</tr>
	<tr HEIGHT="20" style="font-size: 10pt">
		<td></td>
		<td COLSPAN="2">
			Si la Información es correcta haga click en Aceptar. <br>
			Si es necesario vuelva a los pasos correspondientes si desea modificar algún dato.<br><br>
		</td>
	</tr>
	<tr>
		<td></td>		
		<td>
			<input type="radio" name="optEnviarBoleta" checked>Enviar boleta a domicilio<br>
			<input type="radio" name="optGuardarBoleta">Guardar boleta en AFEX (6 meses)<br><br>
		</td>
		
	</tr>
	<tr>
		<td></td>
		<td colspan="3">
		<table cellspacing="5" cellpadding="2">
		<tr>
			<td id="tdAnterior3" onclick="MostrarPaso3('tabPaso2')">
				<img src="../images/BotonAnterior.jpg" style="cursor: hand" WIDTH="80" HEIGHT="20">
			</td>		
			<td width="100%"></td>
			<td id="tdAceptar" onclick>
				<img src="../images/BotonAceptar.jpg" style="cursor: hand" WIDTH="70" HEIGHT="20">
			</td>
		</tr>
		</table>
		</td>
	</tr>
</table>
</form>
</td></tr>
</table>
</body>
<!--#INCLUDE virtual="/Compartido/Rutinas.htm" -->
<!--#INCLUDE virtual="/Compartido/Transferencias.htm" -->
</html>
