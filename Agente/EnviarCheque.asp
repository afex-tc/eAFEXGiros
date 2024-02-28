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
<%
	Dim sMoneda, sVigencia, cObservado
	Dim nMonto, nParidad, nEquivalente, nTarifaSugerida, nTarifaCobrada
	Dim nGastoExtranjera, nGastoNacional, nTotal, sPais
	
	On Error Resume Next
	
	sMoneda = Trim(request.Form("cbxMoneda"))
	If sMoneda = "" Then
		sMoneda = Session("MonedaExtranjera")
	End If

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
	sEncabezadoTitulo = "Enviar un Cheque"

	Sub txtMonto_OnBlur()
		If Trim(frmCheque.txtMonto.value) = "" Then
			frmCheque.txtMonto.value = "0,00"
		Else
			frmCheque.txtMonto.value = FormatNumber(frmCheque.txtMonto.value, 2)
		End If		
		'cbxMoneda_onClick
		CalcularUSD
		CalcularMontos
	End Sub 

	Sub txtTarifaCobrada_OnBlur()
		If Trim(frmCheque.txtTarifaCobrada.value) = "" Then
			frmCheque.txtTarifaCobrada.value = "0,00"
		Else
			frmCheque.txtTarifaCobrada.value = FormatNumber(frmCheque.txtTarifaCobrada.value, 2)
		End If
		CalcularMontos
	End Sub 
	
	Sub txtParidad_onBlur()
		If Trim(frmCheque.txtParidad.value) = "" Then
			frmCheque.txtParidad.value = "0,00"
		Else
			frmCheque.txtParidad.value = FormatNumber(frmCheque.txtParidad.value, 8)
		End If		
		CalcularUSD
		CalcularMontos
	End Sub

	Sub imgCalcular_onClick()
		CalcularUSD
		CalcularMontos
	End Sub
	
	
	Sub imgAceptar_onCLick()
	
		If Not CajaPregunta("AFEX En Linea", "Está seguro que desea enviar el cheque?") Then
			Exit Sub
		End If

		If Not ValidarDatos() Then
			Exit Sub
		End If
		HabilitarControles
		frmCheque.action = "GrabarEnvioCheque.asp"
		frmCheque.submit 
		frmCheque.action = ""		
	End Sub

	'Objetivo:	calcular los montos de la página
	Sub CalcularMontos()
	
		
		' si la tarifa cobrada es distinta de la sugerida no se calcula
		If frmCheque.txtTarifaCobrada.value = frmCheque.txtTarifaSugerida.value Then
			frmCheque.txtTarifaCobrada.value = "10,00" 'frmCheque.txtTarifaCobrada.value = "5,00"
		End If
		
		' calcula la tarifa
		frmCheque.txtTarifaSugerida.value = "10,00" 'frmCheque.txtTarifaSugerida.value = "5,00"
		' calcula la comisión
		frmCheque.txtGastoExtranjera.value = "8,00" 'frmCheque.txtGastoExtranjera.value = "3,00"
		
		frmCheque.txtGastoNacional.value = FormatNumber(cDbl(0 & frmCheque.txtGastoExtranjera.value) * _
												   cCur(0 & "<%=cObservado%>"), 0)
		' calcula el total
		frmCheque.txtTotal.value = CalcularTotal
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
	'	If frmCheque.cbxRate.value = "1" Then
	'		frmCheque.txtParidad.value  = FormatNumber( _											
	'						frmCheque.cbxParidades(frmCheque.cbxMoneda.selectedIndex).text _
	'						, 8)
	'	End If
		CalcularUSD
		
		' calcula los montos
		CalcularMontos
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
		If frmCheque.cbxRate.value = "1" Then
			frmCheque.txtParidad.value  = FormatNumber( _											
							frmCheque.cbxParidades(frmCheque.cbxMoneda.selectedIndex).text _
							, 8)
			frmCheque.txtParidad.disabled = True
		Else
			frmCheque.txtParidad.disabled = False
		End If
		CalcularUSD
		CalcularMontos		
	End Sub
	
	Sub CalcularUSD
		'Calcula el Equivalente en Dolares
		frmCheque.txtEquivalente.value = FormatNumber( _
				Round(CDbl(0 & frmCheque.txtMonto.value) * CDbl(0 & frmCheque.txtParidad.value), 7), 2)
	End Sub

	'Objetivo:	calcular el total de la transferencia
	Function CalcularTotal()
		CalcularTotal = 0
	
		' verifica si hay monto
		If frmCheque.txtMonto.value = Empty Or frmCheque.txtMonto.value = "0,00" Then Exit Function
		
		CalcularTotal = FormatNumber(cDbl(0 & frmCheque.txtTarifaCobrada.value) + _
									cDbl(0 & frmCheque.txtEquivalente.value), 2)
	End Function	


	Function ValidarDatos()
		ValidarDatos = False
		If frmCheque.cbxMoneda.value = empty Then
			MsgBox "Debe seleccionar una moneda para el cheque",,"AFEX En Linea"
			frmCheque.cbxMoneda.focus
			Exit Function
		End If
		If cCur(0 & frmCheque.txtMonto.value) <= 0 Then
			MsgBox "Debe ingresar un monto válido para el cheque",,"AFEX En Linea"
			frmCheque.txtMonto.select
			frmCheque.txtMonto.focus
			Exit Function
		End If
		If cCur(0 & frmCheque.txtParidad.value) <= 0 Then
			MsgBox "Debe ingresar un rate válido para el cheque",,"AFEX En Linea"
			frmCheque.txtParidad.select
			frmCheque.txtParidad.focus
			Exit Function
		End If
		ValidarDatos = True
	End Function

	Sub CargarMenu()
		Dim sId

		frmCheque.objmenu.bgColor = document.bgColor 
		frmCheque.objmenu.stylesheet = "../Estilos/Cliente.css"
		sId = frmCheque.objmenu.addparent("Opciones")

		frmCheque.objMenu.addchild sId, "", "", "Principal"
		frmCheque.objMenu.addchild sId, "", "", ""
		frmCheque.objMenu.addchild sId, "", "", ""
	End Sub

	sub window_onload()
		CargarMenu
		cbxMoneda_onClick
		frmCheque.txtmonto.select
		frmCheque.txtmonto.focus
	end sub

-->
</script>

<body>
<!-- #INCLUDE virtual="/Compartido/Encabezado.htm" -->
<form id="frmCheque" method="post">
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
<!-- Paso 1 -->
<table ID="tabPaso1" CELLSPACING="0" CELLPADDING="0" BORDER="0" HEIGHT="100" WIDTH="560" STYLE="BORDER-BOTTOM: steelblue 2px solid; BORDER-LEFT: steelblue 2px solid; BORDER-RIGHT: steelblue 2px solid; BORDER-TOP: steelblue 2px solid;  HEIGHT: 100px; LEFT: 4px; POSITION: relative; TOP: 0px; WIDTH: 566px">
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
		<!--
		<td><img border="0" id="imgTransferencia" src="../images/Transferencia.jpg" style="cursor: hand" WIDTH="19" HEIGHT="22" onClick="UltimasTransferencias"></td>
		<td style="cursor: hand" onClick="UltimasTransferencias">Ultimas Transferencias</td>
		-->
	</tr>
	</table>		
	<table>
	<tr>
	</tr>
	</table>
</td></tr>
<tr><td class="titulo">Valores</td></tr>		
<tr><td>
	<table ID="tabTransfer">	
	<tr HEIGHT="15">
	<td></td>
	<td>Monto<br>
		<input STYLE="HEIGHT: 22px; TEXT-ALIGN: right; WIDTH: 115px" NAME="txtMonto" SIZE="15" onkeypress="IngresarTexto(1)" value="<%=FormatNumber(nMonto, 2)%>">
	</td>				
	<td>Monedas de Pago<br>
		<select NAME="cbxMoneda" style="HEIGHT: 22px; WIDTH: 166px">
		<%
			'CargarMonedasTransfer sMoneda
			CargarMonedasCheques
		%>
		</select>							
		<!--
		<select name="cbxParidades" STYLE="Display: none">			
		<%
			CargarParidadesTransfer2 afxTCTransferencia, sMoneda
		%>
		</select>
		-->
		</td>
		<td>
			<!--
			Rate<br>
			<select NAME="cbxRate" style="HEIGHT: 22px; WIDTH: 106px">
				<option SELECTED VALUE="1">Automática</option>
				<option VALUE="2">Manual</option>					
			</select>
			-->
		</td>			
	<!--		
	<td>Valuta<br>
		<select NAME="cbxValuta" style="HEIGHT: 22px; WIDTH: 101px">
			<%
				CargarValuta "48"
			%>
		</select>
	</td>			
	-->
	</tr>
	<tr HEIGHT="15">
	<td></td>
	<td>Rate<br>
		<input STYLE="HEIGHT: 22px; TEXT-ALIGN: right; WIDTH: 93px" NAME="txtParidad" SIZE="15" onkeypress="IngresarTexto(1)" value="<%=FormatNumber(nParidad, 8)%>">
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
	</tr>
	</table>
</td></tr>
<tr ID="trDatosTransfer1"><td class="titulo">Datos del Beneficiario&nbsp;</td></tr>
<tr ID="trDatosTransfer2"><td>
	<table ID="tabBeneficiario" width="100%" border="0">
   <tbody>
	<tr><td>
		<table width="100%">
		<tr HEIGHT="15">
		<td>Nombre <br><input NAME="txtBeneficiario" SIZE="25" style="HEIGHT: 22px; WIDTH: 300px" onkeypresss="IngresarTexto(2)" onblur="frmCheque.txtBeneficiario.value=MayMin(frmCheque.txtBeneficiario.value)"></td>
		<td colspan="2">Pais<br>
			<select name="cbxPaisB" style="width: 150px">
				<%	
					CargarUbicacion 1, "", ""
				%>
			</select>
		</td>
		<td width="100%"></td>
		</tr>
		<tr>
		<td align="right" valign="right" colspan="4">
			<img align="right" alt border="0" hspace="0" id="imgAceptar" src="../images/BotonAceptar.jpg" style="FLOAT: right; cursor: hand" WIDTH="70" HEIGHT="20">
		</td>
		</tr>
		</table>
	</td></tr>
<!--	
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
		</tr>
		<tr HEIGHT="15">
		<td></td>
		<td>ABA<br><input NAME="txtAba" SIZE="25" style="HEIGHT: 22px; WIDTH: 126px" disabled></td>
		<td>Invoice<br><input ID="txtInvoice" NAME="txt" SIZE="25" style="HEIGHT: 22px; WIDTH: 179px" disabled></td>				
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
-->
	</tbody>
	</table>
</td></tr>
</tbody>
</table>
<table BORDER="0" cellspacing="0" cellpadding="0" STYLE="LEFT: 437px; POSITION: absolute; TOP: 50px">	
	<tr><td>
	    <object align="left" id="objMenu" style="HEIGHT: 60px; LEFT: 0px; POSITION: relative; TOP: 0px; WIDTH: 160px" type="text/x-scriptlet" width="170" VIEWASTEXT border="0" valign="top"><param NAME="Scrollbar" VALUE="0"><param NAME="URL" VALUE="../Scriptlets/Menu.htm"></object>
	</td></tr>
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

		End Select
	End Sub
	
	Sub ImprimirTransferencia()
		Dim sTo, sFonoCliente, sFaxCliente, sNombreContacto
		Dim sParamatros, sReleaseD, sValueD, sNombreCliente
		Dim sDireccionB
		
		sNombreCliente = trim(trim(frmCheque.txtNombres.value) & " " & _
						 trim(frmCheque.txtApellidoP.value) & " " & _
						 trim(frmCheque.txtApellidoM.value) & _
						 trim(frmCheque.txtRazonSocial.value))
		
		sDireccionB = ""
		
		If trim(frmCheque.txtDireccionB.value) <> Empty Then
			sDireccionB = trim(trim(frmCheque.txtDireccionB.value) & ", " & _
						  trim(frmCheque.txtCiudadB.value))
		End If					  
		
		sFonoCliente = ""
		sFaxCliente = ""
		
		If frmCheque.txtFono.value > 0 Then
			sFonoCliente = "(" + frmCheque.txtPaisFono.value & " " & _
								 frmCheque.txtAreaFono.value & ")" & " " & _
								 frmCheque.txtFono.value
		End If
		If frmCheque.txtFax.value > 0 Then
			sFaxCliente = "(" + frmCheque.txtPaisFax.value & " " & _
 								frmCheque.txtAreaFax.value & ")" & " " & _
								frmCheque.txtFax.value
		End If
		
		sReleaseD = date
		If frmCheque.cbxValuta.value = "24" Then
			sValueD = date + 1
		ElseIf frmCheque.cbxValuta.value = "48" Then
			sValueD = date + 2
		Else
			sValueD = date
		End If
		
		sTo = "<%=Session("NombreCliente")%>"
		
		sParamatros = "../Reportes/transferencia.rpt?init=actx&prompt0= &prompt1= &prompt2= " & _
	 				  "&prompt3=" & frmCheque.cbxMoneda.value & _
					  "&prompt4=" & frmCheque.txtBancoDestino.value & _
					  "&prompt5=" & frmCheque.txtCuentaDestino.value & _
					  "&prompt6=" & frmCheque.txtNombreB.value & _
					  "&prompt7=" & frmCheque.txtABA.value & _
					  "&prompt8=" & sDireccionB & _
					  "&prompt9=" & _
					  "&prompt10=" & frmCheque.txtMonto.value & _
					  "&prompt11= &prompt12= " & _
					  "&prompt13=" & frmCheque.txtMensaje.value & _
					  "&prompt14= " & _
					  "&prompt15=" & frmCheque.txtInvoice.value & _
					  "&prompt16=" & frmCheque.txtIdentificacion.value & _
					  "&prompt17=" & sReleaseD & _
					  "&prompt18=" & sValueD & _
					  "&prompt19=" & frmCheque.txtBancoIntermedio.value & _
					  "&prompt20=" & frmCheque.txtCuentaIntermedio.value & _
					  "&prompt21=" & sTo &  _
					  "&prompt22=AFEX TRANSFERENCIAS" & _
					  "&prompt23=" & sNombreCliente & _
					  "&prompt24=" & FormatoRut(frmCheque.txtRut.value) & _
					  "&prompt25=" & sFonoCliente & _
					  "&prompt26=" & sFaxCliente & _
					  "&prompt27=" & sNombreContacto & _
					  "&prompt28=" & "<%=session("NombreCliente")%>"
					  
		sParamatros = replace(sParamatros, "ñ", "n")
		sParamatros = replace(sParamatros, "Ñ", "N")
		
		window.open sParamatros, _
				    "", "dialogHeight= 800pxl; dialogWidth= 800pxl; " & _
					"dialogTop= 0; dialogLeft= 0; resizable=no; " & _
					"status=no; scrollbars=no"

	End Sub
-->
</script>
<!--#INCLUDE virtual="/Compartido/Rutinas.htm" -->
<!--#INCLUDE virtual="/Compartido/Transferencias.htm" -->
</html>
