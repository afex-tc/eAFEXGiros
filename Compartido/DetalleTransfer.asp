<%@ Language=VBScript %>
<!--#INCLUDE virtual="/compartido/Constantes.asp" -->
<%
	'Se asegura que la página no se almacene en la memoria cache
	Response.Expires = 0
	If session("CodigoCliente") = "" Then
		response.Redirect "../Compartido/TimeOut.htm"
		response.end
	End If
%>
<%
	'Variables
	Dim rsTrf, cCorrelativo, afxTrf, sCliente
	
	cCorrelativo = Request("Correlativo")
	sCliente = Request("Cliente")
	
	On Error Resume Next	

	
	'Set afxTrf = Server.CreateObject("AFEXProducto.Transferencia")
	'Set rsTrf = afxTrf.Buscar(Session("afxCnxAFEXchange"), 6, "Correlativo_transferencia = " & cCorrelativo)
	Set rsTrf = ObtenerTransfer()
	
	If Err.number <> 0 Then
		Set rsTrf = Nothing
		'Set afxTrf = Nothing
		MostrarErrorMS "Transferencia 1"
	End If
	'If afxTrf.ErrNumber <> 0 Then			
	'	Set rsTrf = Nothing
	'	MostrarErrorAFEX afxTrf, "Transferencia 2"
	'End If
	
	'set afxTrf = Nothing
	
	Function ObtenerTransfer()
		Dim rs, sSQL
		
		On Error Resume Next
		
		sSQL = "Correlativo_transferencia = " & cCorrelativo
						
		If Err.number <> 0 Then
			Set Cnnx = Nothing
			MostrarErrorMS "Detalle Transfer"
		End If

		Set rs = BuscarTRF(Session("afxCnxAFEXchange"), 6, sSQL, 0, True, False)		
		
		If Err.number <> 0 Then
			Set rs = Nothing
			MostrarErrorMS "Detalle Transfer"
		End If
		
		Set ObtenerTransfer = rs
		Set rs = Nothing
	End Function

%>
<!--#INCLUDE virtual="/Compartido/Transferencias.asp" -->
<!--#INCLUDE virtual="/Compartido/Rutinas.asp" -->
<!--#INCLUDE virtual="/Sucursal/Rutinas.asp" -->
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
	
	sEncabezadoFondo = "Información"
	sEncabezadoTitulo = "Detalle Transferencia"

	Sub CargarMenu()
		Dim sId

		frmTransfer.objmenu.bgColor = document.bgColor 
		frmTransfer.objmenu.stylesheet = "../Estilos/Cliente.css"
		sId = frmTransfer.objmenu.addparent("Opciones")
		<%	
			If rsTrf.Recordcount > 0 Then				
		%>
				frmTransfer.objMenu.addchild sId, "Imprimir Transferencia", "Imprimir", "Principal"
				<% If rsTrf("estado_transferencia") = 9 Then %>
						frmTransfer.objMenu.addchild sId, "Actualizar Estado", "Estado", "Principal"
				<% End If %>
				'If <%=cInt(0 & Request("Tipo"))%> = 1 Then
				'frmTransfer.objMenu.addchild sId, "Atención de Clientes", "ATC", "Principal"
				'Else
				'	frmTransfer.objMenu.addchild sId, "", "", ""
				'End If
				frmTransfer.objMenu.addchild sId, "", "", ""
		<%	
			End If 
		%>
	End Sub

	'Objetivo:	calcular el total de la transferencia
	Function CalcularTotal()
		CalcularTotal = 0
	
		' verifica si hay monto
		If frmTransfer.txtMonto.value = Empty Or frmTransfer.txtMonto.value = "0,00" Then Exit Function
		
		CalcularTotal = FormatNumber(cDbl(0 & frmTransfer.txtTarifaCobrada.value) + _
									cDbl(0 & frmTransfer.txtEquivalente.value), 2)
	End Function	

	sub window_onload()
		frmTransfer.txtTotal.value = CalcularTotal
		CargarMenu
		Set rsTrf = Nothing
	end sub
-->
</script>

<body>
<!-- #INCLUDE virtual="/Compartido/Encabezado.htm" -->
<form id="frmTransfer" method="post">
<input type="hidden" name="txtRut" value="<%=rsTrf("rut_cliente")%>">
<input type="hidden" name="txtPaisFono" value="<%=rsTrf("pais_telefono_particular")%>">
<input type="hidden" name="txtAreaFono" value="<%=rsTrf("area_telefono_particular")%>">
<input type="hidden" name="txtFono" value="<%=rsTrf("numero_telefono_particular")%>">
<input type="hidden" name="txtPaisFax" value="<%=rsTrf("pais_fax_particular")%>">
<input type="hidden" name="txtAreaFax" value="<%=rsTrf("area_fax_particular")%>">
<input type="hidden" name="txtFax" value="<%=rsTrf("numero_fax_particular")%>">
<input type="hidden" name="txtReleaseD" value="<%=rsTrf("fecha_release")%>">
<input type="hidden" name="txtValueD" value="<%=rsTrf("fecha_value")%>">
<input type="hidden" name="txtIdentificacion" value="<%=rsTrf("identificacion")%>">
<input type="hidden" name="txtMensaje" value="<%=rsTrf("mensaje_destino")%>">

<!-- Paso 1 -->
<table class="Borde" ID="tabPaso1" CELLSPACING="0" CELLPADDING="0" BORDER="0" HEIGHT="100" WIDTH="560" STYLE="  HEIGHT: 100px; LEFT: 4px; POSITION: relative; TOP: 0px; WIDTH: 566px">
<tr><td colspan="2" align="center"><font FACE="Verdana" SIZE="2"><%=sVigencia%></font></td></tr>
<tr><td class="titulo">Datos del Remitente</td></tr>
<tr><td>
	<table ID="tabRemitente" width="100%">
	<tr HEIGHT="15">
		<td></td>
		<td>Nombres<br>
			<input NAME="txtNombres" SIZE="25" style="HEIGHT: 22px; WIDTH: 400px" disabled value="<%=Trim(rsTrf("nombre_completo_cliente"))%>">
		</td>
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
		<input STYLE="HEIGHT: 22px; TEXT-ALIGN: right; WIDTH: 115px" NAME="txtMonto" SIZE="15" onkeypress="IngresarTexto(1)" disabled value="<%=FormatNumber(rsTrf("monto_transferencia"), 2)%>">
	</td>				
	<td>Monedas de Pago<br>
		<select NAME="cbxMoneda" style="HEIGHT: 22px; WIDTH: 166px" disabled>
		<%
			CargarMonedasTransfer rsTrf("codigo_moneda")
		%>
		</select>							
		<select name="cbxParidades" STYLE="Display: none" disabled>	
		<%
			CargarParidadesTransfer afxTCTransferencia, rsTrf("codigo_moneda")
		%>
		</select></td>
		<td>Rate<br>
		<select NAME="cbxRate" style="HEIGHT: 22px; WIDTH: 106px" disabled>
			<% If rsTrf("tipo_paridad") = 1 Then %>
				<option VALUE="1">Automática</option>
			<% ElseIf rsTrf("tipo_paridad") = 2 Then %>
				<option VALUE="2">Manual</option>					
			<% End If %>
		</select></td>			
	<td>Valuta<br>
		<select NAME="cbxValuta" style="HEIGHT: 22px; WIDTH: 101px" disabled>
			<% If cInt(0 & EvaluarVar(rsTrf("tiempo_despacho"), "")) = 0 Then %>
				<option VALUE="0">Express</option>
			<% ElseIf rsTrf("tiempo_despacho") = 24 Then %>
				<option VALUE="1">24 hrs</option>
			<% ElseIf rsTrf("tiempo_despacho") = 48 Then %>
				<option VALUE="2">48 hrs</option>
			<% End If %>
		</select>
	</td>			
	</tr>
	<tr HEIGHT="15">
	<td></td>
	<td>Rate<br>
		<%
			Dim nParidad
			If rsTrf("Paridad")=0 And IsNull(rsTrf("Moneda_Equivalente")) Then
				nParidad = 1
			Else
				nParidad = rsTrf("Paridad")
			End If
		%>
		<input STYLE="HEIGHT: 22px; TEXT-ALIGN: right; WIDTH: 93px" NAME="txtParidad" SIZE="15" disabled value="<%=FormatNumber(rsTrf("Paridad"), 8)%>">
	</td>
	<td>Equivalente USD<br>
		<input STYLE="HEIGHT: 22px; TEXT-ALIGN: right; WIDTH: 115px" NAME="txtEquivalente" SIZE="15" disabled value="<%=FormatNumber(rsTrf("monto_equivalente"), 2)%>">
	</td>
	<td>Tarifa Sugerida<br>
		<input STYLE="HEIGHT: 22px; TEXT-ALIGN: right; WIDTH: 100px" NAME="txtTarifaSugerida" SIZE="15" disabled value="<%=FormatNumber(rsTrf("tarifa_sugerida"), 2)%>">
	</td>
	<td>Tarifa Cobrada<br>
		<input STYLE="HEIGHT: 22px; TEXT-ALIGN: right; WIDTH: 100px" ID NAME="txtTarifaCobrada" SIZE="15" onkeypress="IngresarTexto(1)" disabled value="<%=FormatNumber(rsTrf("tarifa_cobrada"), 2)%>">
	</td>
	</tr>
	<tr HEIGHT="15">
	<td></td>
	<td>Gasto USD<br>
		<input STYLE="HEIGHT: 22px; TEXT-ALIGN: right; WIDTH: 93px" ID NAME="txtGastoExtranjera" SIZE="15" disabled value="<%=FormatNumber(rsTrf("comision"), 2)%>">
	</td>
	<!--<td>Gasto Pesos<br>-->
		<input Type="Hidden" STYLE="HEIGHT: 22px; TEXT-ALIGN: right; WIDTH: 100px" ID NAME="txtGastoNacional" SIZE="15" disabled value="0,00">
	<!--</td>-->
	<td>Total a Pagar<br>
		<input STYLE="HEIGHT: 22px; TEXT-ALIGN: right; WIDTH: 115px" NAME="txtTotal" SIZE="15" disabled value="0,00">
	</td>
	<td>Numero<br>
		<input STYLE="HEIGHT: 22px; TEXT-ALIGN: right; WIDTH: 100px" NAME="txtNumero" SIZE="15" disabled value="<%=rsTrf("numero_transferencia")%>">
	</td>
	<td>Estado<br>
		<% If rsTrf("estado_transferencia") <> 9 Then %> 
				<input STYLE="HEIGHT: 22px; WIDTH: 100px" NAME="txtEstado" disabled value="<%=Trim(rsTrf("nombre_estado"))%>">
		<% Else %>
				<select NAME="cbxEstado" style="HEIGHT: 22px; WIDTH: 100px">
					<%
						CargarEstadoTRF rsTrf("estado_transferencia")
					%>
				</select>
		<% End If %>
	</td>
	</tr>
	</table>
</td></tr>
<tr ID="trDatosTransfer1"><td class="titulo">Datos de la Transferencia &nbsp;
	<font ID="dspDatosTransfer" FACE="Marlett" SIZE="3">t</font> 
</td></tr>
<tr ID="trDatosTransfer2" style="DISPLAY: "><td>
	<table ID="tabBeneficiario" width="100%" border="0">
   <tbody>
	<tr><td>
		<table width="100%">
		<tr><td colspan="3" class="titulo" style="BACKGROUND-COLOR: lightgrey">Origen</td></tr>
		<tr HEIGHT="15">
		<td></td>
		<td>Banco <br><input NAME="txtBancoOrigen" SIZE="25" style="HEIGHT: 22px; WIDTH: 301px" value="<%=rsTrf("nombre_banco_origen")%>" disabled></td>
		<td>Cuenta Corriente<br><input NAME="txtCuentaOrigen" SIZE="25" style="HEIGHT: 22px; WIDTH: 179px" value="<%=rsTrf("cuenta_corriente_origen")%>" disabled></td>
		</tr>
		</table>
	</td></tr>
	<tr><td>
		<table width="100%" border="0">		
		<tr><td colspan="3" class="titulo" style="BACKGROUND-COLOR: lightgrey">Destino</td></tr>
		<tr HEIGHT="15">
		<td></td>
		<td>Banco <br><input NAME="txtBancoDestino" SIZE="25" style="HEIGHT: 22px; WIDTH: 299px" value="<%=rsTrf("nombre_banco_destino")%>" disabled></td>
		<td>Cuenta Corriente<br><input name="txtCuentaDestino" SIZE="25" style="HEIGHT: 22px; WIDTH: 179px" value="<%=rsTrf("cuenta_corriente_destino")%>" disabled></td>
		</tr>
		<tr HEIGHT="15">
		<td></td>
		<td colspan="2">Nombre Beneficiario<br><input name="txtNombreB" SIZE="25" style="HEIGHT: 22px; WIDTH: 299px" value="<%=rsTrf("nombre_titular_destino")%>" disabled></td>
		</tr>
		<tr HEIGHT="15">
		<td></td>
		<td>Dirección<br><input NAME="txtDireccionB" SIZE="50" style="HEIGHT: 22px; WIDTH: 321px" value="<%=rsTrf("direccion_destino")%>" disabled></td>
		<td>Ciudad<br><input NAME="txtCiudadB" style="HEIGHT: 22px; WIDTH: 153px" value="<%=rsTrf("nombre_ciudad_destino")%>" disabled></td>
		</tr>
		<tr HEIGHT="15">
		<td></td>
		<td>ABA<br><input NAME="txtAba" SIZE="25" style="HEIGHT: 22px; WIDTH: 126px" value="<%=rsTrf("numero_aba")%>" disabled></td>
		<td>Invoice<br><input ID="txtInvoice" NAME="txt" SIZE="25" style="HEIGHT: 22px; WIDTH: 179px" value="<%=rsTrf("invoice")%>" disabled></td>				
		</tr>
		</table>
	</td></tr>
	<tr><td width="60%">
		<table width="100%" border="0">		
		<tr><td colspan="3" class="titulo" style="BACKGROUND-COLOR: lightgrey">Intermediario</td></tr>
		<tr HEIGHT="15">
		<td></td>
		<td>Banco<br><input NAME="txtBancoIntermedio" SIZE="25" style="HEIGHT: 22px; WIDTH: 303px" value="<%=rsTrf("banco_intermediario")%>" disabled></td>
		<td>Cuenta Corriente<br><input NAME="txtCuentaIntermedio" SIZE="25" style="HEIGHT: 22px; WIDTH: 179px" value="<%=rsTrf("cuenta_intermediario")%>" disabled></td>
		</tr>
		<tr HEIGHT="15">
		<td></td>
		<td>Dirección<br><input NAME="txtDireccionIntermedio" SIZE="50" style="HEIGHT: 22px; WIDTH: 321px" value="<%=rsTrf("direccion_intermediario")%>" disabled></td>
		<td>Ciudad<br><input NAME="txtCiudadIntermedio" style="HEIGHT: 22px; WIDTH: 121px" value="<%=rsTrf("nombre_ciudad_intermediario")%>" disabled></td>
		</tr>
		</table>
	</td>
	</td></tr>
	</tbody>
	</table>
</td></tr>
</tbody>
</table>
<table BORDER="0" cellspacing="0" cellpadding="0" STYLE="LEFT: 421px; POSITION: absolute; TOP: 90px">	
	<tr><td>
	    <object align="left" id="objMenu" style="HEIGHT: 60px; LEFT: 0px; POSITION: relative; TOP: -10px; WIDTH: 160px" type="text/x-scriptlet" width="170" VIEWASTEXT border="0" valign="top"><param NAME="Scrollbar" VALUE="0"><param NAME="URL" VALUE="http:../Scriptlets/Menu.htm"></object>
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

				ElseIf Right(varEventData, 6) = "Estado" Then
					ActualizarEstado
					
				ElseIf Right(varEventData, 3) = "ATC" Then
					window.navigate "..\Agente\AtencionClientes.asp?Accion=<%=afxAccionBuscar%>&Campo=<%=afxCampoCodigoExchange%>&Argumento=<%=sCliente%>"

				End If

		End Select
	End Sub
	
	Sub ActualizarEstado()
		Dim cnn, sSQL
		
		If Not CajaPregunta("AFEX En Linea", "Está seguro que desea cambiar el estado?") Then
			Exit Sub
		End If

		window.navigate "ActualizarEstadoTRF.asp?corr=<%=cCorrelativo%>&es=" & frmTransfer.cbxEstado.value
	End Sub
	
	Sub ImprimirTransferencia()
		Dim sTo, sFonoCliente, sFaxCliente, sNombreContacto
		Dim sParamatros, sReleaseD, sValueD, sNombreCliente
		Dim sDireccionB
		
		If frmTransfer.txtDireccionB.value <> Empty Then
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
		
		If frmTransfer.txtReleaseD.value = Empty Then
			sReleaseD = date
		Else
			sReleaseD = cdate(frmTransfer.txtReleaseD.value)
		End If
		If frmTransfer.txtValueD.value = Empty Then
			sValueD = date + frmTransfer.cbxValuta.value
		Else
			sValueD = cdate(frmTransfer.txtValueD.value)
		End If
		
		sTo = "<%=Session("NombreCliente")%>"
		
		sParamatros = "../Agente/ImprimirTransfer.asp?" & _
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
					  "&rf=" & frmTransfer.txtIdentificacion.value & _
					  "&rd=" & sReleaseD & _
					  "&vd=" & sValueD & _
					  "&ib=" & frmTransfer.txtBancoIntermedio.value & _
					  "&ca=" & frmTransfer.txtCuentaIntermedio.value & _
					  "&dd=" & sTo &  _
					  "&prompt22=AFEX TRANSFERENCIAS" & _
					  "&nb=" & frmTransfer.txtNombres.value & _
					  "&rt=" & FormatoRut(frmTransfer.txtRut.value) & _
					  "&tl=" & sFonoCliente & _
					  "&fx=" & sFaxCliente & _
					  "&co=" & sNombreContacto & _
					  "&prompt28=" & "<%=session("NombreCliente")%>"
					  
		sParamatros = replace(sParamatros, "ñ", "n")
		sParamatros = replace(sParamatros, "Ñ", "N")
		
		window.open sParamatros, _
				    "", "dialogHeight= 800pxl; dialogWidth= 800pxl; " & _
					"dialogTop= 0; dialogLeft= 0; resizable=no; " & _
					"status=no; scrollbars=yes"

	End Sub

	Sub ImprimirTransferenciaOld()
		Dim sTo, sFonoCliente, sFaxCliente, sNombreContacto
		Dim sParamatros, sReleaseD, sValueD, sDireccionB
		
		If frmTransfer.txtDireccionB.value <> Empty Then
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
		
		If frmTransfer.txtReleaseD.value = Empty Then
			sReleaseD = date
		Else
			sReleaseD = cdate(frmTransfer.txtReleaseD.value)
		End If
		If frmTransfer.txtValueD.value = Empty Then
			sValueD = date + frmTransfer.cbxValuta.value
		Else
			sValueD = cdate(frmTransfer.txtValueD.value)
		End If
		
		sTo = "<%=Session("NombreCliente")%>"
		
		sParamatros = "../Reportes/transferencia.rpt?init=actx&prompt0= &prompt1= &prompt2= " & _
	 				  "&prompt3=" & frmTransfer.cbxMoneda.value & _
					  "&prompt4=" & frmTransfer.txtBancoDestino.value & _
					  "&prompt5=" & frmTransfer.txtCuentaDestino.value & _
					  "&prompt6=" & frmTransfer.txtNombreB.value & _
					  "&prompt7=" & frmTransfer.txtABA.value & _
					  "&prompt8=" & sDireccionB & _
					  "&prompt9=" & _
					  "&prompt10=" & frmTransfer.txtMonto.value & _
					  "&prompt11= &prompt12= " & _
					  "&prompt13=" & frmTransfer.txtMensaje.value & _
					  "&prompt14= " & _
					  "&prompt15=" & frmTransfer.txtInvoice.value & _
					  "&prompt16=" & frmTransfer.txtIdentificacion.value & _
					  "&prompt17=" & sReleaseD & _
					  "&prompt18=" & sValueD & _
					  "&prompt19=" & frmTransfer.txtBancoIntermedio.value & _
					  "&prompt20=" & frmTransfer.txtCuentaIntermedio.value & _
					  "&prompt21=" & sTo &  _
					  "&prompt22=AFEX TRANSFERENCIAS" & _
					  "&prompt23=" & frmTransfer.txtNombres.value & _
					  "&prompt24=" & FormatoRut(frmTransfer.txtRut.value) & _
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
