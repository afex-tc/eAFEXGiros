<%@ Language=VBScript %>
<!--#INCLUDE virtual="/Compartido/Rutinas.asp" -->
<!--#INCLUDE virtual="/Compartido/Rutinas.htm" -->

<%
	Dim sPaisB, sCiudadB, sNombrePaisB, sNombreciudadB, sVerCompra, sVerVenta, sVerGiro, sVerTCredito
		
	sPaisB = Request.Form("cbxPaisB")
	sCiudadB = Request.Form("cbxCiudadB")
	sNombrePaisB = Request.Form("txtPaisB")
	sNombreCiudadB = Request.Form("txtCiudadB")
		
	sVerCompra = "none"
	sVerVenta = "none"
	sVerGiro = "none"	
	sVerTCredito = "none"
	Select Case Request("Ver") 
		Case 1
			sVerCompra = ""		
		Case 2
			sVerVenta = ""
		Case 3
			sVerGiro = ""
		Case 4
			sVerGiro = ""
			sNombrePaisB = "CHILE"
		Case 5
			sVerTCredito = ""
			
	End Select
	
%>

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link href="../CSS/Linksnuevos.css" rel="stylesheet" type="text/css">
<link href="../CSS/Links_T1T2T3.css" rel="stylesheet" type="text/css">
</HEAD>

<script language="vbscript">
<!--
	Sub window_onLoad()
		Dim sValor
		
		Select Case <%=Request("Ver")%>
			Case 1	' comprar
				sValor = window.showModalDialog("Tarifas.asp?Tipo=1")
				if Len(sValor) > 10 Then
					msgbox sValor,,"Comprar"
		
				else
					frmComprar.txtTCC.value = ""
					frmComprar.txtTCC.value = formatnumber(round(sValor, 0), 0)
				end if
				frmComprar.txtMontoC.focus
			
			Case 2	' vender
				sValor = window.showModalDialog("Tarifas.asp?Tipo=2")
				if Len(sValor) > 10 Then
					msgbox sValor,,"Comprar"
		
				else
					frmVender.txtTCV.value = ""
					frmVender.txtTCV.value = formatnumber(round(sValor, 0), 0)
				end if
				frmVender.txtMontoV.focus
		
			Case 3 	' giro internacional
				frmGiro.txtNombresB.focus
				frmGiro.cbxMoneda.value = "USD"
				frmGiro.cbxMoneda.disabled = True
		
			Case 4	' giro nacional
				frmGiro.txtNombresB.focus
				frmGiro.cbxMoneda.value = "CLP"
			
			Case 5	' Tarjeta de credito
				sValor = window.showModalDialog("Tarifas.asp?Tipo=2")
				if Len(sValor) > 10 Then
					msgbox sValor,,"Pagar Tarjeta"
		
				else
					frmTCredito.txtTCT.value = ""
					frmTCredito.txtTCT.value = formatnumber(round(sValor, 0), 0)
				end if				
				frmTCredito.txtMontoTC.focus
			
		End Select
	End Sub

	Sub cbxPaisB_onChange()
		frmGiro.txtPaisB.value = frmGiro.cbxPaisB.options(frmGiro.cbxPaisB.selectedIndex).text
		
		frmGiro.action = "OperacionesOnLine.asp?Ver=3"
		frmGiro.submit()
		frmGiro.action = ""
	End Sub
	
	Sub cbxCiudadB_onChange()
		frmGiro.txtCiudadB.value = frmGiro.cbxCiudadB.options(frmGiro.cbxCiudadB.selectedIndex).text
		
		CalcularTarifa	
	End Sub
	
	Sub cbxMoneda_onChange()		
		
		CalcularTarifa	
	End Sub
	
	Sub CalcularTarifa()
		Dim sValor
		Dim sPais		
		
		' saca la tarifa para la ciudad seleccionada
		If frmGiro.txtCiudadB.value = Empty Or frmGiro.txtMonto.value = Empty Or frmGiro.cbxMoneda.value = Empty Then exit sub	
		
		If <%=Request("Ver")%> = 4 Then
			sPais = "CL"
		Else
			sPais = frmGiro.cbxPaisB.value
		End If
		
		sValor = window.showModalDialog("Tarifas.asp?Tipo=3&Monto=" & frmGiro.txtMonto.value & _
																	"&Pais=" & sPais & "&Ciudad=" & frmGiro.cbxCiudadB.value & _
																	"&Moneda=" & frmGiro.cbxMOneda.value)
		if Len(sValor) > 10 Then
			msgbox sValor,,"Giro"
		
		else
			frmGiro.txtTarifa.value = ""
			frmGiro.txtTarifa.value = formatnumber(sValor , 2)
			frmGiro.txtTotal.value = formatnumber(ccur(frmGiro.txtMonto.value) + ccur(frmGiro.txtTarifa.value), 2)
		end if		
	End Sub
		
	Sub cmdEnviar_onClick()
		' verifica los campos vacios
		if frmGiro.txtNombresB.value = Empty And frmGiro.txtApellidosB.value = Empty And _
			frmGiro.txtPaisB.value = Empty And frmGiro.txtCiudadB.value = Empty And _
			frmGiro.txtFonoB.value = Empty And frmGiro.txtMonto.value = Empty Then
			
			msgbox "Debe completar todos los datos Solicitados.",,"Enviar Giro"
			exit sub
		End If
		
		' habilita los objetos		
		frmGiro.cbxMoneda.disabled = False
		frmGiro.txtTarifa.disabled = False
		frmGiro.txtTotal.disabled = False
	
		frmGiro.action = "EnviarMail.asp?Accion=3"
		frmGiro.submit()
		frmGiro.action = ""		
	End Sub
	
	Sub cmdEnviarC_onClick()
		' verifica los campos vacios
		if frmComprar.txtMontoC.value = Empty Or frmComprar.txtBancoC.value = Empty Or frmComprar.txtCtaCteC.value = Empty Then		
			msgbox "Debe ingresar todos los datos solicitados.",,"Comprar US$"
			exit sub
		End If
		
		' habilita los objetos
		frmComprar.txtTCC.disabled = False
		frmComprar.txtTotalC.disabled = False
	
		frmComprar.action = "EnviarMail.asp?Accion=1"
		frmComprar.submit()
		frmComprar.action = ""		
	End Sub
	
	Sub cmdEnviarV_onClick()
		' verifica los campos vacios
		if frmVender.txtMontoV.value = Empty Or frmVender.txtBancoV.value = Empty Or frmVender.txtCtaCteV.value = Empty Then		
			msgbox "Debe ingresar todos los datos solicitados.",,"Vender US$"
			exit sub
		End If
		
		' habilita los objetos
		frmVender.txtTCV.disabled = False
		frmVender.txtTotalV.disabled = False
	
		frmVender.action = "EnviarMail.asp?Accion=2"
		frmVender.submit()
		frmVender.action = ""
	End Sub

	Sub cmdEnviarTC_onClick()
		' verifica los campos vacios
		if frmTCredito.txtMontoTC.value = Empty Or frmTCredito.txtNTarjeta.value = Empty Or frmTCredito.txtBanco.value = Empty Or _
			frmTCredito.txtTCT.value = Empty Or frmTCredito.txtTotalT.value = Empty Then
			msgbox "Debe ingresar todos los datos solicitados.",,"Pagar Tarjeta"
			exit sub
		End If
		
		' habilita los objetos
		frmTCredito.txtTCT.disabled = False
		frmTCredito.txtTotalT.disabled = False
		
		frmTCredito.action = "EnviarMail.asp?Accion=5"
		frmTCredito.submit()
		frmTCredito.action = ""
	End Sub

	Sub txtMonto_onBlur()
		If isnumeric(frmGiro.txtMonto.value) Then
			frmGiro.txtMonto.value = formatnumber(frmGiro.txtMonto.value, 2)
			CalcularTarifa
		
		else
			frmGiro.txtMonto.value = Empty
		End If
		
	End Sub

	Sub txtMontoC_onBlur()
		If isnumeric(frmComprar.txtMontoC.value) Then
			frmComprar.txtMontoC.value = formatnumber(round(frmComprar.txtMontoC.value, 0), 0)
			frmComprar.txtTotalC.value = formatnumber(round(ccur(frmComprar.txtMontoC.value) * ccur(frmComprar.txtTCC.value), 0), 0)			
		else
			frmComprar.txtMontoC.value = Empty
			frmComprar.txtTotalC.value = Empty
		end if
		
	End Sub
	Sub txtMontoV_onBlur()
		If isnumeric(frmVender.txtMontoV.value) Then
			frmVender.txtMontoV.value = formatnumber(round(frmVender.txtMontoV.value, 0), 0)
			frmVender.txtTotalV.value = formatnumber(round(ccur(frmVender.txtMontoV.value) * ccur(frmVender.txtTCV.value), 0), 0)
		else
			frmVender.txtMontoV.value = Empty
			frmVender.txtTotalV.value = Empty
		End If		
	End Sub	
	
	Sub txtMontoTC_onBlur()
		If isnumeric(frmTCredito.txtMontoTC.value) Then
			frmTCredito.txtMontoTC.value = formatnumber(round(frmTCredito.txtMontoTC.value, 0), 0)
			frmTCredito.txtTotalT.value = formatnumber(round(ccur(frmTCredito.txtMontoTC.value) * ccur(frmTCredito.txtTCT.value), 0), 0)
		else
			frmTCredito.txtMontoTC.value = Empty
			frmTCredito.txtTotalT.value = Empty
		End If
		
	End Sub
-->
</script>

<BODY  leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0" marginwidth="0" style="font-family: Tahoma; font-size: 11px;">

	<form method="post" name="frmGiro" action="" style="display: <%=sVerGiro%>">
		<table>
			<tr>
				<td><b>Beneficiario</b></td>
			</tr>
			<tr>
				<td>Nombres</td>
				<td><input type="text" maxlength="30" name="txtNombresB" onKeyPress="IngresarTexto(2)" value="<%=Request.Form("txtNombresB")%>" size="30">
					&nbsp;&nbsp;&nbsp;Apellidos&nbsp;&nbsp;&nbsp;<input type="text" maxlength="30" name="txtApellidosB" onKeyPress="IngresarTexto(2)" value="<%=Request.Form("txtApellidosB")%>" size="30">
				</td>
			</tr>
			
			<%If Request("Ver") = 3 Then%>
			<tr>
				<td>Pais</td>
				<td>
					<select name="cbxPaisB">
					<% CargarUbicacion 1, "", sPaisB %>
					</select>
				</td>				
			</tr>
			<%End If%>
			<tr>
				<td>Ciudad</td>		
				<td>
					<select name="cbxCiudadB">
					<% 
						If Request("Ver") = 4 Then
							CargarCiudadesPais "CL", ""
						Else
							CargarCiudadesPais sPaisB, "" 
						End If 
					%>
					</select>
				</td>
			</tr>			
			<tr>
				<td>Dirección</td>
				<td><input type="text" maxlength="100" name="txtDireccionB" value="<%=Request.Form("txtDireccionB")%>" size="50"></td>				
			</tr>
			<tr>
				<td>Teléfono</td>
				<td><input type="text" maxlength="10" name="txtFonoB" onKeyPress="IngresarTexto(1)" value="<%=Request.Form("txtFonoB")%>"></td>
			</tr>
			<tr>
				<td>Moneda</td>
				<td>
					<select name="cbxMoneda">
						<option selected value="USD">DOLAR</option>
						<option value="CLP">PESOS</option>
					</select>
				</td>				
			</tr>
			<tr>
				<td>Monto</td>
				<td><input type="text" maxlength="10" name="txtMonto" value="<%=Request.Form("txtMonto")%>" onKeyPress="IngresarTexto(1)"></td>				
			</tr>
			<tr>
				<td>Tarifa</td>
				<td><input type="text" name="txtTarifa" value="<%=Request.Form("txtTarifa")%>" disabled></td>
			</tr>
			<tr>
				<td>Total</td>
				<td><input type="text" name="txtTotal" value="<%=Request.Form("txtTotal")%>" disabled></td>
			</tr>
			<tr>
		</table>
		<center><input type="button" style="cursor: hand" name="cmdEnviar" value="Enviar"></center>
		<input type="hidden" name="txtPaisB" value="<%=sNombrePaisB%>">
		<input type="hidden" name="txtCiudadB" value="<%=sNombreCiudadB%>">
	</form>
	
	<form method="post" name="frmComprar" action="" style="display: <%=sVerCompra%>">
		<table width="540" border="0" cellspacing="0" class="Borde_tabla_abajo">			
			<tr>
			  <td colspan="3" bgcolor="#31514A" class="textoempresa"><img src="../Img/titulos_virtual_Cambios.jpg" width="130" height="16"></td>
		  </tr>
			<tr>
			  <td colspan="3" class="textoempresa"><table width="538" border="0" cellspacing="0" bgcolor="#CCCCCC">
                <tr>
                  <td width="10"><img src="../Img/img_pagempresa/Link_01.jpg" width="10" height="8" /></td>
                  <td width="521"><strong class="textoempresa">AFEX Compra US$</strong></td>
                </tr>
              </table></td>
		  </tr>
			<tr>
				<td width="172" class="textoempresa">Monto US$</td>
				<td width="58" class="textoempresa">:</td>
				<td width="298"><input name="txtMontoC" type="text" class="Borde_tabla_abajo" onKeyPress="IngresarTexto(1)" maxlength="10"></td>				
			</tr>
			<tr>
				<td class="textoempresa">Tipo Cambio</td>
				<td class="textoempresa">:</td>
				<td><input name="txtTCC" type="text" disabled class="Borde_tabla_abajo"></td>				
			</tr>
			<tr>
				<td class="textoempresa">Total en $</td>
				<td class="textoempresa">:</td>
				<td><input name="txtTotalC" type="text" disabled class="Borde_tabla_abajo"></td>
			</tr>
			<tr>
				<td colspan="2">&nbsp;</td>
			</tr>
			<tr>
				<td colspan="3"><table width="371" border="0" cellspacing="1">
                  <tr>
                    <td width="10"><img src="../Img/img_pagempresa/Link_01.jpg" width="10" height="8" /></td>
                    <td width="354"><b class="Estilo2">Datos para que AFEX deposite los Pesos de esta operaci&oacute;n:</b></td>
                  </tr>
                </table></td>
			</tr>
			<tr>
				<td class="textoempresa">Banco</td>
				<td class="textoempresa">:</td>
				<td><input name="txtBancoC" type="text" class="Borde_tabla_abajo" onKeyPress="IngresarTexto(2)" size="30" maxlength="30"></td>
			</tr>
			<tr>
				<td class="textoempresa">Cta. Cte.</td>
				<td class="textoempresa">:</td>
				<td><input name="txtCtaCteC" type="text" class="Borde_tabla_abajo" size="30" maxlength="30"></td>
			</tr>
			<tr>
			  <td class="textoempresa">&nbsp;</td>
			  <td class="textoempresa"><div align="right"><img src="../Img/igual.jpg" width="13" height="9"></div></td>
			  <td>
			    <div align="left">
			      <input name="cmdEnviarC" type="button" class="textoempresa" style="cursor: hand" value="Enviar">
		          </div></td></tr>		
	  </table>
		<center>
		</center>		
	</form>
	
	<form method="post" name="frmVender" action="" style="display: <%=sVerVenta%>">
		<table width="540" border="0" cellspacing="0" class="Borde_tabla_abajo">			
			<tr>
			  <td colspan="3" bgcolor="#31514A" class="textoempresa"><img src="../Img/titulos_virtual_Cambios.jpg" width="130" height="16"></td>
		  </tr>
			<tr>
			  <td colspan="3" class="textoempresa"><table width="538" border="0" cellspacing="0" bgcolor="#CCCCCC">
                <tr>
                  <td width="11"><img src="../Img/img_pagempresa/Link_01.jpg" width="10" height="8" /></td>
                  <td width="520" class="textoempresa">&nbsp;<strong>AFEX Vende US$</strong></td>
                </tr>
              </table></td>
		  </tr>
			<tr>
				<td width="174" class="textoempresa">Monto US$</td>
				<td width="57" class="textoempresa">:</td>
				<td width="297"><input name="txtMontoV" type="text" class="Borde_tabla_abajo" onKeyPress="IngresarTexto(1)" maxlength="10"></td>				
			</tr>
			<tr>
				<td class="textoempresa">Tipo Cambio</td>
				<td class="textoempresa">:</td>
				<td><input name="txtTCV" type="text" disabled class="Borde_tabla_abajo"></td>				
			</tr>
			<tr>
				<td class="textoempresa">Total en $</td>
				<td class="textoempresa">:</td>
				<td><input name="txtTotalV" type="text" disabled class="Borde_tabla_abajo"></td>
			</tr>
			<tr>
				<td colspan="2">&nbsp;</td>
			</tr>
			<tr>
				<td colspan="3"><table width="360" border="0">
                  <tr>
                    <td width="10"><img src="../Img/img_pagempresa/Link_01.jpg" width="10" height="8" /></td>
                    <td width="340"><b class="Estilo2">Datos para que AFEX deposite los Dolares de esta operaci&oacute;n:</b></td>
                  </tr>
                </table></td>
			</tr>
			<tr>
				<td class="textoempresa">Banco</td>
				<td class="textoempresa">:</td>
				<td><input name="txtBancoV" type="text" class="Borde_tabla_abajo" onKeyPress="IngresarTexto(2)" size="30" maxlength="30"></td>
			</tr>
			<tr>
				<td class="textoempresa">Cta. Cte.</td>
				<td class="textoempresa">:</td>
				<td><input name="txtCtaCteV" type="text" class="Borde_tabla_abajo" size="30" maxlength="30"></td>
			</tr>
			<tr>
			  <td class="textoempresa">&nbsp;</td>
			  <td class="textoempresa"><div align="right"><img src="../Img/igual.jpg" width="13" height="9"></div></td>
			  <td>
			    <div align="left">
			      <input name="cmdEnviarV" type="button" class="textoempresa" style="cursor: hand" value="Enviar">
		          </div></td></tr>
	  </table>
		<center>
		</center>		
	</form>
	
	<form method="post" name="frmTCredito" action="" style="display: <%=sVerTCredito%>">
		<table width="540" border="0" cellspacing="0" class="Borde_tabla_abajo">			
			<tr>
			  <td colspan="3"><table width="539" border="0" cellspacing="0">
                <tr>
                  <td colspan="2" bgcolor="#31514A"><img src="../Img/titulos_virtual_pagos.jpg" width="109" height="16"></td>
                </tr>
                <tr>
                  <td colspan="15"></td>
                </tr>
                <tr>
                  <td width="10"><img src="../Img/img_pagempresa/Link_01.jpg" width="10" height="8" /></td>
                  <td width="318" class="Estilo2">Pagar con tarjeta de Cr&eacute;dito </td>
                </tr>
              </table></td>
		  </tr>
			<tr>
				<td width="175" class="textoempresa">Monto US$</td>
				<td width="50" class="textoempresa">:</td>
				<td width="308"><input name="txtMontoTC" type="text" class="Borde_tabla_abajo" onKeyPress="IngresarTexto(1)" maxlength="10"></td>				
			</tr>
			<tr>
				<td class="textoempresa">Nº de Cta. de Tarjeta</td>
				<td class="textoempresa">:</td>
				<td><input name="txtNTarjeta" type="text" class="Borde_tabla_abajo" maxlength="20"></td>				
			</tr>
			<tr>
				<td class="textoempresa">Banco emisor de la Tarjeta</td>
				<td class="textoempresa">:</td>
				<td><input name="txtBanco" type="text" class="Borde_tabla_abajo" onKeyPress="IngresarTexto(2)" maxlength="30"></td>
			</tr>
			<tr>
				<td class="textoempresa">Tipo Cambio</td>
				<td class="textoempresa">:</td>
				<td><input name="txtTCT" type="text" disabled class="Borde_tabla_abajo"></td>
			</tr>
			<tr>
				<td class="textoempresa">Pesos a depositar</td>
				<td class="textoempresa">:</td>
				<td><input name="txtTotalT" type="text" disabled class="Borde_tabla_abajo"></td>
			</tr>
			<tr>
			  <td class="textoempresa">&nbsp;</td>
			  <td class="textoempresa"><div align="right"><img src="../Img/igual.jpg" width="13" height="9"></div></td>
			  <td>
			    <div align="left">
			      <input name="cmdEnviarTC" type="button" class="textoempresa" style="cursor: hand" value="Enviar">
		          </div></td></tr>
	  </table>
		<center>
		</center>
	</form>
</BODY>
</HTML>