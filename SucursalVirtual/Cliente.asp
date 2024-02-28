<%@ Language=VBScript %>
<!--#INCLUDE virtual="/Compartido/Rutinas.asp" -->
<!--#INCLUDE virtual="/Compartido/Rutinas.htm" -->
<%
'	Dim sIDR, sTipoIDR
'	Dim rsCliente
	
'	sIDR = Request.Form("txtIDR")
'	sTipoIDR = Request.Form("cbxTipoIDR")
	
'	If sIDR <> Empty Then
'		Set rsCliente = BuscarCliente(sTipoIDR, sIDR, "", "")	
'		If rsCliente.EOF Then
'			rsCliente.Close
'			Set rsCliente = Nothing		
'			
'			sMensaje = "Este Identificador no se encuentra en nuestros registros. Debe registrarse como Cliente."
'			sPagina = "enviargiro.asp"
'		Else
'			sNombreR = MayMin(EvaluarVar(rsCliente("nombre_completo"), ""))
'			'sPaisPass = EvaluarVar(rs("codigo_paispas"), "")
'			sDireccionR = MayMin(EvaluarVar(rsCliente("direccion"), ""))
			'sPaisR = EvaluarVar(rs("codigo_pais"), "")
'			'sCiudadR = EvaluarVar(rs("codigo_ciudad"), "")
'			'sComuna = EvaluarVar(rs("codigo_comuna"), "")
'			'sPaisFono = EvaluarVar(rs("ddi_pais"), "")
'			'sAreaFono = EvaluarVar(rs("ddi_area"), "")
'			sFonoR = "(" & EvaluarVar(rsCliente("ddi_pais"), "") & " " & EvaluarVar(rsCliente("ddi_area"), "") & ") " & EvaluarVar(rsCliente("telefono"), "")
'			'sPaisFono2 = EvaluarVar(rs("ddi_pais2"), "")
'			'sAreaFono2 = EvaluarVar(rs("ddi_area2"), "")
'			'sFono2 = EvaluarVar(rs("telefono2"), "")
'			sPaisR = MayMin(EvaluarVar(rsCliente("pais"), ""))
'			sCiudadR = MayMin(EvaluarVar(rsCliente("ciudad"), ""))
'			'sNombreComuna = MayMin(EvaluarVar(rs("comuna"), ""))
'			'sNombrePaisPass = MayMin(EvaluarVar(rs("paispas"), ""))	
'		End If
'	End If
	
	
'	Set rsCliente = Nothing	
%>

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link href="../CSS/CSS_SucursalVirtual.css" rel="stylesheet" type="text/css">
<link href="../CSS/Links_T1T2T3.css" rel="stylesheet" type="text/css">
</HEAD>

<script language="vbscript">
<!--
	Sub window_onLoad()
	'	frmCliente.txtIDR.focus
	'	If "<%=sTipoIDR%>" <> Empty Then			
	'		frmCliente.cbxTipoIDR.value = "<%=sTipoIDR%>"
	'	
	'	Else
	'		frmCliente.cbxTipoIDR.value = 1
	'	End If
	'	
	'	if "<%=sNombreR%>" <> Empty then
	'		frmCliente.optCompra.disabled = false
	'		frmCliente.optVenta.disabled = false 
	'		frmCliente.optGiro.disabled = false  
	'	end if
	'	
	'	<%If Request("Ver") = 1 Then%>
	'		frmCliente.optCompra.checked = True 
	'	<%ElseIF Request("Ver") = 2 Then%>
	'		frmCliente.optVenta.checked = True 
	'	<%ElseIf Request("Ver") = 3 Then%>	
	'		frmCliente.optGiro.checked = True 
	'	<%End If%>
	
		frmCliente.optGiro.click()		
	End Sub

'	Sub cmdBuscar_onClick()
'		If frmCliente.txtIDR.value = Empty Then Exit Sub
'		
'		frmCliente.action = "OperacionesOnLine.asp"
'		frmCliente.submit()
'		frmCliente.action = ""
'	End Sub

'	Sub txtIDR_onBlur()
'		If frmCliente.cbxTipoIDR.value = 1 And frmCliente.txtIDR.value <> Empty Then 
'			sRut = ValidarRut(frmCliente.txtIDR.value)
'			If sRut = Empty Then
'				msgbox "El número de Rut no es válido"
'				frmCliente.txtIDR.select
'				frmCliente.txtIDR.focus()
'			Else
'				frmCliente.txtIDR.value = sRut
'			End If
'		End If
'	End Sub
	
	Sub optCompra_onClick()
	'	dim sValor
	'	
		frmCliente.optVenta.checked = False
		frmCliente.optGiro.checked = False
		frmCliente.optGiroNac.checked = False
	'	frmGiro.style.display = "none"
	'	frmVender.style.display = "none"
	'	frmComprar.style.display = ""
	'	
	'	sValor = window.showModalDialog("Tarifas.asp?Tipo=1")
	'	if Len(sValor) > 10 Then
	'		msgbox sValor,,"Comprar"
	'	
	'	else
	'		frmComprar.txtTCC.value = ""
	'		frmComprar.txtTCC.value = formatnumber(sValor, 2)
	'	end if
		window.open "operacionesonline.asp?Ver=1", "Principal"	
	End Sub
	Sub optVenta_onClick()
		frmCliente.optCompra.checked = False
		frmCliente.optGiro.checked = False
		frmCliente.optGiroNac.checked = False
		frmCliente.optTransfer.checked = False
		frmCliente.optTCredito.checked = False
	
		window.open "operacionesonline.asp?Ver=2", "Principal"
	End Sub
	Sub optGiro_onClick()
		frmCliente.optVenta.checked = False
		frmCliente.optCompra.checked = False
		frmCliente.optGiroNac.checked = False
		frmCliente.optTransfer.checked = False
		frmCliente.optTCredito.checked = False
			
		window.open "operacionesonline.asp?Ver=3", "Principal"
	End Sub	
	Sub optGiroNac_onClick()
		frmCliente.optVenta.checked = False
		frmCliente.optCompra.checked = False
		frmCliente.optGiro.checked = False
		frmCliente.optTransfer.checked = False
		frmCliente.optTCredito.checked = False	
		
		window.open "operacionesonline.asp?Ver=4", "Principal"
	End Sub
	Sub optTransfer_onClick()
		frmCliente.optVenta.checked = False
		frmCliente.optCompra.checked = False
		frmCliente.optGiro.checked = False
		frmCliente.optGiroNac.checked = False
		frmCliente.optTCredito.checked = False	
		
		window.open "EnviarTransfer.asp", "Principal"
	End Sub	
	Sub optTCredito_onClick()
		frmCliente.optVenta.checked = False
		frmCliente.optCompra.checked = False
		frmCliente.optGiro.checked = False
		frmCliente.optTransfer.checked = False
		frmCliente.optGiroNac.checked = False	
		
		window.open "operacionesonline.asp?Ver=5", "Principal"
	End Sub		
-->
</script>

<BODY style="font-family: Tahoma; font-size: 11px;">

	<div><b><%=Session("NombreOperador")%></b></div>

	<form method="post" name="frmCliente" action="">
		<table width="100%" border="0" class="Borde_tabla_abajo">
			<!--
			<tr>
				<td><b>Cliente</b></td>
			</tr>			
			<tr>			
				<td>
					<select name="cbxTipoIDR">
						<option value="1" >Rut</option>
						<option value="2" >Pasaporte</option>
					</select>
				</td>
				<td colspan="2">
					<input type="text" name="txtIDR" value="<%=sIDR%>">&nbsp;&nbsp;&nbsp;
					<input type="button" style="cursor: hand" name="cmdBuscar" value="...">
					&nbsp;&nbsp;
					<font color="red" size="2px">
						<%=sMensaje%>
						<% If sMensaje <> Empty Then %>
							<a href="<%=sPagina%>">Reg.</a>
						<% End If %>
					</font>
				</td>
			</tr>
			
			<tr>
				<td>Nombre</td>				
				<td><input type="text" name="txtNombreR" size="50" disabled value="<%=sNombreR%>"></td>
			
				<td><b><%=Request("Cliente")%></b></td>
			</tr>
			
			<tr>
				<td>País</td>
				<td><input type="text" name="txtPaisR" disabled value="<%=sPaisR%>"></td>	
			</tr>
			<tr>
				<td>Ciudad</td>
				<td><input type="text" name="txtCiudadR" disabled value="<%=sCiudadR%>"></td>	
			</tr>
			<tr>
				<td>Teléfono</td>
				<td><input type="text" name="txtFonoR" disabled value="<%=sFonoR%>"></td>
			</tr>
			<tr>
				<td>Dirección</td>
				<td><input type="text" name="txtDireccionR" size="50" disabled value="<%=sDireccionR%>"></td>				
			</tr>
			<tr>
				<td>&nbsp;</td>
			</tr>
			<tr>
				<td>&nbsp;</td>
			</tr>
			-->
			<tr>
			  <td width="30%">
					<input type="radio" name="optCompra">
					<span class="textoempresa">Comprar US$				</span></td>
				<td>
					<input type="radio" name="optGiro">
					<span class="textoempresa">Enviar Giro Internacional				</span></td>
				<td>
					<input type="radio" name="optTransfer">
					<span class="textoempresa">Enviar Transferencia				</span></td>
			</tr>
			<tr>
			  <td>
					<input type="radio" name="optVenta">
					<span class="textoempresa">Vender US$				</span></td>
				<td>
					<input type="radio" name="optGiroNac">
					<span class="textoempresa">Enviar Giro Nacional				</span></td>					
				<td>
					<input type="radio" name="optTCredito">
					<span class="textoempresa">Pagar T. de Crédito en US$ </span></td>				
			</tr>
	  </table>
	</form>	
</BODY>
</HTML>