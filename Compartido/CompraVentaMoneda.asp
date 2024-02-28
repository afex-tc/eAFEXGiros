<%@ Language=VBScript %>
<!--#INCLUDE virtual="/compartido/TimeOut.asp" -->
<!--#INCLUDE virtual="/compartido/Constantes.asp" -->
<%
	Dim sMoneda, nTipo
	
	sMoneda = Session("MonedaExtranjera")
	nTipo = cInt(0 & Request("Tipo")) '1 Cliente; 2 Agente
	If nTipo = afxCliente Then sHabilitado = "disabled"
	
%>
<!--#INCLUDE virtual="/compartido/Transferencias.asp" -->
<!--#INCLUDE virtual="/compartido/Rutinas.asp" -->
<!--#INCLUDE virtual="/compartido/Errores.asp" -->
<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link rel="stylesheet" type="text/css" href="../Estilos/Cliente.css">
</head>
<script LANGUAGE="VBScript">
<!--
	Sub cbxOperacion_onKeyPress()
		cbxMoneda_onClick		
	End Sub
	
	Sub cbxOperacion_onKeyDown()
		cbxMoneda_onClick
	End Sub

	Sub cbxOperacion_onKeyUp()
		cbxMoneda_onClick
	End Sub

	Sub cbxOperacion_onKeyPress()
		cbxMoneda_onClick		
	End Sub

	Sub cbxOperacion_onClick()
		cbxMoneda_onClick		
	End Sub
	
	Sub cbxMoneda_onKeyDown()
		cbxMoneda_onClick
	End Sub

	Sub cbxMoneda_onKeyUp()
		cbxMoneda_onClick
	End Sub

	Sub cbxMoneda_onClick()
		If frmCompraVenta.cbxOperacion.value = 1 Then
			frmCompraVenta.txtTipoCambio.value  = FormatNumber( _											
							frmCompraVenta.cbxTCCompra(frmCompraVenta.cbxMoneda.selectedIndex).text _
							, 4)		
		Else
			frmCompraVenta.txtTipoCambio.value  = FormatNumber( _											
							frmCompraVenta.cbxTCVenta(frmCompraVenta.cbxMoneda.selectedIndex).text _
							, 4)		
		End IF
		MostrarTotal
	End Sub
	
	Sub txtMonto_OnKeyPress()
		 IngresarTexto 1
	End Sub
	
	Sub txtMonto_onBlur()
		If Trim(frmCompraVenta.txtMonto.value) = "" Then
			frmCompraVenta.txtMonto.value = "0,00"
		Else
			frmCompraVenta.txtMonto.value = FormatNumber(frmCompraVenta.txtMonto.value, 2)
		End If		
		MostrarTotal
	End Sub
		
	Sub imgCalcular_onClick()
		MostrarTotal
	End Sub 

	Sub MostrarTotal()
		frmCompraVenta.txtTotal.value = FormatNumber(cDbl(frmCompraVenta.txtMonto.value) * cDbl(frmCompraVenta.txtTipoCambio.value), 0)
	End Sub

	Sub window_onLoad()
		cbxMoneda_onClick
	End Sub	
	
	Sub tdAceptar_onClick()
		If Not ValidarDatos() Then
			Exit Sub
		End If		
		HabilitarControles
		<% If nTipo = afxCliente Then %>
				frmCompraVenta.action = "GrabarCompraVenta.asp"
		<% Else %>
				
		<% End If %>
		frmCompraVenta.submit()
		frmCompraVenta.action = ""
	End Sub 

	Function ValidarDatos()
		ValidarDatos = False
		If Trim(frmCompraVenta.cbxOperacion.value) = "" Then
			MsgBox "Debe seleccionar el tipo de operación que desea",,"AFEX"
			Exit Function
		End If
		If Trim(frmCompraVenta.cbxMoneda.value) = "" Then
			MsgBox "Debe seleccionar la moneda de la operación",,"AFEX"
			Exit Function
		End If
		If cCur(0 & Trim(frmCompraVenta.txtMonto.value)) = 0 Then
			MsgBox "Debe ingresar el monto de la operación",,"AFEX"
			Exit Function
		End If	
		ValidarDatos = True
	End Function

	Sub optEnviarBoleta_onClick()
		window.frmCompraVenta.optGuardarBoleta.checked=False		
	End Sub

	Sub optGuardarBoleta_onClick()
		window.frmCompraVenta.optEnviarBoleta.checked=False		
	End Sub

-->
</script>

<body>
<script LANGUAGE="VBScript">
<!--

	Const sEncabezadoFondo = "Transacciones"
	Const sEncabezadoTitulo = "Compra y Venta de Monedas"
	Const sClass = "TituloPrincipal"
	
-->
</script>
<!--#INCLUDE virtual="/Compartido/Encabezado.htm" -->
<form id="frmCompraVenta" method="post">
<table class="borde" ID="tabPaso1" CELLSPACING="0" CELLPADDING="1" BORDER="0" WIDTH="560" STYLE="LEFT: 2px; POSITION: relative; TOP: 0px">
	<tr bgcolor="#dcf9ff">
		<td WIDTH="4"></td>
		<td HEIGHT="80" COLSPAN="6">
			 Seleccione el tipo de operación que desea realizar, la moneda y el producto. Luego ingrese el monto y si lo requiere puede modificar el tipo de cambio.
			 
		</td>	
	</tr>
	<tr HEIGHT="15">
		<td colspan="4" height="15" bgcolor="steelblue"><font face="Verdana,Helvetica" color="white" size="2"><b>Datos de la operación</b></font></td>
	</tr>
	<tr HEIGHT="20">
		<td COLSPAN="4"></td>		
	</tr>
	<tr HEIGHT="40">
		<td></td>
		<td VALIGN="center" WIDTH="170">Qué operación desea hacer?<br>
			<select NAME="cbxOperacion" style="HEIGHT: 22px; WIDTH: 124px">
				<option SELECTED VALUE="2">Comprar</option>
				<option VALUE="1">Vender</option>
			</select>
		</td>
		<td COLSPAN="2">Seleccione la moneda<br>
			<select NAME="cbxMoneda" style="HEIGHT: 22px; WIDTH: 180px">
			<%
				CargarMonedas sMoneda
			%>
			</select>							
			<select name="cbxTCCompra" STYLE="Display: none">	
			<%
				CargarParidades afxTCCompra, sMoneda
			%>
			</select>
			<select name="cbxTCVenta" STYLE="Display: none">	
			<%
				CargarParidades afxTCVenta, sMoneda
			%>
			</select>
		</td>
	</tr>
	<tr>
		<td></td>
		<td VALIGN="center">Tipo de Cambio<br>
			<input STYLE="HEIGHT: 22px; TEXT-ALIGN: right; WIDTH: 123px" SIZE="10" ALIGN="right" NAME="txtTipoCambio" VALUE="0,00" disabled>
		</td>
		<td VALIGN="center" COLSPAN="2">Monto de la operación<br>
			<input STYLE="HEIGHT: 22px; TEXT-ALIGN: right; WIDTH: 137px" SIZE="10" ALIGN="rigth" NAME="txtMonto" VALUE="0,00">
			<img border="0" height="20" id="imgCalcular" name="imgCalcular" src="../images/BotonCalcular.jpg" style="LEFT: 0px; POSITION: relative; TOP: 5px; cursor: hand" width="70">
		</td>
	</tr>
	<tr HEIGHT="40">
		<td></td>
		<td COLSPAN="3" VALIGN="center"><strong>Total de la operación</strong><br>
			<input DISABLED STYLE="BACKGROUND: lightgrey; FONT-SIZE: 12pt; FONT-WEIGHT: bold; HEIGHT: 22px; TEXT-ALIGN: right; WIDTH: 159px; FONT-COLOR: black" SIZE="10" ALIGN="rigth" NAME="txtTotal" VALUE="0"><br><br>
		</td>
	</tr>
	<tr>
		<td></td>		
		<td>
			<input type="radio" name="optEnviarBoleta" checked>Enviar boleta a domicilio<br>
			<input type="radio" name="optGuardarBoleta">Guardar boleta en AFEX (6 meses)<br><br>
		</td>
		
	</tr>
	<tr HEIGHT="40">
		<td COLSPAN="3"></td>		
		<td align="right"><img border="0" id="tdAceptar" onclick src="../images/BotonAceptar.jpg" style="POSITION: relative; cursor: hand" WIDTH="70" HEIGHT="20"></td>
	</tr>
	
</table>
</form>
<!--#INCLUDE virtual="/compartido/Rutinas.htm" -->
</body>
</html>
