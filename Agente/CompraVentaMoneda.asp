<%@ Language=VBScript %>
<%
	'Se asegura que la página no se almacene en la memoria cache
	Response.Expires = 0
	If Session("CodigoCliente") = "" Then
		response.Redirect "../Compartido/TimeOut.htm"
		response.end
	End If
%>
<!--#INCLUDE virtual="/Compartido/Constantes.asp" -->
<%
	Dim sMoneda
	
	sMoneda = Session("MonedaExtranjera")
	
%>
<!--#INCLUDE virtual="/Compartido/Transferencias.asp" -->
<!--#INCLUDE virtual="/Compartido/Rutinas.asp" -->
<!--#INCLUDE virtual="/Compartido/Errores.asp" -->
<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link rel="stylesheet" type="text/css" href="../Estilos/Cliente.css">
<style>
	INPUT
		{
			font-size:	8pt;	
		}
	SELECT
		{
			font-size:	8pt;	
		}
</style>
</head>
<script LANGUAGE="VBScript">
<!--
	Sub cbxOperacion_onKeyPress()
		cbxMoneda_onClick
		cbxMoneda1_onClick
		cbxMoneda2_onClick
	End Sub
	
	Sub cbxOperacion_onKeyDown()
		cbxMoneda_onClick
		cbxMoneda1_onClick
		cbxMoneda2_onClick
	End Sub

	Sub cbxOperacion_onKeyUp()
		cbxMoneda_onClick
		cbxMoneda1_onClick
		cbxMoneda2_onClick
	End Sub

	Sub cbxOperacion_onKeyPress()
		cbxMoneda_onClick		
		cbxMoneda1_onClick
		cbxMoneda2_onClick
	End Sub

	Sub cbxOperacion_onClick()
		cbxMoneda_onClick		
		cbxMoneda1_onClick
		cbxMoneda2_onClick
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
		MostrarSubTotal
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
		MostrarSubTotal
	End Sub

	Sub txtTipoCambio_onBlur()
		If Trim(frmCompraVenta.txtTipoCambio.value) = "" Then
			frmCompraVenta.txtTipoCambio.value = "0,0000"
		Else
			frmCompraVenta.txtTipoCambio.value = FormatNumber(frmCompraVenta.txtTipoCambio.value, 4)
		End If		
		MostrarSubTotal
	End Sub

	Sub MostrarSubTotal()
		frmCompraVenta.txtSubTotal.value = FormatNumber(cDbl(frmCompraVenta.txtMonto.value) * cDbl(frmCompraVenta.txtTipoCambio.value), 0)
		MostrarTotal
	End Sub

	Sub MostrarTotal()
		frmCompraVenta.txtTotal.value = FormatNumber(cDbl(frmCompraVenta.txtSubTotal.value) + cDbl(frmCompraVenta.txtSubTotal1.value) + cDbl(frmCompraVenta.txtSubTotal2.value), 0)		
	End Sub
	
	Sub cbxOperacion1_onKeyPress()
		cbxMoneda1_onClick		
	End Sub
	
	Sub cbxOperacion1_onKeyDown()
		cbxMoneda1_onClick
	End Sub

	Sub cbxOperacion1_onKeyUp()
		cbxMoneda1_onClick
	End Sub

	Sub cbxOperacion1_onKeyPress()
		cbxMoneda1_onClick		
	End Sub

	Sub cbxOperacion1_onClick()
		cbxMoneda1_onClick		
	End Sub
	
	Sub cbxMoneda1_onKeyDown()
		cbxMoneda1_onClick
	End Sub

	Sub cbxMoneda1_onKeyUp()
		cbxMoneda1_onClick
	End Sub

	Sub cbxMoneda1_onClick()
		If frmCompraVenta.cbxOperacion.value = 1 Then
			frmCompraVenta.txtTipoCambio1.value  = FormatNumber( _											
							frmCompraVenta.cbxTCCompra(frmCompraVenta.cbxMoneda1.selectedIndex).text _
							, 4)		
		Else
			frmCompraVenta.txtTipoCambio1.value  = FormatNumber( _											
							frmCompraVenta.cbxTCVenta(frmCompraVenta.cbxMoneda1.selectedIndex).text _
							, 4)		
		End IF
		MostrarSubTotal1
	End Sub
	
	Sub txtMonto1_OnKeyPress()
		 IngresarTexto 1
	End Sub
	
	Sub txtMonto1_onBlur()
		If Trim(frmCompraVenta.txtMonto1.value) = "" Then
			frmCompraVenta.txtMonto1.value = "0,00"
		Else
			frmCompraVenta.txtMonto1.value = FormatNumber(frmCompraVenta.txtMonto1.value, 2)
		End If		
		MostrarSubTotal1
	End Sub

	Sub txtTipoCambio1_onBlur()
		If Trim(frmCompraVenta.txtTipoCambio1.value) = "" Then
			frmCompraVenta.txtTipoCambio1.value = "0,0000"
		Else
			frmCompraVenta.txtTipoCambio1.value = FormatNumber(frmCompraVenta.txtTipoCambio1.value, 4)
		End If		
		MostrarSubTotal1
	End Sub

	Sub MostrarSubTotal1()
		frmCompraVenta.txtSubTotal1.value = FormatNumber(cDbl(frmCompraVenta.txtMonto1.value) * cDbl(frmCompraVenta.txtTipoCambio1.value), 0)
		MostrarTotal
	End Sub

	Sub cbxOperacion2_onKeyPress()
		cbxMoneda2_onClick
	End Sub
	
	Sub cbxOperacion2_onKeyDown()
		cbxMoneda2_onClick
	End Sub

	Sub cbxOperacion2_onKeyUp()
		cbxMoneda2_onClick
	End Sub

	Sub cbxOperacion2_onKeyPress()
		cbxMoneda2_onClick		
	End Sub

	Sub cbxOperacion2_onClick()
		cbxMoneda2_onClick
	End Sub
	
	Sub cbxMoneda2_onKeyDown()
		cbxMoneda2_onClick
	End Sub

	Sub cbxMoneda2_onKeyUp()
		cbxMoneda2_onClick
	End Sub

	Sub cbxMoneda2_onClick()
		If frmCompraVenta.cbxOperacion.value = 1 Then
			frmCompraVenta.txtTipoCambio2.value  = FormatNumber( _
							frmCompraVenta.cbxTCCompra(frmCompraVenta.cbxMoneda2.selectedIndex).text _
							, 4)
		Else
			frmCompraVenta.txtTipoCambio2.value  = FormatNumber( _
							frmCompraVenta.cbxTCVenta(frmCompraVenta.cbxMoneda2.selectedIndex).text _
							, 4)		
		End IF
		MostrarSubTotal2
	End Sub
	
	Sub txtMonto2_OnKeyPress()
		 IngresarTexto 1
	End Sub
	
	Sub txtMonto2_onBlur()
		If Trim(frmCompraVenta.txtMonto2.value) = "" Then
			frmCompraVenta.txtMonto2.value = "0,00"
		Else
			frmCompraVenta.txtMonto2.value = FormatNumber(frmCompraVenta.txtMonto2.value, 2)
		End If		
		MostrarSubTotal2
	End Sub

	Sub txtTipoCambio2_onBlur()
		If Trim(frmCompraVenta.txtTipoCambio2.value) = "" Then
			frmCompraVenta.txtTipoCambio2.value = "0,0000"
		Else
			frmCompraVenta.txtTipoCambio2.value = FormatNumber(frmCompraVenta.txtTipoCambio2.value, 4)
		End If		
		MostrarSubTotal2
	End Sub

	Sub MostrarSubTotal2()
		frmCompraVenta.txtSubTotal2.value = FormatNumber(cDbl(frmCompraVenta.txtMonto2.value) * cDbl(frmCompraVenta.txtTipoCambio2.value), 0)
		MostrarTotal 
	End Sub

	Sub window_onLoad()
		cbxMoneda_onClick
	End Sub	
	
	Sub tdAceptar_onClick()
		'If Not ValidarDatos() Then
		'	Exit Sub
		'End If		
		If Not CajaPregunta("AFEX En Linea", "Coloque la boleta en la impresora y haga click en Aceptar") Then						
			Exit Sub
		End If
		HabilitarControles
		frmCompraVenta.action = "GrabarCompraVenta.asp"
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

	Sub tdCalcular_onClick()		
	End Sub
	
	Sub ImprimirBoleta()
		Dim afxTM
		
		Set afxTM = CreateObject("AfexPrinter.TM295")

		afxTM.Inicializar
		afxTM.Habilitar MsComm1
		afxTM.EncabezadoBoleta MSComm1, 1002, Date, "Compra moneda", Time
		afxTM.DetalleBoleta MSComm1, frmCompraVenta.cbxMoneda.value, _
				frmCompraVenta.txtMonto.value, frmCompraVenta.txtTipoCambio.value, frmCompraVenta.txtSubTotal.value 
		afxTM.PieBoleta MSComm1, frmCompraVenta.txtTotal.value 
		afxTM.Deshabilitar MSComm1
				
		Set afxTM = Nothing
	End Sub
-->
</script>

<body style="font-size: 8pt">
<script LANGUAGE="VBScript">
<!--

	Const sEncabezadoFondo = "Transacciones"
	Const sEncabezadoTitulo = "Compra y Venta de Monedas"
	Const sClass = "TituloPrincipal"
	
-->
</script>
<!--#INCLUDE virtual="/Compartido/Encabezado.htm" -->
<!--
<OBJECT classid="clsid:648A5600-2C6E-101B-82B6-000000000014" id="MSComm1" style="LEFT: 0px; TOP: 0px" VIEWASTEXT>
	<PARAM NAME="_ExtentX" VALUE="1005">
	<PARAM NAME="_ExtentY" VALUE="1005">
	<PARAM NAME="_Version" VALUE="393216">
	<PARAM NAME="CommPort" VALUE="1">
	<PARAM NAME="DTREnable" VALUE="1">
	<PARAM NAME="Handshaking" VALUE="0">
	<PARAM NAME="InBufferSize" VALUE="512">
	<PARAM NAME="InputLen" VALUE="0">
	<PARAM NAME="NullDiscard" VALUE="0">
	<PARAM NAME="OutBufferSize" VALUE="512">
	<PARAM NAME="ParityReplace" VALUE="63">
	<PARAM NAME="RThreshold" VALUE="0">
	<PARAM NAME="RTSEnable" VALUE="0">
	<PARAM NAME="BaudRate" VALUE="9600">
	<PARAM NAME="ParitySetting" VALUE="0">
	<PARAM NAME="DataBits" VALUE="8">
	<PARAM NAME="StopBits" VALUE="0">
	<PARAM NAME="SThreshold" VALUE="0">
	<PARAM NAME="EOFEnable" VALUE="0">
	<PARAM NAME="InputMode" VALUE="0">
</OBJECT>
-->
<form id="frmCompraVenta" method="post">
<table class="borde" ID="tabPaso1" CELLSPACING="0" CELLPADDING="1" BORDER="0" WIDTH="560" STYLE="LEFT: 6px; POSITION: relative; TOP: 0px">

	<tr HEIGHT="15">
		<td colspan="6" height="15" bgcolor="steelblue"><font face="Verdana,Helvetica" color="white" size="2"><b>Datos de la operación</b></font></td>
	</tr>
	<tr HEIGHT="10">		
		<td COLSPAN="4"></td>		
	</tr>
	<tr>
		<td width></td>
		<td VALIGN="center" colspan>Operacion&nbsp;<br>
			<select NAME="cbxOperacion" style="HEIGHT: 22px; WIDTH: 80px">
				<option SELECTED VALUE="1">Comprar</option>
				<option VALUE="2">Vender</option>
			</select><br>
		</td>
	</tr>
	<strong>
	<tr ALIGN="center">
		<td></td>
		<td>Moneda</td>
		<td>Producto</td>
		<td>Monto</td>
		<td>Tipo de<br>Cambio</td>
		<td>Total</td>
	</tr>
	</strong>
	<tr>
		<td></td>
		<td>
			<select NAME="cbxMoneda" style="HEIGHT: 22px; WIDTH: 160px">
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
		<td VALIGN="center">
			<select NAME="cbxProducto" style="HEIGHT: 22px; WIDTH: 80px">
				<option SELECTED VALUE="1">Efectivo</option>
				<option VALUE="2">Cheque</option>
				<option value="3">Transferencia</option>
			</select>
		</td>
		<td>
			<input STYLE="HEIGHT: 22px; TEXT-ALIGN: right; WIDTH: 100px" SIZE="10" ALIGN="rigth" NAME="txtMonto" VALUE="0,00" onkeypress="IngresarTexto(1)">
		</td>
		<td>
			<input STYLE="HEIGHT: 22px; TEXT-ALIGN: right; WIDTH: 80px" SIZE="10" ALIGN="right" NAME="txtTipoCambio" VALUE="0,00" onkeypress="IngresarTexto(1)">
		</td>
		<td>
			<input DISABLED STYLE="BACKGROUND: lightgrey; FONT-SIZE: 8pt; FONT-WEIGHT: bold; HEIGHT: 22px; TEXT-ALIGN: right; WIDTH: 100px; FONT-COLOR: black" SIZE="10" ALIGN="rigth" NAME="txtSubTotal" VALUE="0">
		</td>
	</tr>
	<tr>
		<td></td>
		<td>
			<select NAME="cbxMoneda1" style="HEIGHT: 22px; WIDTH: 160px">
			<%
				CargarMonedas sMoneda
			%>
			</select>							
		</td>
		<td VALIGN="center">
			<select NAME="cbxProducto1" style="HEIGHT: 22px; WIDTH: 80px">
				<option SELECTED VALUE="1">Efectivo</option>
				<option VALUE="2">Cheque</option>
				<option value="3">Transferencia</option>
			</select>
		</td>
		<td>
			<input STYLE="HEIGHT: 22px; TEXT-ALIGN: right; WIDTH: 100px" SIZE="10" ALIGN="rigth" NAME="txtMonto1" VALUE="0,00" onkeypress="IngresarTexto(1)">
		</td>
		<td>
			<input STYLE="HEIGHT: 22px; TEXT-ALIGN: right; WIDTH: 80px" SIZE="10" ALIGN="right" NAME="txtTipoCambio1" VALUE="0,00" onkeypress="IngresarTexto(1)">
		</td>
		<td>
			<input DISABLED STYLE="BACKGROUND: lightgrey; FONT-SIZE: 8pt; FONT-WEIGHT: bold; HEIGHT: 22px; TEXT-ALIGN: right; WIDTH: 100px; FONT-COLOR: black" SIZE="10" ALIGN="rigth" NAME="txtSubTotal1" VALUE="0">
		</td>
	</tr>
	<tr>
		<td></td>
		<td>
			<select NAME="cbxMoneda2" style="HEIGHT: 22px; WIDTH: 160px">
			<%
				CargarMonedas sMoneda
			%>
			</select>							
		</td>
		<td VALIGN="center">
			<select NAME="cbxProducto2" style="HEIGHT: 22px; WIDTH: 80px">
				<option SELECTED VALUE="1">Efectivo</option>
				<option VALUE="2">Cheque</option>
				<option value="3">Transferencia</option>
			</select>
		</td>
		<td>
			<input STYLE="HEIGHT: 22px; TEXT-ALIGN: right; WIDTH: 100px" SIZE="10" ALIGN="rigth" NAME="txtMonto2" VALUE="0,00" onkeypress="IngresarTexto(1)">
		</td>
		<td>
			<input STYLE="HEIGHT: 22px; TEXT-ALIGN: right; WIDTH: 80px" SIZE="10" ALIGN="right" NAME="txtTipoCambio2" VALUE="0,00" onkeypress="IngresarTexto(1)">
		</td>
		<td>
			<input DISABLED STYLE="BACKGROUND: lightgrey; FONT-SIZE: 8pt; FONT-WEIGHT: bold; HEIGHT: 22px; TEXT-ALIGN: right; WIDTH: 100px; FONT-COLOR: black" SIZE="10" ALIGN="rigth" NAME="txtSubTotal2" VALUE="0">
		</td>
	</tr>
	<tr>
		<td colspan="4"></td>		
		<td align="right">
			<img border="0" id="imgCalcular" name="imgCalcular" onClick="MostrarTotal()" src="../images/BotonCalcular.jpg" style="LEFT: 0px; POSITION: relative; TOP: 5px; cursor: hand" WIDTH="70" HEIGHT="20">&nbsp;Total
		</td>
		<td>
			<input DISABLED STYLE="BACKGROUND: lightgrey; FONT-SIZE: 8pt; FONT-WEIGHT: bold; HEIGHT: 22px; TEXT-ALIGN: right; WIDTH: 100px; FONT-COLOR: black" SIZE="10" ALIGN="rigth" NAME="txtTotal" VALUE="0">
		</td>
	</tr>
	<tr HEIGHT="40">
		<td></td>		
		<td COLSPAN="5" align="right">
			<img border="0" id="tdAceptar" onclick src="../images/BotonAceptar.jpg" style="POSITION: relative; cursor: hand" WIDTH="70" HEIGHT="20">&nbsp;&nbsp;&nbsp;
		</td>
	</tr>
	
</table>
</form>
<!--#INCLUDE virtual="/Compartido/Rutinas.htm" -->
</body>
</html>
