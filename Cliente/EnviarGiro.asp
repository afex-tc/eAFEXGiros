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
	Dim nAccion, sPais, sCiudad
	Dim sTarifa, sTotal
	Dim nMonto, nTarifa, nGastoTransfer, nComisionCaptador, nComisionPagador
	Dim nComisionMatriz, nAfectoIva, nDDIpais, nDDICiudad, nTotal
		
	On Error Resume Next

	sPais = Request("Pais")
	sCiudad = Request("Ciudad")
	nAccion = cInt(0 & Request("Accion"))
	nTarifa = 0
	Cargar
	
	Sub Cargar
		sPais = Request.Form("cbxPais")
		sCiudad = Request.Form("cbxCiudad")
		sPagador = "AL"
		sMoneda = request.Form("cbxMoneda")
		nMonto = cDbl(0 & Request.Form("txtMonto"))		
		nDDIPais = ObtenerDDI(1, sPais)
		nDDICiudad = ObtenerDDI(2, sCiudad)
		'sApellidos = Request.Form("txtApellidos")
		If Request.Form("txtMonto") <> "" Then
			nTarifa = ObtenerTarifa(cDbl(Request.Form("txtMonto")), Request.Form("cbxPais"), "***", Request.Form("cbxMoneda"))
			sTarifa = Formatnumber(nTarifa, 2)
			sTotal = Formatnumber(cDbl(Request.Form("txtMonto")) + cDbl(nTarifa), 2)
		End If
	End Sub
				
	Function ObtenerTarifa(ByVal Monto, Byval Pais, ByVal Ciudad, Byval Moneda)
		Dim afxGiro
		Dim nTarifa
		
		ObtenerTarifa = 0
		Set afxGiro = server.CreateObject("AfexGiroXP.Giro")
		
		On Error Resume Next
		nTarifa = afxGiro.ObtenerTarifa(session("afxCnxAFEXpress"), "AF", Monto, Pais, Ciudad, Moneda, Moneda)
		
		If Err.number <> 0 Then
			Set afxGiro = Nothing
			response.Redirect "Error.asp?Titulo=Error en Simulación de Giro&Number=" & Err.Number   & "&Source=" & Err.Source & "&Description=" & Err.description
		End If
		If afxGiro.ErrNumber <> 0 Then
			Set afxGiro = Nothing			
			response.Redirect "Error.asp?Titulo=Error en HágaseCliente&Number=" & afxGiro.ErrNumber  & "&Source=" & afxGiro.ErrSource & "&Description=" & replace(afxGiro.ErrDescription, vbCrLf , "^")	
		End If
		
		ObtenerTarifa = nTarifa
		Set afxGiro = Nothing
	End Function
	
%>
<!--#INCLUDE virtual="/Compartido/Rutinas.asp" -->
<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link rel="stylesheet" type="text/css" href="../Estilos/Cliente.css">
</head>
<script LANGUAGE="VBScript">
<!--

	Sub MostrarCotizacion()
		
		frmGiro.txtTarifa.value = FormatNumber("<%=nTarifa%>", 2)
		frmGiro.txtTotal.value = FormatNumber( cCur(cDbl(0 & frmGiro.txtTarifa.value)) + cCur(cDbl(0 & frmGiro.txtMonto.value)), 2)
		
	End Sub

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
		lblNombre.innerText = frmGiro.txtNombre.value & " " & frmGiro.txtApellido.value 
		lblMonto.innerText = frmGiro.txtMonto.value
		lblTarifa.innerText = frmGiro.txtTarifa.value
		lblTotal.innerText = frmGiro.txtTotal.value
		lblMoneda.innerText = frmGiro.cbxMoneda(frmGiro.cbxMoneda.selectedIndex).Text
		lblMoneda2.innerText = frmGiro.cbxMoneda(frmGiro.cbxMoneda.selectedIndex).Text
		lblMoneda3.innerText = frmGiro.cbxMoneda(frmGiro.cbxMoneda.selectedIndex).Text
		lblDireccion.innerText = frmGiro.txtDireccion.value 
		lblCiudad.innerText = frmGiro.cbxCiudad(frmGiro.cbxCiudad.selectedIndex).Text 
		lblPais.innerText = frmGiro.cbxPais(frmGiro.cbxPais.selectedIndex).Text 		
		lblTelefono.innerText = "(" & Trim(frmGiro.txtPaisFono.value) & Trim(frmGiro.txtAreaFono.value) & ") " & Trim(frmGiro.txtFono.value)
	End Sub

-->
</script>
<body>
<script LANGUAGE="VBScript">
<!--

	Const sEncabezadoFondo = "Transacciones"
	Const sEncabezadoTitulo = "Enviar un Giro"
	Const sClass = "TituloPrincipal"
		
	Sub cbxPais_onblur()
		If frmGiro.cbxPais.value = "" Then Exit Sub
		If frmGiro.cbxPais.value = "<%=sPais%>" Then Exit Sub			
		frmGiro.action = "EnviarGiro.asp?Accion=<%=afxAccionPais%>"
		frmGiro.submit 
		frmGiro.action = ""
	End Sub
	
	Sub cbxCiudad_onblur()
		If frmGiro.cbxCiudad.value = "" Then Exit Sub
		If frmGiro.cbxCiudad.value = "<%=sCiudad%>" Then Exit Sub			
		frmGiro.action = "EnviarGiro.asp?Accion=<%=afxAccionCiudad%>"
		frmGiro.submit 
		frmGiro.action = ""
	End Sub

	Sub txtMonto_OnKeyPress()
		IngresarTexto 1
	End Sub 

	Sub txtMonto_onBlur()
		Dim nPos
		
		If Trim(frmGiro.txtMonto.value) = "" Then
			frmGiro.txtMonto.value = "0"
		Else
			nPos = Instr(frmGiro.txtMonto.value, ",")
			If nPos > 0 Then
				If cCur(0 & Mid(frmGiro.txtMonto.value, nPos)) > 0 Then
					msgbox "El monto del giro no puede incluir decimales"
					frmGiro.txtMonto.focus
					frmGiro.txtMonto.select
					Exit Sub
				End If
			End If
		End If
		
		frmGiro.txtMonto.value = FormatNumber(frmGiro.txtMonto.value, 2)
		If cCur(frmGiro.txtMonto.value) <> cCur(0 & "<%=nMonto%>") Then
			HabilitarControles
			frmGiro.action = "EnviarGiro.asp?Accion=<%=afxAccionMonto%>"
			frmGiro.submit 
			frmGiro.action = ""
		End IF
	End Sub 

	Sub tdCalcular_onClick()
		'MostrarCotizacion	
		frmGiro.action = "EnviarGiro.asp?nAccion=<%=afxAccionPais%>"
		frmGiro.submit()
		frmGiro.action = ""
	End Sub 

	Sub window_onload()
		'frmGiro.txtMonto.value = "<%=Request.Form("txtMonto")%>"
		MostrarCotizacion()
		ActivarFoco
	End Sub	

	Sub tdAceptar_onClick()
		If Not ValidarDatos() Then
			Exit Sub
		End If
		HabilitarControles
		frmGiro.action = "GrabarEnvioGiro.asp"
		frmGiro.submit()
		frmGiro.action = ""
	End Sub 

	Sub ActivarFoco()
		<% Select Case nAccion %>
		<%	Case afxAccionPais %>
				frmGiro.cbxCiudad.focus 
		<%	Case afxAccionCiudad, afxAccionMonedaPago %>
				frmGiro.txtMonto.focus
				frmGiro.txtMonto.select
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
		If Trim(frmGiro.cbxPais.value) = "" Then
			MsgBox "Debe ingresar el pais donde enviará el dinero",,"AFEX"
			Exit Function
		End If
		If Trim(frmGiro.cbxCiudad.value) = "" Then
			MsgBox "Debe ingresar la ciudad donde enviará el dinero",,"AFEX"
			Exit Function
		End If
		If Trim(frmGiro.cbxMoneda.value) = "" Then
			MsgBox "Debe seleccionar la moneda de envío",,"AFEX"
			Exit Function
		End If
		If cCur(0 & Trim(frmGiro.txtMonto.value)) = 0 Then
			MsgBox "Debe ingresar el monto del giro",,"AFEX"
			Exit Function
		End If	
		ValidarPaso1 = True
	End Function
	
	Function ValidarPaso2()
		ValidarPaso2 = False
		If Trim(frmGiro.txtNombre.value) = "" Then
			MsgBox "Debe ingresar el nombre de quien recibirá el dinero",,"AFEX"
			Exit Function
		End If
		If Trim(frmGiro.txtApellido.value) = "" Then
			MsgBox "Debe ingresar el apellido de quien recibirá el dinero",,"AFEX"
			Exit Function
		End If
		If Trim(frmGiro.txtDireccion.value) = "" Then
			MsgBox "Debe ingresar la dirección de quien recibirá el dinero",,"AFEX"
			Exit Function
		End If
		If Trim(frmGiro.txtFono.value) = "" Then
			MsgBox "Debe ingresar el teléfono de quien recibirá el dinero",,"AFEX"
			Exit Function
		End If				
		If frmGiro.cbxPais.value = "CL" Then
			If frmGiro.cbxCiudad.value = "SCL" Then
				If Len(Trim(frmGiro.txtfono.value)) <> 7 Then
					Msgbox "El número de teléfono para Santiago de Chile debe ser de siete dígitos",, "AFEX"
					Exit Function
				End If
			Else
				If Len(Trim(frmGiro.txtfono.value)) <> 6 Then
					Msgbox "El número de teléfono para Regiones de Chile debe ser de seis dígitos",, "AFEX"
					Exit Function
				End If			
			End If
		End If
		ValidarPaso2 = True
	End Function

	Sub VerDemo_OnClick()
		window.open "DemoEnviarGiro.htm", null, _
				"title=no, top=0, left=0, height=405, width=570, status=no,toolbar=no,menubar=no,location=no"
	End Sub

	Sub optEnviarBoleta_onClick()
		window.frmGiro.optGuardarBoleta.checked=False		
	End Sub

	Sub optGuardarBoleta_onClick()
		window.frmGiro.optEnviarBoleta.checked=False		
	End Sub
	
-->
</script>
<!--#INCLUDE virtual="/Compartido/Encabezado.htm" -->
<table id="tabHagaseCliente" cellspacing="0" cellpadding="0" border="0" style="LEFT: 5px; POSITION: relative">
<tr height="20"><td>
      <object height="20" id="objTab" style="HEIGHT: 20px; WIDTH: 465px" type="text/x-scriptlet" VIEWASTEXT><param NAME="Scrollbar" VALUE="0"><param NAME="URL" VALUE="../Scriptlets/Tab.htm"></object>
</td></tr>
<tr ID="tabPaso1"><td>

<form id="frmGiro" method="post">
<table cellspacing="0" class="borde" ID="tabPaso11" border="0" style="HEIGHT: 280px; LEFT: 4px; POSITION: relative; TOP: 0px; WIDTH: 530px">
	<tr class="descripcion">
		<td colspan="5"><table><tr>
			<td WIDTH="5"></td>
			<td HEIGHT="50" COLSPAN="2">
				  <br>Seleccione el país y la ciudad donde desea enviar el dinero. Luego ingrese el monto y haga click en calcular.<br>Revise los datos. Si está conforme haga click en «siguiente» para ir al siguiente paso.<br><br>
			</td>	
			<td id="VerDemo"><img id="VerDemo" src="../images/BotonVerDemo.jpg" style="cursor: hand" WIDTH="80" HEIGHT="20"></td>
		</tr></table>
		</td>
	</tr>
	<tr HEIGHT="15">
		<td colspan="4" class="titulo">Datos del Envío</td>
	</tr>
	<tr HEIGHT="60">
		<td></td>
		<td>A qué país desea enviar el dinero?<br>
		<select NAME="cbxPais" style="width: 150px">
			<% CargarUbicacion 1, "", sPais %>
			</select></td>
		<td colspan="1">A qué ciudad desea enviar el dinero?<br>
		<select NAME="cbxCiudad" style="width: 150px">
			<% CargarCiudadesPais sPais, sCiudad %>
		</select></td>			
	</tr>
	<tr HEIGHT="60">
		<td></td>
			<td>En qué moneda desea enviar el giro?<br>
			<select NAME="cbxMoneda">
				<option SELECTED VALUE="USD">Dolar Americano</option>
			</select></td>
		<td colspan="1">Cuánto dinero desea enviar?<br>
			<table cellspacing="5" cellpadding="2">
			<tr><td>
				<input STYLE="TEXT-ALIGN: right" ID="txtMonto" NAME="txtMonto" SIZE="15" value="<%=Request.Form("txtMonto")%>">
				</td>
				<td id="tdCalcular">
					<img align="absmiddle" src="../images/BotonCalcular.jpg" style="cursor: hand" WIDTH="70" HEIGHT="20">
				</td>
			</tr>
			</table>
			</td>
		</tr>
	<tr HEIGHT="60" ID="tdTarifa" NAME="tdTarifa">
		<td></td>
		<td>Esta es la tarifa que debe pagar<br><strong><input NAME="txtTarifa" STYLE="FONT-SIZE: 10pt; HEIGHT: 20px; TEXT-ALIGN: right; WIDTH: 123px" value="<%=sTarifa%>" disabled></strong></td>
		<td colspan="2">Este es el total que debe pagar<br><strong><input NAME="txtTotal" STYLE="FONT-SIZE: 10pt; HEIGHT: 20px; TEXT-ALIGN: right; WIDTH: 123px" value="<%=sTotal%>" disabled></strong></td>
	</tr>
	<tr>
		<td></td>
		<td COLSPAN="4"></td>
	</tr>
	<tr>
		<td ></td>
		<td colspan="1" style="font-size: 7pt">*Nota: Los valores incluyen IVA</td>
		<td>
		<table cellspacing="5" cellpadding="2">
		<tr>
			<td width="100%"></td>
			<td id="tdsiguiente" onclick="MostrarPaso3('tabPaso2')">
				<img align="absmiddle" src="../images/Botonsiguiente.jpg" id="imgsiguiente" style="cursor: hand" WIDTH="80" HEIGHT="20">
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
		<td HEIGHT="50" COLSPAN="3">
			  <br>Ingrese los datos de quien recibirá el dinero. Luego haga click en «siguiente» para ir al siguiente paso.<br><br><br>
			</td>	
		</tr>
	<tr HEIGHT="15">
		<td colspan="4" class="titulo">Datos del Envío</td>
	</tr>
	<tr HEIGHT="60">
		<td></td>
		<td>Quién recibira el dinero?<br>Nombres<br><input NAME="txtNombre" SIZE="25" style="width: 300px" onkeypress="IngresarTexto(2)" onBlur="frmGiro.txtNombre.value=MayMin(frmGiro.txtNombre.value)" value="<%=request.Form("txtNombre")%>"></td>
		<td COLSPAN="2"><br>Apellidos<br><input NAME="txtApellido" SIZE="25" style="width: 200px" onkeypress="IngresarTexto(2)" onBlur="frmGiro.txtApellido.value=MayMin(frmGiro.txtApellido.value)" value="<%=request.Form("txtApellido")%>"></td>
	</tr>
	<tr HEIGHT="60">
		<td></td>
		<td COLSPAN="2">Dirección donde enviará el dinero?<br>
		<input NAME="txtDireccion" SIZE="50" style="width: 300px" onBlur="frmGiro.txtDireccion.value=MayMin(frmGiro.txtDireccion.value)" value="<%=request.Form("txtDireccion")%>"></td>
		<td colspan="1">Teléfono<br>
			<input disabled name="txtPaisFono" style="width: 40px" value="<%=nDDIPais%>">
			<input disabled name="txtAreaFono" style="width: 40px" value="<%=nDDICiudad%>">
			<input name="txtFono" style="width: 90px" size="10" onKeyPress="IngresarTexto(1)" value="<%=request.Form("txtFono")%>">
		</td>
	</tr>
	<tr HEIGHT="60">
		<td></td>
		<td COLSPAN="3">Si desea puede enviarle un mensaje aquí<br>
			<input name="txtMensaje" style="WIDTH: 526px" SIZE="256" value="<%=request.Form("txtMensaje")%>">
		</td>
	</tr>
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
</table></td></tr>

<tr ID="tabPaso3" style="DISPLAY: none"><td>
<!-- Paso 3 -->	
<table cellspacing="0" class="borde" ID="tabPaso33" border="0" style="HEIGHT: 280px; LEFT: 4px; POSITION: relative; TOP: 0px; WIDTH: 530px">
	<tr class="descripcion">
		<td WIDTH="5"></td>
		<td HEIGHT="50" COLSPAN="3">
		  <br>Lea atentamente el resumen de su operación. Si está seguro de la información y desea realizar la operación haga click en «aceptar». Si tiene dudas, o algún dato debe ser corregido, vuelva a los pasos anteriores para modificarlo.<br><br>
		</td>	
		</tr>
	<tr HEIGHT="15">
		<td colspan="4" class="titulo">Resumen</td>
	</tr>
	<tr HEIGHT="20" style="font-size: 10pt">
		<td></td>
		<td COLSPAN="2">
			<br>
			Usted desea enviar un giro por <strong ID="lblMonto" NAME="lblMonto"></strong>&nbsp;
			<strong ID="lblMoneda" NAME="lblMoneda"></strong>&nbsp;
			a <strong ID="lblNombre" NAME="lblNombre"></strong>&nbsp;
			que vive en <strong ID="lblDireccion" NAME="lblDireccion"></strong>,&nbsp; 
			<strong ID="lblCiudad" NAME="lblCiudad"></strong>,&nbsp; 
			<strong ID="lblPais" NAME="lblPais"></strong>,&nbsp;
			y se le dará aviso al teléfono&nbsp;<strong id="lblTelefono"></strong>.
			<br>				
			La tarifa cobrada es de <strong ID="lblTarifa" NAME="lblTarifa"></strong>&nbsp;
			<strong ID="lblMoneda2"></strong>,&nbsp; 
			y debe cancelar a AFEX la suma total de 
			<strong ID="lblTotal"></strong>&nbsp;
			<strong ID="lblMoneda3"></strong>.
		</td>
	</tr>
	<tr HEIGHT="20" style="font-size: 10pt">
		<td></td>
		<td COLSPAN="2">
			Si la Información es correcta haga click en Aceptar. <br>
			Si es necesario vuelva a los pasos correspondientes si desea modificar algún dato.
		</td>
	</tr>
	<tr HEIGHT="20">
		<td></td>
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
</html>
