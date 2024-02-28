<%@ Language=VBScript %>
<%

	Sub CargarTipoGiro()
		Dim sSelect
	
		On Error Resume Next
		Response.write "<option value=0>NACIONALES</option>"
		Response.write "<option value=1>INTERNACIONALES</option>"

	End Sub

%>
<!--#INCLUDE virtual="/Compartido/Rutinas.asp" -->
<!--#INCLUDE virtual="/Sucursal/Rutinas.asp" -->
<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<title>Configuración Consulta de Giros</title>
<link rel="stylesheet" type="text/css" href="../Estilos/Cliente.css">
</head>

<script LANGUAGE="VBScript">
<!--
	On Error Resume Next
	Dim nTipo

	Sub public_put_Tipo(iTipoConsulta)
		Dim objTemp
		Dim sTr
		Dim i

		nTipo = iTipoConculta
		Select Case iTipoConsulta			
			Case <%=afxGirosPendientes%>
				sTr= ""
			
			Case <%=afxGirosCartola%>
				sTr = ""				
				
			Case <%=afxGirosRecibidos%>
				sTr = "Giro"

				If <%=Session("Categoria")%> <> 4 Then				
					trCaptador.style.display = ""
				End If
				
			Case <%=afxGirosEnviados%>
				sTr = "Giro"

				If <%=Session("Categoria")%> <> 4 Then					
					trPagador.style.display = ""
				End If
				
			Case 7	'Consulta de Transferencias
				sTr = "Transfer"
			
			Case 9, 10
				sTr = ""
				
			Case 11
				sTr = "Tarjeta"
				'trTarjeta.style.display = ""
				
		End Select		
		Set objTemp = document.all.tags("TR")
		For i = 0 To objTemp.length - 1
			If objTemp(i).title = sTr Then objTemp(i).style.display = ""
		Next 
		Set objTemp = document.all.tags("TD")
		For i = 0 To objTemp.length - 1
			If objTemp(i).title = sTr Then objTemp(i).style.display = ""
		Next 

	End Sub
	
	Function public_get_Tipo()
		public_get_Tipo = nTipo
	End Function

	Sub imgAceptar_onClick()
		'window.navigate "ListaGiros.asp?Tipo=1&Titulo=" + "<%=sTitulo%>"
		window.event.returnValue = False
		window.external.raiseEvent "Aceptar", "Aceptar"
	End Sub		

	Function public_put_Captador(sAgente)
		cbxCaptador.value = sAgente
		Select Case sAgente
		Case ""
			cbxCaptador.selectedIndex = 0
		Case "MB"
			cbxCaptador.selectedIndex = 1
		Case "MG"
			cbxCaptador.selectedIndex = 2
		Case "UT"
			cbxCaptador.selectedIndex = 4
		Case "*"
			cbxCaptador.selectedIndex = 3
		End Select
	End Function
	
	Function public_get_Captador()
		public_get_Captador = cbxCaptador.value 
	End Function

	Function public_put_Sucursal(sSucursal)
		cbxSucursal.value = sSucursal
		If sSucursal = "AW" Then
			cbxSucursal.disabled = False
		Else
			cbxSucursal.disabled = True
		End If
	End Function
	
	Function public_get_Sucursal()
		public_get_Sucursal = cbxSucursal.value 
		If cbxSucursal.value = "AW" Then
			cbxSucursal.disabled = False
		Else
			cbxSucursal.disabled = True
		End If
	End Function

	Function public_put_Pagador(sAgente)
		cbxPagador.value = sAgente
		Select Case sAgente
		Case ""
			cbxPagador.selectedIndex = 0
		Case "ME"
			cbxPagador.selectedIndex = 1
		Case "UT"
			cbxPagador.selectedIndex = 3
		Case "*"
			cbxPagador.selectedIndex = 2
		End Select
	End Function
	
	Function public_get_Pagador()
		public_get_Pagador = cbxPagador.value 
	End Function

	Function public_put_CodigoCliente(sCliente)
		txtCodigoCliente.value =  sCliente
		If sCliente <> "" Then
			txtNombres.disabled = True
			txtApellidos.disabled = True
		End If
	End Function

	Function public_get_CodigoCliente()
		public_get_CodigoCliente = txtCodigoCliente.value 
	End Function

	Function public_put_NombreCliente(sCliente)
		txtNombres.value =  sCliente
	End Function

	Function public_get_NombreCliente()
		public_get_NombreCliente = txtNombres.value 
	End Function

	Function public_put_ApellidoCliente(sCliente)
		txtApellidos.value =  sCliente
	End Function
	
	Function public_get_ApellidoCliente()
		public_get_ApellidoCliente = txtApellidos.value 
	End Function

	Function public_put_Desde(sDesde)
		txtDesde.value =  sDesde
	End Function

	Function public_get_Desde()
		public_get_Desde = txtDesde.value 
	End Function

	Function public_put_Hasta(sHasta)
		txtHasta.value =  sHasta
	End Function

	Function public_get_Hasta()
		public_get_Hasta = txtHasta.value 
	End Function

	Sub txtDesde_onBlur()
		Dim sFecha
		sFecha = txtDesde.value
		If ValidarFechas(sFecha) Then
			txtDesde.value = sFecha
		Else
			txtDesde.focus 
			txtDesde.select 
		End If
	End Sub

	Function public_put_Moneda(sMoneda)
		Dim i		
		If sMoneda <> Empty Then
			For i = 0 To cbxMoneda.length - 1
				cbxMoneda.selectedIndex = i
				If Ucase(cbxMoneda.value) = Ucase(sMoneda) Then
					Exit For	
				End If
			Next
		End If
	End Function

	Function public_get_Moneda()
		public_get_Moneda = cbxMoneda.value 
	End Function

	Function public_put_TipoGiro(Byval Tipo)
		cbxTipoGiro.selectedIndex = cInt(0 & Tipo)
	End Function

	Function public_get_TipoGiro()
		public_get_TipoGiro = cbxTipoGiro.value 
	End Function

	Sub txtHasta_onBlur()
		Dim sFecha

		sFecha = txtHasta.value
		If ValidarFechas(sFecha) Then
			txtHasta.value = sFecha
		Else
			txtHasta.focus 
			txtHasta.select 
		End If
	End Sub
	
	Function ValidarFecha(ByRef Fecha)
		Fecha = UCase(Trim(Fecha))
		Fecha = Replace(Fecha, "-", "")
		Fecha = Replace(Fecha, "/", "")
		ValidarFecha = True		

		Select Case Fecha
		Case "" 
			Fecha = Date()
			Exit Function
		Case "HOY"
			fecha = Date()
			Exit Function
		Case "AYER"
			Fecha = Date() - 1
			Exit Function
		End Select

		ValidarFecha = False
		If Len(Fecha) < 6 Then
			MsgBox "Debe ingresar una fecha válida con uno de estos formatos: " & vbCrLf & vbCrLf & _
					 "dd-mm-aaaa, dd-mm-aa, ddmmaa, HOY o AYER"
			Exit Function
		End If
		If Len(Fecha) = 7 Then
			MsgBox "Debe ingresar una fecha válida con uno de estos formatos: " & vbCrLf & vbCrLf & _
					 "dd-mm-aaaa, dd-mm-aa, ddmmaa, HOY o AYER"
			Exit Function
		End If
		If Len(Fecha) > 8 Then
			MsgBox "Debe ingresar una fecha válida con uno de estos formatos: " & vbCrLf & vbCrLf & _
					 "dd-mm-aaaa, dd-mm-aa, ddmmaa, HOY o AYER"
			Exit Function
		End If
		If Len(Fecha) = 6 Then
			Fecha = Left(Fecha, 2) & "-" & Mid(fecha, 3, 2) & "-"  & Right(Fecha, 2)
			Fecha = cDate(fecha)
		ElseIf Len(Fecha) = 8 Then
			Fecha = Left(Fecha, 2) & "-" & Mid(fecha, 3, 2) & "-" & Right(Fecha, 4)
			Fecha = cDate(fecha)
		End If
		ValidarFecha = True
	End Function
	
	Function ValidarFechas(ByRef Fecha)
		Fecha = UCase(Trim(Fecha))
		Fecha = Replace(Fecha, "-", "")
		Fecha = Replace(Fecha, "/", "")
		ValidarFechas = True		

		Select Case Fecha
		Case "" 
			Fecha = Date()
			Exit Function
		Case "HOY"
			fecha = Date()
			Exit Function
		Case "AYER"
			Fecha = Date() - 1
			Exit Function
		End Select

		ValidarFechas = False
		If Len(Fecha) < 6 Then
			MsgBox "Debes ingresar una fecha válida con uno de estos formatos: " & vbCrLf & vbCrLf & _
					 "dd-mm-aaaa, dd-mm-aa, ddmmaa, HOY o AYER"
			Exit Function
		End If
		If Len(Fecha) = 7 Then
			MsgBox "Debes ingresar una fecha válida con uno de estos formatos: " & vbCrLf & vbCrLf & _
					 "dd-mm-aaaa, dd-mm-aa, ddmmaa, HOY o AYER"
			Exit Function
		End If
		If Len(Fecha) > 8 Then
			MsgBox "Debes ingresar una fecha válida con uno de estos formatos: " & vbCrLf & vbCrLf & _
					 "dd-mm-aaaa, dd-mm-aa, ddmmaa, HOY o AYER"
			Exit Function
		End If
		If Len(Fecha) = 6 Then
			Fecha = Left(Fecha, 2) & "-" & Mid(fecha, 3, 2) & "-"  & Left(Trim(cStr(Year(Date))), 2) & Right(Fecha, 2)
		ElseIf Len(Fecha) = 8 Then
			Fecha = Left(Fecha, 2) & "-" & Mid(fecha, 3, 2) & "-" & Right(Fecha, 4)
		End If
		ValidarFechas = True
	End Function
		
//-->
</script>
<body>
<center>
<input type="hidden" id="txtCodigoCliente">
<table id="tabConsulta" class="borde" BORDER="0" cellpadding="4" cellspacing="0" style="HEIGHT: 195px" width="300px">
<tr><td class="Titulo" colspan="2" style="FONT-SIZE: 10pt; HEIGHT: 5px">Datos de la consulta</td></tr>
<tr id="trAgente" title="Giro" style="DISPLAY: none">
	<td colspan="2">
		<table width="100%" cellpadding="0" cellspacing="0">
		<tr id="trCaptador" style="display: none">
			<td>Agente captador<br>
				<select name="cbxCaptador" style="HEIGHT: 22px; WIDTH: 400px">					
					<option value="">Agentes AFEX</option>
					<option value="MB">Money Gram</option>
					<option value="MG">MoneyBroker</option>
					<!--<option value="UT">Uniteller</option> APPL-6080_MS_26-09-2014-->
					<option value="*">Todos los agentes</option>
				</select>
			</td>
		</tr>
		<tr id="trPagador" style="display: none">
			<td>Agente pagador<br>
				<select name="cbxPagador" style="HEIGHT: 22px; WIDTH: 400px">
					<option value="">Agentes AFEX</option>
					<option value="ME">Money Gram</option>
					<!--<option value="UT">Uniteller</option>APPL-6080_MS_26-09-2014-->
					<option value="*">Todos los agentes</option>
				</select>
			</td>
		</tr>
		<tr id="trMoneda" style="display: ">		
		<td>
			<table>
			<tr>
			<td>Giros<br>
				<select name="cbxTipoGiro" style="width: 200px">
					<%	
						CargarTipoGiro
					%>
				</select>
			</td>
			<td>Moneda<br>
				<select name="cbxMoneda" style="width: 200px">
					<%	
						CargarMonedaGiro "", Session("PaisMatriz"), "", ""
					%>
				</select>
			</td>
			</tr>
			</table>
		</td>
		</tr>
		</table>
	</td>
</tr>
<tr id="trTarjeta" title="Tarjeta" style="DISPLAY: none">
	<td>
		<table width="100%" cellpadding="0" cellspacing="0">
		<tr>
			<td>Sucursal<br>
				<select name="cbxSucursal" style="HEIGHT: 22px; WIDTH: 290px">
					<% CargarSucursal sSucursal %>
				</select>
			</td>
		</tr>
		</table>
	</td>
</tr>
<!--<tr id="trBanco" title="Transfer" style="DISPLAY: none"><td colspan="2">		<table width="100%" cellpadding="0" cellspacing="0">		<tr>			<td>Banco origen<br>				<select name="cbxBancoOrigen" style="HEIGHT: 22px; WIDTH: 199px" disabled>					<option value="99" selected>Todos los bancos</option>					<option value="1">Bank Of America</option>					<option value="2">Deutsche Bank</option>				</select>			</td>			<td>Banco destino<br><input SIZE="40" VALUE="Todos los bancos" id="txtBancoDestino" style="HEIGHT: 22px; WIDTH: 209px"></td>		</tr>		</table></td></tr>-->
<tr align="center"><td title="Giro" id="Giro" style="display: none">
		<table>
		<tr><td id="tdNombres">Nombres<br><input SIZE="30" VALUE id="txtNombres"></td></tr>
		<tr><td id="tdApellidos">Apellidos<br><input SIZE="30" VALUE id="txtApellidos"></td></tr>
<!--		
		<tr style="display: none">
			<td>Estado<br>
			<select name="cbxEstado" disabled>
				<option value="99" selected>Todos</option>
				<option value="1">Enviado</option>
				<option value="2">Recibido</option>
			</select>
			</td>
		</tr>
-->
		</table>
	</td>
	<td>	
		<table id="tabPeriodo" class="bordeinactivo" cellspacing="0" cellpadding="3">
        <tbody>
		<tr>
			<td colspan="2" class="tituloinactivo">Periodo&nbsp;&nbsp;<a style="font-size: 8pt">(ddmmyy)</a></td>
		</tr>
		<tr>
			<td>Desde</td> 
			<td><input SIZE="8" VALUE="01-01-2002" id="txtDesde"></td>
		</tr>
		<tr>
			<td>Hasta</td>
			<td><input SIZE="8" VALUE="01-01-2002" id="txtHasta"></td>
			</td>
		</tr>
		</table></td></tr>
<tr align="middle">
	<td colspan="2"><img id="imgAceptar" src="../images/BotonAceptar.jpg" style="CURSOR: hand" WIDTH="70" HEIGHT="20"></td>
</tr></tbody></table>
</center>
<!--#INCLUDE virtual="/Compartido/Rutinas.htm" -->
</body>
</html>