<%@ Language=VBScript %>
<%
	Response.Expires = 0
	Response.buffer = true	
	If Not Session("SesionActiva") Then
'		response.Redirect "../Compartido/TimeOut.htm"
		response.end
	End If
%>
<%
	
%>
<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<title>Paridades <%=Request("cmn")%></title>
<link rel="stylesheet" type="text/css" href="../Compartido/VisorAP.css">
</head>
<!--#INCLUDE virtual="/afexvisorap/visorap/Rutinas.htm" -->
<!--#INCLUDE virtual="/afexvisorap/Compartido/Boton.htm" -->
<script LANGUAGE="VBScript">
<!--
	Dim bCargado
	Dim nUSDCompra, nUSDVenta
	
	bCargado = false
	Sub window_onload()
		HabilitarPrecios
		Cargar
		If bCargado Then CalcularPrecios
	End Sub
	
	Sub Cargar
		Dim rs
		
		On Error Resume Next
		Set rs = ObtenerVisorAP("<%=Session("cnxVisorAP")%>", "<%=Request("mn")%>", "<%=Date()%>")
		If Err.number <> 0 Then
			Err.Clear 
			Set rs = Nothing
			Exit Sub
		End If
		
		bCargado = True
		
		tdMoneda.innerText  = rs("alias_moneda")
		txtTCCompra.value = FormatNumber(rs("pm_compra"), 2)
		txtTCVenta.value= FormatNumber(rs("pm_venta"), 2)
		txtParidadCompra.value = rs("paridad_compra")
		txtParidadVenta.value = rs("paridad_venta")
		txtParidadTrf.value = rs("paridad_transferencia")
		
		Set rs = ObtenerVisorAP("<%=Session("cnxVisorAP")%>", "USD")
		If Err.number <> 0 Then
			Err.Clear 
			Set rs = Nothing
			Exit Sub
		End If
	   nUSDCompra = CCur(rs("pr_compra"))
		nUSDVenta = CCur(rs("pr_venta"))		
		Set rs = Nothing
	End Sub

	Sub CalcularPrecios()
		If txtParidadCompra.value = 0 Then txtParidadCompra.value = 1
		If txtParidadVenta.value = 0 Then txtParidadVenta.value = 1
		If txtParidadtrf.value = 0 Then txtParidadTrf.value = 1
	   txtTCCompra.value = FormatNumber(nUSDCompra / CCur(0 & txtParidadCompra.value ), 2)
	   txtTCVenta.value = FormatNumber(nUSDVenta / CCur(0 & txtParidadVenta.value ), 2)
	   txtTCTrf.value = FormatNumber(nUSDVenta / CCur(0 & (1 / txtParidadTrf.value)), 4)
	End Sub

	Sub AplicarPR
		AgregarPR
	End Sub
	
	Sub AceptarPR
		AgregarPR
		window.close 
	End Sub
	
	Sub AgregarPR
		Dim bOk
		
		If MsgBox("Está seguro que desea guardar los cambios?", vbYesNo+vbQuestion) <> vbYes Then Exit Sub
		
		bOk = AgregarPRPro("<%=Session("cnxVisorAP")%>", "<%=Request("mn")%>", _
								cCur(0 & txtTCCompra.value), cCur(0 & txtTCVenta.value), _
								"ADMIN", cCur(0 & txtParidadCompra.value), cCur(0 & txtParidadVenta.value), _
								cDbl(0 & txtParidadTrf.value))
		If Not bOk Then
			MsgBox "Se produjo un error al intentar guardar los cambios"
			Exit Sub
		End If
		
	End Sub

	Function AgregarPRPro(ByVal Conexion, _
                         ByVal Moneda, ByVal PMCompra, ByVal PMVenta, _
                         ByVal Usuario, _
								 ByVal ParidadCompra, ByVal ParidadVenta, _
								 ByVal ParidadTransfer)
		Dim sSQL
		Dim BD
   
		On Error Resume Next
		AgregarPRPro = False
		sSQL = "APInsertarPRPro " & EvaluarStr(Moneda) & ", '" & FormatoFechaSQL(Date) & "', " & EvaluarStr(Time) & ", " & _
		                        FormatoNumeroSQL(PMCompra) & ", " & FormatoNumeroSQL(PMVenta) & ", " & _
		                        EvaluarStr(Usuario) & ", " & EvaluarVar(Orden, "Null") & ", " & _
		                        FormatoNumeroSQL(ParidadCompra) & ", " & FormatoNumeroSQL(ParidadVenta) & ", " & _
		                        FormatoNumeroSQL(ParidadTransfer)
   
		'Conexion
		Set BD = CreateObject("ADODB.Connection")
		BD.Open Conexion                          'Abre la conexion
		If Err.Number <> 0 Then
			Set BD = Nothing
			Exit Function
		End If
		   
		'Consulta
		BD.BeginTrans
		BD.Execute sSQL                           'Ejecuta la consulta
		If Err.Number <> 0 Then 
			BD.RollbackTrans 
			BD.Close 
			Set BD = Nothing
			Exit Function
		End If
		   
		AgregarPRPro = True
		BD.CommitTrans
		BD.Close
		Set BD = Nothing
	End Function


	Function ActualizarParidadesTrf(ByVal Conexion)
		Dim sSQL
		Dim BD
   
		On Error Resume Next
		ActualizarParidadesTrf = False
	   sSQL = "APActualizarParidadTransfer"
		'Conexion
		Set BD = CreateObject("ADODB.Connection")
		BD.Open Conexion                          'Abre la conexion
		If Err.Number <> 0 Then
			Set BD = Nothing
			Exit Function
		End If
		   
		'Consulta
		BD.BeginTrans
		BD.Execute sSQL                           'Ejecuta la consulta
		If Err.Number <> 0 Then 
			BD.RollbackTrans 
			BD.Close 
			Set BD = Nothing
			Exit Function
		End If
		   
		ActualizarParidadesTrf = True
		BD.CommitTrans
		BD.Close
		Set BD = Nothing
	End Function


	Sub MostrarAP()
		window.showModalDialog "APAperturaPunta.asp?mn=<%=Request("mn")%>&cmn=<%=Request("cmn")%>&tcc=" & txtTCCompra.value  & "&tcv=" & txtTCVenta.value & "&tct=" & txtTCTrf.value & "&ptr=" & txtParidadTrf.value & "&tcusv=" & nUSDVenta , , "dialogWidth:26; dialogHeight:24"
	End Sub

	Sub HabilitarParidades
		tbParidades.disabled = Not chkParidades.checked
	End Sub

	Sub HabilitarPrecios
		tbPrecios.disabled = Not chkPrecios.checked
	End Sub
	
	Sub ActualizarPTrf
		Dim bOk
	   If MsgBox("Está seguro que desea cargar las paridades para cheques y transferencias?", vbExclamation + vbYesNo) <> vbYes Then Exit Sub
	   
      On Error Resume Next
      bOk = ActualizarParidadesTrf("<%=Session("cnxVisorAP")%>")
		If Not bOk Then 
			MsgBox "Se produjo un error al intentar actualizar las paridades de transferencias", vbCritical
			Exit Sub
		End If
		window.close 
	End Sub
	
//-->
</script>
<body style="margin-left: 10px; margin-top: 10px">
<input type="hidden" id="txtMoneda">
<center>
<table cellpadding="0" cellspacing="1" style="border: 1px solid silver; width: 100%" ,>
<tr class="titulo"><td id="tdMoneda"></td></tr>
<tr style="background-color: '#EFEFEF'"><td>Paridades</td></tr>
<tr>
	<td>
		<table id="tbParidades" cellpadding="5">
		<tr align="center">
			<td>Compra<br><input id="txtParidadCompra" style="width: 80px" onKeyPress="IngresarNumero" onBlur="CalcularPrecios"></td>
			<td>Venta<br><input id="txtParidadVenta" style="width: 80px" onKeyPress="IngresarNumero" onBlur="CalcularPrecios"></td>
			<td align="center">Transfer<br><input id="txtParidadTrf" style="width: 80px" onKeyPress="IngresarNumero" onBlur="CalcularPrecios"></td>
		</tr>		
		<tr style="height: 16px;">
			<td colspan="3">
			<table>
			<tr align="center">
				<td width="100px"></td>
				<td class="boton" style="height: 20px; width: 180px;" onClick="ActualizarPTrf" onMouseOver="MouseOver()" onMouseOut="MouseOut()">Paridades Trf</td>
				<td class="boton" style="height: 20px; width: 180px;" onClick="MostrarAP" onMouseOver="MouseOver()" onMouseOut="MouseOut()">Productos</td>
			</tr>
			</table>
			</td>
		</tr>
		</table>
	<td>
</tr>
<tr height="10px"><td></td></tr>
<tr style="background-color: '#EFEFEF'"><td>Precios</td></tr>
<tr><td><input id="chkPrecios" type="checkbox" onClick="HabilitarPrecios">Cambiar Precios</td></tr>
<tr>
	<td>
		<table id="tbPrecios" cellpadding="5" align="center">
		<tr>
			<td>Compra<br><input id="txtTCCompra" onKeyPress="IngresarNumero" style="width: 80px"></td>
			<td>Venta<br><input id="txtTCVenta" onKeyPress="IngresarNumero" style="width: 80px"></td>
			<td>Transfer<br><input id="txtTCTrf" onKeyPress="IngresarNumero" style="width: 80px"></td>
		</tr>		
		</table>
	<td>
</tr>
<tr height="10px"><td></td></tr>
</table>
<table>
<tr height="20px"><td></td></tr>
<tr align="center" style="height: 20px;">
	<td class="boton" style="width: 100px;" onClick="AceptarPR" onMouseOver="MouseOver()" onMouseOut="MouseOut()">Aceptar</td>
	<td class="boton" style="width: 100px;" onClick="AplicarPR" onMouseOver="MouseOver()" onMouseOut="MouseOut()">Aplicar</td>
	<td class="boton" style="width: 100px;" onClick="window.close" onMouseOver="MouseOver()" onMouseOut="MouseOut()">Salir</td>
</tr>
</table>
</center>
</body>
</html>
