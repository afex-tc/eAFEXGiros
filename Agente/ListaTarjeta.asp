<%@ Language=VBScript %>
<%
	'Se asegura que la página no se almacene en la memoria cache
	Response.Expires = 0
	If Session("CodigoAgente") = "" Then
		response.Redirect "../Compartido/TimeOut.htm"
		response.end
	End If
%>
<%

	'Variables de módulo
	'Variables para encabezado
	Dim sEncabezadoFondo
	Dim sEncabezadoTitulo
	
	Dim sTitulo, nTipo, sSucursal
	Dim dDesde, dHasta, sCliente, sUsuario
	
	sTitulo = Request("Titulo")
	nTipo = 11
	nTipoLlamada = Request("TipoLlamada")
	'mostrarerrorms Request("TipoLlamada") & ", " & nTipoLlamada
	dDesde = request("Desde")
	dHasta = request("Hasta")
	sCliente = request("Cliente")
	sSucursal = Request("sc")
	
	If sSucursal = Empty Then
		sSucursal = Trim(Session("CodigoAgente"))
	Else
		sSucursal = Request("sc")
	End If
	
	'sUsuario = Request("Usuario")
	If dDesde = "" Then dDesde = Date() '- 31
	If dHasta = "" Then dHasta = Date()
	
	If Trim(sTitulo) = "" Then sTitulo = "Tarjetas"
	sEncabezadoFondo = "Consultas"
	sEncabezadoTitulo = sTitulo

	Function ObtenerTarjetas()
		'Dim afxTransfer, rs, sSQL
		Dim rs, sSQL
		
		On Error Resume Next
		
		sSQL = "select t.*, m.nombre_moneda, isnull(tc.valor, 0) as valor " & _
			   "from tarjeta t " & _
			   "inner join moneda m on m.codigo_moneda = t.codigo_moneda " & _
			   "left outer join vtipo_cambio tc on tc.codigo_moneda = t.codigo_moneda " & _
			   "and	tc.fecha_termino is null " & _
			   "and	tc.sw_tipo = 2 " & _
			   "where " & _
			   "fecha between '" & Trim(cdate(dDesde)) & "' " & _
			   "and '" & Trim(cdate(dHasta)) & "' " & _
			   "and codigo_agente = '" & sSucursal & "' and t.tipo_pin=2 " & _
			   "order by numero_boleta"

		Set rs = CreateObject("ADODB.Recordset") 
		
		rs.Open sSQL, Session("afxCnxAFEXpress"), 3, 1
		
		If Err.number <> 0 Then
			Set rs = Nothing
			MostrarErrorMS "ObtenerTarjetas", Err.Description
			Exit Function
		End If
		
		Set ObtenerTarjetas = rs
		Set rs = Nothing
	End Function
%> 
<!--#INCLUDE virtual="/Compartido/Errores.asp" -->
<html>
<head>
<META name=VI60_defaultClientScript content=JavaScript>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<title>AFEX</title>
<link rel="stylesheet" type="text/css" href="../Estilos/Cliente.css">
<link rel="stylesheet" type="text/css" href="../Estilos/Reportes.css">
</head>
<script LANGUAGE="VBScript">
<!--
	Sub imgAceptar_onClick()
		If window.tbReporte.style.display = "" then
			window.tbReporte.style.display = "none"
		Else 
			window.tbReporte.style.display = ""
		End If
	End Sub		
	
	Function imgAceptar_onMouseOver()
		window.imgAceptar.style.cursor = "Hand"		
	End Function

	Sub window_onload()
		objConsulta.Desde = cdate("<%=dDesde%>")
		objConsulta.Hasta = cdate("<%=dHasta%>")
		objConsulta.Tipo = <%=nTipo%>
		objConsulta.CodigoCliente = "<%=sCliente%>"
		objConsulta.Sucursal = "<%=sSucursal%>"
	End Sub
	
	
//-->
</script>
<body>
<!--#INCLUDE virtual="/Compartido/Encabezado.asp" -->
<!--#INCLUDE virtual="/Compartido/Rutinas.htm" -->
<table border="0" cellspacing="0" cellpadding="0" style="LEFT: 0px; POSITION: relative; TOP: 0px">
<tr><td align="middle">
      <object id="objConsulta" style="HEIGHT: 200px; LEFT: 0px; TOP: 0px; WIDTH: 544px" type="text/x-scriptlet" width="544" VIEWASTEXT><param NAME="Scrollbar" VALUE="0"><param NAME="URL" VALUE="http:ConfiguracionConsulta.asp"></object>
</td></tr>
<tr height="10"><td>
	<table cellspacing="1" cellpadding="1" ID="tbReporte" border="0" ALIGN="center" STYLE="COLOR: #505050; FONT-FAMILY: Verdana; FONT-SIZE: 10px; POSITION: relative; TOP: 0px">
<%
		Dim rsTar, nTotal, nTransfer, nCantidad, sMoneda
		Dim sCodigoMoneda
		
		nCantidad = 0
		sMoneda = Empty
		
		set rsTar = ObtenerTarjetas()

		Do Until rsTar.EOF 
			If sMoneda <> rsTar("nombre_moneda") Then
				If sMoneda <> Empty Then
%>
					<tr style="height: 20px" CLASS="Encabezado">
						<td style="background-color: white"></td>
						<td align="left"><b><%=FormatNumber(nCantidad, 0)%>&nbsp;Tarjetas</b></td>			
						<td align="right"><b>Total</b></td>
						<% If sCodigoMoneda = "CLP" Then %>
							<td align="right"><b><%=FormatNumber(nTotal, 0)%></b></td>
						<% Else %>
							<td align="right"><b><%=FormatNumber(nTotal, 2)%></b></td>
						<% End If %>
					</tr>
					<tr><td><br><br><br></td></tr>
<%				
					nTotal = 0
					nCantidad = 0
				End If
				sCodigoMoneda = rsTar("codigo_moneda")
				sMoneda = rsTar("nombre_moneda")
%>		
				<tr>
					<td colspan=2 style="font-size: 12px; font-weight: bold"><%=sMoneda%></td>
				</tr>
				<tr CLASS="Encabezado">
					<td WIDTH="100">
						<b>Fecha</b>
					</td>
					<td WIDTH="100">
						<b>Nº Boleta</b>
					</td>
					<td WIDTH="100">
						<b>Nº PIN</b>
					</td>
					<td WIDTH="100">
						<b>Monto</b>
					</td>
					<td WIDTH="150">
						<b>Usuario</b>
					</td>
				</tr>
		<% End If %>
			<a href="VenderTarjeta.asp?hb=disabled&mn=<%=rsTar("codigo_moneda")%>&mto=<%=rsTar("monto")%>&bs=<%=rsTar("numero_boleta")%>&np=<%=rsTar("numero_pin")%>&tc=<%=rsTar("valor")%>" onmouseout="window.status=''" onmouseover="window.status='Ver detalle'" onclick="">
				<tr style="HEIGHT: 25px" language="javascript" onmouseover="javascript:this.bgColor='#a4dded'" onmouseout="javascript:this.bgColor='#DAF6FF'" bgColor="#dbf7ff" style="cursor: hand">
					<td align="center"><%=rsTar("fecha")%></td>
					<td align="right"><%=rsTar("numero_boleta")%></td>
					<td align="right"><%=rsTar("numero_pin")%></td>
					<% If rsTar("codigo_moneda") = "USD" Then %>
						<td ALIGN="right"><%=FormatNumber(rsTar("monto"), 2)%></td>
					<% Else %>
						<td ALIGN="right"><%=FormatNumber(rsTar("monto"), 0)%></td>
					<% End If %>
					<% If Isnull(rsTar("codigo_usuario")) Then %>
						<td></td>
					<% Else %>
						<td><%=rsTar("codigo_usuario")%></td>
					<% End If %>
				</tr>
			</a>
<%
			nTotal = nTotal + cCur(0 & rsTar("monto"))
			nCantidad = nCantidad + 1
			rsTar.MoveNext
		Loop
%>
		<tr style="height: 20px" CLASS="Encabezado">
			<td style="background-color: white"></td>
			<td align="left"><b><%=FormatNumber(nCantidad, 0)%>&nbsp;Tarjetas</b></td>			
			<td align="right"><b>Total</b></td>
			<% If sCodigoMoneda = "CLP" Then %>
				<td align="right"><b><%=FormatNumber(nTotal, 0)%></b></td>
			<% Else %>
				<td align="right"><b><%=FormatNumber(nTotal, 2)%></b></td>
			<% End If %>
		</tr>
	<!--	<tr style="height: 20px" CLASS="Encabezado">
			<td colspan="3" style="background-color: white"></td>
			<td colspan="2"  ALIGN="left"><b><%=FormatNumber(nCantidad, 0)%>&nbsp;Tarjetas</b></td>			
			<td align="right"><b>Total</b></td>
			<td ALIGN="right"><b><%=FormatNumber(nTotal, 2)%></b></td>
		</tr>
	-->
<%
		Set rsTar = Nothing
%>
</table>
</td></tr>
</table>
</body>
<script>

	Sub objConsulta_OnScriptletEvent(strEventName, varEventData)
		
	   Select Case strEventName
	   
			Case "Aceptar"				
				window.navigate "ListaTarjeta.asp?Titulo=<%=sTitulo%>&Desde=" & objConsulta.Desde & _
				"&Hasta=" & objConsulta.Hasta & "&sc=" & objConsulta.Sucursal
		
		End Select
		
	End Sub
		
</script>

</html>
