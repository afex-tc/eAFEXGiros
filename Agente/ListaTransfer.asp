<%@ Language=VBScript %>
<%
	'Se asegura que la página no se almacene en la memoria cache
	Response.Expires = 0
	If Session("CodigoCliente") = "" Then
		response.Redirect "../Compartido/TimeOut.htm"
		response.end
	End If
%>
<%

	'Variables de módulo
	'Variables para encabezado
	Dim sEncabezadoFondo
	Dim sEncabezadoTitulo
	
	Dim sTitulo, nTipo
	Dim dDesde, dHasta, sCliente, sUsuario
	
	sTitulo = Request("Titulo")
	nTipo = 9
	nTipoLlamada = Request("TipoLlamada")
	'mostrarerrorms Request("TipoLlamada") & ", " & nTipoLlamada
	dDesde = request("Desde")
	dHasta = request("Hasta")
	sCliente = request("Cliente")
	sUsuario = Request("Usuario")
	If dDesde = "" Then dDesde = Date() - 31
	If dHasta = "" Then dHasta = Date()
	
	If Trim(sTitulo) = "" Then sTitulo = "Transferencias"
	sEncabezadoFondo = "Consultas"
	sEncabezadoTitulo = sTitulo

	Function ObtenerTransfer()
		Dim afxTransfer, rs, sSQL
		
		On Error Resume Next
		
		sSQL = "fecha_transferencia between '" & Trim(cdate(dDesde)) & "'" & _
									 "  and '" & Trim(cdate(dHasta)) & "' "
		If Trim(sCliente) <> "" Then
			sSQL = sSQL & " and codigo_cliente = '" & trim(sCliente) & "'"
		End If

		If Trim(sUsuario) <> "" Then
			sSQL = sSQL & " and usuario = '" & trim(sUsuario) & "'"
		End If
		'mostrarerrorms ssql
						
		'Set afxTransfer = Server.CreateObject("AfexProducto.Transferencia")
		'Set rs = afxTransfer.Buscar(Session("afxCnxAFEXchange"), 6, sSQL, 0, True, True)
		Set rs = BuscarTRF(Session("afxCnxAFEXchange"), 6, sSQL, 0, True, True)
		
		If Err.number <> 0 Then
			Set rs = Nothing
			'Set afxTransfer = Nothing
			MostrarErrorMS "Lista Transferencias 1"
		End If
		'If afxTransfer.ErrNumber <> 0 Then			
		'	Set rs = Nothing
		'	MostrarErrorAFEX afxTransfer, "Lista Transferencias 2" 
		'End If
		
		Set ObtenerTransfer = rs
		Set rs = Nothing
		'Set afxTransfer = Nothing
	End Function
	

%> 
<!--#INCLUDE virtual="/Compartido/Rutinas.asp" -->
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
	<tr CLASS="Encabezado">
		<td WIDTH="200">
			<b>Fecha</b>
		</td>
		<td WIDTH="300">
			<b>Banco Destino</b>
		</td>
		<td WIDTH="300">
			<b>Cuenta Destino</b>
		</td>
		<td WIDTH="100">
			<b>Estado</b>
		</td>
		<td WIDTH="250">
			<b>Moneda</b>
		</td>
		<td WIDTH="200">
			<b>Monto</b>
		</td>
		<td WIDTH="200">
			<b>Monto USD</b>
		</td>
	</tr>
<%
		Dim rsTrf, nTotal, nTransfer, nCantidad
		nCantidad = 0
		set rsTrf = ObtenerTransfer()

		Do Until rsTrf.EOF 
			'If sDetalle <> "Detalle1" Then
			'	sDetalle = "Detalle1"
			'Else
			'	sDetalle = "Detalle2"
			'End If
			
%>		
			<a href="../Compartido/DetalleTransfer.asp?Correlativo=<%=rsTrf("correlativo_transferencia")%>&Cliente=<%=rsTrf("codigo_cliente")%>&Tipo=<%=nTipoLlamada%>" onmouseout="window.status=''" onmouseover="window.status='Ver detalle'" onclick="">
			<!--<a href="EnviarTransfer.asp?Correlativo=<%=rsTrf("correlativo_transferencia")%>" onmouseout="window.status=''" onmouseover="window.status='Ver detalle'" onclick="">-->
<!--			<tr CLASS="<%=sDetalle%>" sonmouseover="Seleccionar()" sonmouseout="QuitarSeleccion()" >-->
				<!-- <td onmouseover="window.status ='=rsTrf("nombre_completo")'">=rsTrf("nombre_completo")</td> -->
			<tr style="HEIGHT: 25px" language="javascript" onmouseover="javascript:this.bgColor='#a4dded'" onmouseout="javascript:this.bgColor='#DAF6FF'" bgColor="#dbf7ff" style="cursor: hand">
				<td><%=rsTrf("fecha_transferencia")%></td>
				<td><%=rsTrf("nombre_banco_destino")%></td>
				<td><%=rsTrf("cuenta_corriente_destino")%></td>
				<td><%=rsTrf("nombre_estado")%></td>
				<td><%=rsTrf("nombre_moneda")%></td>
				<td ALIGN="right"><%=FormatNumber(rsTrf("monto_transferencia"), 2)%></td>
				<% If cCur(0 & rsTrf("monto_equivalente")) = 0 Then %>
					<td ALIGN="right"><%=FormatNumber(rsTrf("monto_transferencia"), 2)%></td>
				<% Else %>
					<td ALIGN="right"><%=FormatNumber(rsTrf("monto_equivalente"), 2)%></td>
				<% End If %>
				<!--<td><%=sNegocio%></td>-->
			</tr>
			</a>
<%
			If cCur(0 & rsTrf("monto_equivalente")) = 0 Then
				nTotal = nTotal + cCur(0 & rsTrf("monto_transferencia"))
			Else
				nTotal = nTotal + cCur(0 & rsTrf("monto_equivalente"))
			End If
			nCantidad = nCantidad + 1
			rsTrf.MoveNext
		Loop
%>
		<tr style="height: 20px" CLASS="Encabezado">
			<td colspan="3" style="background-color: white"></td>
			<td colspan="2"  ALIGN="left"><b><%=FormatNumber(nCantidad, 0)%>&nbsp;Transferencias</b></td>			
			<td align="right"><b>Total</b></td>
			<td ALIGN="right"><b><%=FormatNumber(nTotal, 2)%></b></td>
		</tr>
<%
		Set rsTrf = Nothing
%>
</table>
</td></tr>
</table>
</body>
<script>

	Sub objConsulta_OnScriptletEvent(strEventName, varEventData)
		
	   Select Case strEventName
	   
			Case "Aceptar"				
				window.navigate "ListaTransfer.asp?Titulo=<%=sTitulo%>&Desde=" & objConsulta.Desde & _
				"&Hasta=" & objConsulta.Hasta & "&Cliente=" & objConsulta.CodigoCliente & "&Usuario=<%=sUsuario%>"
		
		End Select
		
	End Sub
		
</script>

</html>
