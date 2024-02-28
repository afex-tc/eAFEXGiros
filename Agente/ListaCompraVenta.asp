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
<!--#INCLUDE virtual="/Compartido/Errores.asp" -->
<%
	'Variables de módulo
	'Variables para encabezado
	Dim sEncabezadoFondo
	Dim sEncabezadoTitulo
	
	Dim sTitulo
	Dim nTipo	
	Dim afxWeb
	Dim rsSP, sDesde, sHasta, nTipoLlamada
	
	'Rescata parámetros
	sTitulo = Trim(Request("Titulo"))
	nTipo = cInt(0 & Request("Tipo"))
	sDesde = Request("Desde")
	sHasta = Request("Hasta")
	nTipoLlamada = cInt(0 & request("TipoLlamada"))
	If sDesde = "" Then sDesde = Date()
	If sHasta = "" Then sHasta = Date()
			
	'**Dependiendo del tipo de pantalla, rescata la lista	
	On Error Resume Next
	Set afxWeb = Server.CreateObject("AFEXWeb.Web")	
	Set rsSP = afxWeb.ObtenerMovimientos(Session("afxCnxAFEXweb"), Session("CodigoCaja"), cDate(sDesde),  cDate(sHasta))
	
	If Err.number <> 0 Then
		Set rsSP = Nothing
		Set afxWeb = Nothing
		MostrarErrorMS ""
	End If
	If afxWeb.ErrNumber <> 0 Then			
		Set rsSP = Nothing
		MostrarErrorAFEX afxWeb, ""
	End If
	sEncabezadoFondo = "Consultas"
	sEncabezadoTitulo = sTitulo

	'**
%>
<html>
<head>
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
		objConsulta.Desde = cDate("<%=sDesde%>")
		objConsulta.Hasta = cDate("<%=sHasta%>")
		objConsulta.Tipo = <%=Request("Tipo")%>
	End Sub
	
	
//-->
</script>
<body>
<!--#INCLUDE virtual="/Compartido/Encabezado.asp" -->

<table border="0" cellspacing="0" cellpadding="0" style="LEFT: 0px; POSITION: relative; TOP: 0px">
<tr>
	<td align="middle">
		<!--Si la página es para mostrar los SPs pendientes del cliente, no se muestra el 
		filtro-->
		<OBJECT height=224 id=objConsulta style="HEIGHT: 224px; LEFT: 0px; TOP: 0px; WIDTH: 544px" 
		type=text/x-scriptlet width=544 VIEWASTEXT>
		<PARAM NAME="Scrollbar" VALUE="0">
		<PARAM NAME="URL" VALUE="http://afexweb/AfexMoneyWeb/Agente/ConfiguracionConsulta.asp"></OBJECT>
	</td>
</tr>
<tr height="10"><td>
	<table cellspacing="1" cellpadding="1" ID="tbReporte" border="0" ALIGN="center" STYLE="COLOR: #505050; FONT-FAMILY: Verdana; FONT-SIZE: 10px; POSITION: relative; TOP: 0px">
	<tr CLASS="Encabezado">
		<td WIDTH="100">
			<b>Fecha</b>
		</td>
		<td WIDTH="100">
			<b>Operación</b>
		</td>
		<td WIDTH="100">
			<b>Producto</b>
		</td>
		<td WIDTH="100">
			<b>Moneda</b>
		</td>
		<td WIDTH="200">
			<b>Monto</b>
		</td>
		<td WIDTH="150">
			<b>Tipo Cambio</b>
		</td>
		<td WIDTH="200">
			<b>Total</b>
		</td>
	</tr>
		<%				
		Dim nTotal, sDetalle, nMonto, nMontoExtranjera
		nTotal = 0
		nMonto = 0 
		nMontoExtranjera = 0

		Do Until rsSP.EOF 
		%>		
			<tr style="HEIGHT: 25px" language="javascript" onmouseover="javascript:this.bgColor='#a4dded'" onmouseout="javascript:this.bgColor='#DAF6FF'" bgColor="#dbf7ff">
				<td><%=rsSP("fecha_solicitud")%></td>
				<td><%=rsSP("nombre_operacion")%></td>
				<td><%=rsSP("nombre_producto")%></td>
				<td><%=rsSP("codigo_moneda")%></td>
		<%
			If rsSP("tipo_operacion") = afxOperacionVenta Then 
				nMonto = cCur(0 & rsSP("monto_nacional")) 
				nMontoExtranjera = cCur(rsSP("monto_extranjera")) * -1
			Else
				nMonto = cCur(0 & rsSP("monto_nacional")) * -1
				nMontoExtranjera = cCur(0 & rsSP("monto_extranjera"))
			End If
		%>
				<td ALIGN="right"><%=formatNumber(nMontoExtranjera, 2, , -1)%></td>
				<td ALIGN="right"><%=FormatNumber(rsSP("tipo_cambio"), 4)%></td>
				<td ALIGN="right"><%=FormatNumber(nMonto, 0, , -1)%></td>
			</tr>
		<%			
			nTotal = nTotal + nMonto
			rsSP.MoveNext
		Loop
		%>
		<tr style="height: 20px" CLASS="Encabezado">
			<td colspan="5" style="background-color: white"></td>
			<td align="left"><b>Total</b></td>
			<td ALIGN="right"><b><%=FormatNumber(nTotal, 0, , -1)%></b></td>
		</tr>
</table>
</td></tr>
</table>
</body>
<script>
	Sub objConsulta_OnScriptletEvent(strEventName, varEventData)
		
	   Select Case strEventName
	   
			Case "Aceptar"
				window.navigate "ListaCompraVenta.asp?Titulo=<%=sTitulo%>&Desde=" & objConsulta.Desde & _
				"&Hasta=" & objConsulta.Hasta & "&Tipo=<%=nTipo%>"
				
		End Select
		
	End Sub
</script>
</html>
