<%@ Language=VBScript %>
<%
	Response.Expires = 0
	If Session("CodigoCliente") = "" Then
		response.Redirect "../Compartido/TimeOut.htm"
		response.end
	End If
%>
<%
	Dim sEncabezadoFondo
	Dim sEncabezadoTitulo

	Dim sTitulo, nTipo
	Dim dDesde, dHasta, sCliente, sAgente, sPorcentaje, nAFEX
	
	sEncabezadoFondo = "Consultas"
	sEncabezadoTitulo = "Movimientos"

	sTitulo = Request("Titulo")
	dDesde = request("Desde")
	dHasta = request("Hasta")
	sAgente = request("Agente")
	sPorcentaje = request("Porcentaje")
	If Request("optAFEX") = Empty Then
		nAFEX = True
	ElseIF Request("optAFEX") = "Verdadero" Then
		nAFEX = True
	Else
	    nAFEX = False
	    
	End If
%>

<HTML>
<HEAD>
<META name="VI60_DefaultClientScript" Content="VBScript">
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<title><%=sTitulo%></title>
<link rel="stylesheet" type="text/css" href="../Estilos/Cliente.css">
<link rel="stylesheet" type="text/css" href="../Estilos/Reportes.css">
</HEAD>
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
		If "<%=dDesde%>" = Empty And "<%=dHasta%>" = Empty Then
			objConsulta.Desde = Date()
			objConsulta.Hasta = Date()
			objConsulta.optAFEX = true
		Else
			objConsulta.Desde = cdate("<%=dDesde%>")
			objConsulta.Hasta = cdate("<%=dHasta%>")
			objConsulta.Agente = "<%=sAgente%>"
			objConsulta.Porcentaje = "<%=sPorcentaje%>"
			objConsulta.optAFEX = <%=nAFEX%>
		End If
	End Sub
//-->
</script>
<BODY><!--#INCLUDE virtual="/Compartido/Encabezado.asp" -->
<table border="0" cellspacing="0" cellpadding="0" style="LEFT: 0px; POSITION: relative; TOP: 0px">
<tr><td align="middle">
      <OBJECT id=objConsulta 
      style="HEIGHT: 200px; LEFT: 0px; TOP: 0px; WIDTH: 450px" 
      type=text/x-scriptlet width=544 VIEWASTEXT><PARAM NAME="Scrollbar" VALUE="0"><PARAM NAME="URL" VALUE="http:ConfiguracionConsulta.asp"></OBJECT>
	</td>
	<td width="30%">
			
	</td>
</tr>
<tr height="10"><td colspan="2">
	<table width="760" cellspacing="1" cellpadding="1" ID="tbReporte" border="0" ALIGN="center" STYLE="COLOR: #505050; FONT-FAMILY: Verdana; FONT-SIZE: 10px; POSITION: relative; TOP: 0px">
<%
	Dim rsCompliance, Sql, rsObs, cObservado
	Dim cRecargoCompra, cRecargoVenta, sOperacion
	Dim dFecha

	
	If dDesde <> Empty And dHasta <> Empty Then
		dFecha = DateAdd("d", 0, dDesde)
		dHasta = DateAdd("d", 0, dHasta)
		Do Until dFecha > dHasta
			Lista
			dFecha = DateAdd("d", 1, dFecha)
		Loop
	End If
	

	Sub Lista	
		Sql = "Select " & _
					  "tipo_cambio_observado " & _
			  "From   " & _
					  "Jornada " & _
			  "Where  " & _
					  "fecha_jornada = convert(char, '" & dFecha & "', 112)"

		Set rsObs = CreateObject("ADODB.Recordset")
		rsObs.Open Sql, Session("afxCnxAFEXchange")
		
		If rsObs.EOF Then
			Sql = "Select " & _
						  "tipo_cambio_observado " & _
				  "From   " & _
						  "Jornada " & _
				  "Where  " & _
						  "fecha_jornada = (SELECT MAX(Fecha_jornada) FROM Jornada WHERE fecha_jornada < convert(char, '" & dFecha & "', 112))"
						  
			Set rsObs = CreateObject("ADODB.Recordset")
			rsObs.Open Sql, Session("afxCnxAFEXchange")
		End If
			
		If Not (rsObs.EOF) Then
			cObservado = rsObs("tipo_cambio_observado")
			If nAFEX Then
				cRecargoCompra = CCur(cObservado) * (1 + (sPorcentaje / 100))
				cRecargoVenta = CCur(cObservado) * (1 - (sPorcentaje / 100))
			Else
				cRecargoCompra = CCur(cObservado) * (1 - (sPorcentaje / 100))
				cRecargoVenta = CCur(cObservado) * (1 + (sPorcentaje / 100))
			End If
		Else	
			cObservado = 0
			cRecargoCompra = 0
			cRecargoVenta = 0
		End If

		Set rsObs = Nothing
			
		Sql = "Select " & _
					  "c.*, m.alias_moneda, " & _
					  "t.nombre_tipo " & _
			  "From " & _
					  "Compliance c " & _
			  "inner join moneda m on m.codigo_moneda = c.codigo_moneda " & _
			  "inner join tipo t on t.codigo_tipo = c.codigo_producto " & _
			  "and	   t.nombre_campo = 'PRODUCTO' " & _
			  "Where " & _
					  "c.fecha_solicitud between convert(char, '" & dFecha & "', 112) and convert(char, '" & dFecha & "', 112) " & _
			  "and     c.codigo_agente = '" & sAgente & "' " 
				  
		If CDbl("0" & sPorcentaje) > 0 Then
			If nAFEX Then
				Sql = Sql & "and	   c.tipo_operacion = 1 " & _
							"and	   c.tipo_cambio >= " & Replace(cRecargoCompra, ",", ".") & " "
			Else
				Sql = Sql & "and	   c.tipo_operacion = 1 " & _
							"and	   c.tipo_cambio <= " & Replace(cRecargoCompra, ",", ".") & " "
			End If

			Sql = Sql & "Union " & _
						"Select " & _
						  "c.*, m.alias_moneda, " & _
						  "t.nombre_tipo " & _
						"From " & _
								  "Compliance c " & _
						"inner join moneda m on m.codigo_moneda = c.codigo_moneda " & _
						"inner join tipo t on t.codigo_tipo = c.codigo_producto " & _
						"and	   t.nombre_campo = 'PRODUCTO' " & _
						"Where " & _
								  "c.fecha_solicitud between convert(char, '" & dFecha & "', 112) and convert(char, '" & dFecha & "', 112) " & _
						"and       c.codigo_agente = '" & sAgente & "' "  & _
						"and	   c.tipo_operacion = 2 "
			
			If nAFEX Then
				Sql = Sql & " and	   c.tipo_cambio <= " & Replace(cRecargoVenta, ",", ".") & " "
			Else
				Sql = Sql & " and	   c.tipo_cambio >= " & Replace(cRecargoVenta, ",", ".") & " "
			End If

		End If

	    Sql = Sql & "Order by " & _
						"c.tipo_operacion, " & _
						"c.codigo_producto, " & _
						"m.alias_moneda, " & _
						"c.monto_extranjera, " & _
						"c.codigo_solicitud"
			
		set rsCompliance = CreateObject("ADODB.Recordset")
		rsCompliance.open Sql, Session("afxCnxCorporativa")
		
		Encabezado
		Do Until rsCompliance.EOF
			Select Case rsCompliance("tipo_operacion")
				Case 1
					sOperacion = "Compra"
						
				Case 2
					sOperacion = "Venta"
						
			End Select
%>			
			<tr style="HEIGHT: 25px" language="javascript" onmouseover="javascript:this.bgColor='#a4dded'" onmouseout="javascript:this.bgColor='#DAF6FF'" bgColor="#dbf7ff" style="cursor: hand">
				<!--<td>=rsCompliance("fecha_solicitud")</td>-->				
				<td><%=sOperacion%></td>
				<td><%=rsCompliance("nombre_tipo")%></td>
				<td><%=rsCompliance("alias_moneda")%></td>
				<td align="right"><%=FormatNumber(rsCompliance("monto_extranjera"), 2)%></td>
				<td align="right"><%=FormatNumber(rsCompliance("tipo_cambio"), 4)%></td>
				<td align="right"><%=FormatNumber(rsCompliance("monto_nacional"), 0)%></td>
				<td><%=rsCompliance("nombre_cliente")%></td>
				<td align="right"><%=rsCompliance("codigo_solicitud")%></td>
			</tr>
<%
			rsCompliance.MoveNext
		Loop
%>
		<tr><td><br><br><br></td></tr>
		<tr style="height: 20px" CLASS="Encabezado">
		<td colspan="3" style="background-color: white"></td>
		<!--<td colspan="2"  ALIGN="left"><b><%=FormatNumber(nCantidad, 0)%>&nbsp;Transferencias</b></td>			
		<td align="right"><b>Total</b></td>
		<td ALIGN="right"><b><%=FormatNumber(nTotal, 2)%></b></td>-->
		</tr>
<%
		Set rsCompliance = Nothing
	End Sub
	

	Sub Encabezado
%>
		<tr style="height: 20px">
		<td CLASS="Titulo" colspan="2" align="left" style="font-size: 8pt">Fecha&nbsp;<%=dFecha%></td>
		<td CLASS="Titulo" colspan="2"  align="left" style="font-size: 8pt">Observado&nbsp;<%=FormatNumber(cObservado, 2)%></td></b>
		<td CLASS="Titulo" colspan="2" align="left" style="font-size: 8pt">Compra&nbsp;<%=cRecargoCompra%></td>		
		<td CLASS="Titulo" colspan="2" align="left" style="font-size: 8pt">Venta&nbsp;<%=cRecargoVenta%></td>		
		</tr>
		<tr CLASS="Encabezado">
			<!--<td WIDTH="150">
				<b>Fecha</b>
			</td>-->
			<td WIDTH="60">
				<b>Operación</b>
			</td>
			<td WIDTH="60">
				<b>Producto</b>
			</td>
			<td WIDTH="100">
				<b>Moneda</b>
			</td>
			<td WIDTH="100">
				<b>Monto</b>
			</td>
			<td WIDTH="80">
				<b>T/C</b>
			</td>
			<td WIDTH="100">
				<b>Monto $</b>
			</td>
			<td WIDTH="400">
				<b>Cliente</b>
			</td>
			<td WIDTH="80">
				<b>Código Solicitud</b>
			</td>
		</tr>
<%				
	End Sub
	
%>
</table>
</td></tr>
</table>
</BODY>
<script>

	Sub objConsulta_OnScriptletEvent(strEventName, varEventData)
	   Select Case strEventName
	   
			Case "Aceptar"				
				window.navigate "ConsultaCompraVenta.asp?Titulo=<%=sTitulo%>&Desde=" & objConsulta.Desde & _
				"&Hasta=" & objConsulta.Hasta & "&Agente=" & objConsulta.Agente & _
				"&Porcentaje=" & objConsulta.Porcentaje & "&optAFEX=" & objConsulta.optAFEX
		
		End Select
	End Sub
		
</script>
</HTML>
