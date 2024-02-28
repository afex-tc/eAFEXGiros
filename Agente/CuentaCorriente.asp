<%@ Language=VBScript %>
<%
	Response.Expires = 0
	Response.buffer = true	
	If Session("CodigoCliente") = "" Then
		response.Redirect "../Compartido/TimeOut.htm"
		response.end
	End If
%>
<!--#INCLUDE virtual="/Compartido/Constantes.asp" -->
<!--#INCLUDE virtual="/Compartido/Errores.asp" -->
<!--#INCLUDE virtual="/Compartido/Rutinas.asp" -->
<!--#INCLUDE virtual="/Compartido/Giros.asp" -->
<%
	'Variables de módulo
	Dim rs, nSaldo
	Dim sDesde, sHasta
	   
	If Request("Desde")="" Then sDesde = Date Else sDesde = cDate(Request("Desde"))
	If Request("Hasta")="" Then sHasta = Date Else sHasta = cDate(Request("Hasta"))
	nSaldo = 0
	
	Set rs = Lista(Session("afxcnxAFEXpress"))
	If rs Is Nothing Then
		Response.Redirect "http:../Compartido/Informacion.asp?Detalle=No se registraron datos para este periodo"
		
	ElseIf rs.EOF Then
		Response.Redirect "http:../Compartido/Informacion.asp?Detalle=No se registraron datos para este periodo"
		
	Else
		
	End If
	
	
	Function Lista(ByVal Conexion)
	   Dim sSQL
	     
	   'Manejo de errores
	   On Error Resume Next
		Set Lista = Nothing   
	   'Crea la consulta
	   sSQL = "Execute CtaCte_AgenteWeb " & FormatoFecha(sDesde) & ", " & FormatoFecha(sHasta) & ", '" & Session("MonedaExtranjera") & "', '" & Session("CodigoAgente") & "'"
	   Set Lista = EjecutarSQLCliente(Conexion, sSQL)
	   'Si se produjeron errores en la consulta
	   If Err.Number <> 0 Then
			MostarErrorMS ""
	   End If
	   
	End Function

	Function FormatoFecha(Byval Fecha)
			fecha=cDate(fecha)
			FormatoFecha = "'" & Year(Fecha) & Right("0" & Month(Fecha), 2) & Right("0" & Day(Fecha), 2) & "'"
	End Function	


	Public Function EvaluarStr(ByVal Valor)
		Dim Devuelve
			
	  EvaluarStr = "'" & Valor & "'"

	End Function
	
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
	'Variables para encabezado	
	Dim sEncabezadoFondo
	Dim sEncabezadoTitulo
	
	sEncabezadoFondo = "&nbsp;"
	sEncabezadoTitulo = "Cuenta Corriente"

	Sub window_onload()
		CtaCte.objConsulta.Desde = cdate("<%=sDesde%>")
		CtaCte.objConsulta.Hasta = cdate("<%=sHasta%>")
	End Sub

	Sub imgAceptar_onClick()
		If Not CajaPregunta("AFEX En Linea", "Está seguro que desea notificar a AFEX Ltda. que la Cuadratura de Cuenta Corriente del periodo <%=sDesde%> al <%=sHasta%> está correcta?") Then
			Exit Sub
		End If
		trTitulo.style.display = ""
		trPeriodo.style.display = ""
		trPie.style.display = ""
		trObj.style.display = "none"
		CtaCte.imgAceptar.style.display = "none"
		CtaCte.imgImprimir.style.display = "none"
		CtaCte.Documento.value = window.tbBody.innerHTML
		CtaCte.action = "EnviarEmailCtaCte.asp"
		CtaCte.submit 
		CtaCte.action = ""
	End Sub
	
	Sub imgImprimir_OnClick()
		tbReporte.border = 1
		trTitulo.style.display = ""
		trPeriodo.style.display = ""
		trPie.style.display = ""
		trObj.style.display = "none"
		CtaCte.imgAceptar.style.display = "none"
		CtaCte.imgImprimir.style.display = "none"
		window.print()		
		tbReporte.border = 0
		trTitulo.style.display = "none"
		trPeriodo.style.display = "none"
		trPie.style.display = "none"
		trObj.style.display = ""
		CtaCte.imgAceptar.style.display = ""
		CtaCte.imgImprimir.style.display = ""		
	End Sub
	
//-->
</script>
<body onmousemove="window.status=''" media="print">
<!--#INCLUDE virtual="/Compartido/Encabezado.htm" -->
<form id="CtaCte" method="post">
<input type="hidden" name="Desde" value="<%=Replace(sDesde, "/", "-")%>">
<input type="hidden" value="<%=Replace(sHasta, "/", "-")%>" name="Hasta">
<input type="hidden" name="Documento" value="">
<table border="0" cellspacing="0" cellpadding="0" style="LEFT: 4px; WIDTH: 90%; POSITION: relative; TOP: 0px" media="print" >
<tr id="trObj"><td>
		<OBJECT id=objConsulta style="HEIGHT: 60px; LEFT: 0px; TOP: 0px; WIDTH: 544px" type=text/x-scriptlet width=544 VIEWASTEXT>
		<PARAM NAME="Scrollbar" VALUE="0">
		<PARAM NAME="URL" VALUE="http:PeriodoConsulta.asp"></OBJECT>
</td></tr>
</table>
<div id="tbBody">
<table  border="0" cellspacing="0" cellpadding="0" style="font-family: Verdana; font-size: 8pt; LEFT: 4px; WIDTH: 90%; POSITION: relative; TOP: 0px" media="print" >
<tr id="trTitulo" style="display: none; font-size: 10pt" ><td>Sres.<br><b><%=Session("NombreCliente")%></b></td></tr>
<tr height="20"><td></td></tr>
<tr id="trPeriodo" style="display: none; font-size: 8pt" ><td>Periodo del <%=Replace(sDesde, "/", "-")%> al <%=Replace(sHasta, "/", "-")%></td></tr>
<tr height="10"><td>
	<table cellspacing="1" cellpadding="5" ID="tbReporte" border="0" ALIGN="center" STYLE="FONT-SIZE: 10px; COLOR: #707070; FONT-FAMILY: Verdana; POSITION: relative; TOP: 0px; HEIGHT: 176px;" bgcolor="#e1e1e1">	
	<tr>
		<td colspan="6" style="FONT-SIZE: 8pt">
			<b><FONT color=black size=2>Cuadratura Cuenta Corriente</FONT> </b>
		</td>
	</tr>
	<tr bgcolor="#f1f1f1" height="20" align="middle">
		<td WIDTH="70%" colspan="4">
			<b>Detalle</b>
		</td>
		<td WIDTH="15%">
			<b>Saldo a favor<br>AFEX</b>
		</td>
		<td WIDTH="15%">
			<b>Saldo a favor<br>Agente</b>
		</td>
	</tr>
	<!-- Saldo Anterior -->
	<tr bgcolor="#fbfbfb" height="20" align="right" style="FONT-SIZE: 8pt; COLOR: #303030">
		<td colspan="4">
			<b><%=rs("Descripcion")%></b>
		</td>
		<td>&nbsp;
			<%If ccur(0 & rs("monto_afex")) <> 0 Then %>
				<% nSaldo = nSaldo - ccur(0 & rs("monto_afex")) %>
				<b><FONT color="red"><%=FormatNumber(cCur(0 & rs("monto_afex")), 2)%></FONT></b>
			<%End If %>
		</td>
		<td>&nbsp;
			<%If ccur(0 & rs("monto_agente")) <> 0 Then %>
			<% nSaldo = nSaldo + ccur(0 & rs("monto_agente")) %>
				<b><FONT color=mediumblue><%=FormatNumber(cCur(0 & rs("monto_agente")), 2)%></FONT></b>
			<%End If %>
		</td>
	</tr>
	<% rs.MoveNext %>
	<!-- Depósitos y Abonos -->	
	<tr bgcolor="#fbfbfb" height="20" align="right" style="FONT-SIZE: 8pt; COLOR: #303030">
		<td colspan="4">
			<b><%=rs("Descripcion")%></b>
		</td>
		<td>&nbsp;
			<%If ccur(0 & rs("monto_afex")) <> 0 Then %>
			<% nSaldo = nSaldo - ccur(0 & rs("monto_afex")) %>
				<b><FONT color="red"><%=FormatNumber(cCur(0 & rs("monto_afex")), 2)%></FONT></b>
			<%End If %>
		</td>
		<td>&nbsp;
			<%If ccur(0 & rs("monto_agente")) <> 0 Then %>
			<% nSaldo = nSaldo + ccur(0 & rs("monto_agente")) %>
				<b><FONT color=mediumblue><%=FormatNumber(cCur(0 & rs("monto_agente")), 2)%></FONT></b>
			<%End If %>
		</td>
	</tr>
	
	<tr height="20" bgcolor="white" >
		<td colspan="4"></td>
		<td bgcolor="#fbfbfb" ></td>
		<td bgcolor="#fbfbfb" ></td>		
	<tr>
	<tr >
		<td colspan="4" style="FONT-SIZE: 8pt">
			<b><FONT color=black size=2 >Giros Enviados</FONT> </b>
		</td>
		<td bgcolor="#fbfbfb" >&nbsp;</td>
		<td bgcolor="#fbfbfb" >&nbsp;</td>
	</tr>
	<tr bgcolor="#f1f1f1" height="20" align="middle">
		<td WIDTH="10%">
			<b>Cantidad</b>
		</td>
		<td WIDTH="20%">
			<b>Monto</b>
		</td>
		<td WIDTH="15%">
			<b>Tarifa Cliente</b>
		</td>
		<td WIDTH="15%">
			<b>Comisión</b>
		</td>
		<td bgcolor="#fbfbfb" >&nbsp;</td>
		<td bgcolor="#fbfbfb" >&nbsp;</td>
	</tr>
	<% rs.MoveNext %>
	<!-- Giros Enviados -->
	<tr bgcolor="#fbfbfb" height="20" align="right" style="FONT-SIZE: 8pt; COLOR: #303030">
		<td>
			<b><%=FormatNumber(cCur(0 & rs("Cantidad")), 0)%></b>
		</td>
		<td>
			<b><%=FormatNumber(cCur(0 & rs("Monto")), 2)%></b>
		</td>
		<td>
			<b><%=FormatNumber(cCur(0 & rs("Tarifa")), 2)%></b>
		</td>
		<td>
			<b><%=FormatNumber(cCur(0 & rs("Comision_Captador")), 2)%></b>
		</td>
		<td>&nbsp;
			<%If ccur(0 & rs("monto_afex")) <> 0 Then %>
			<% nSaldo = nSaldo - ccur(0 & rs("monto_afex")) %>
				<b><FONT color="red"><%=FormatNumber(cCur(0 & rs("monto_afex")), 2)%></FONT></b>
			<%End If %>
		</td>
		<td>&nbsp;
			<%If ccur(0 & rs("monto_agente")) <> 0 Then %>
			<% nSaldo = nSaldo + ccur(0 & rs("monto_agente")) %>
				<b><FONT color=mediumblue><%=FormatNumber(cCur(0 & rs("monto_agente")), 2)%></FONT></b>
			<%End If %>
		</td>
	</tr>
	
	
	<tr height="20" bgcolor="white" >
		<td colspan="4"></td>
		<td bgcolor="#fbfbfb" ></td>
		<td bgcolor="#fbfbfb" ></td>		
	<tr>
	<tr >
		<td colspan="4" style="FONT-SIZE: 8pt">
			<b><FONT color=black size=2 >Giros Anulados</FONT> </b>
		</td>
		<td bgcolor="#fbfbfb" >&nbsp;</td>
		<td bgcolor="#fbfbfb" >&nbsp;</td>
	</tr>
	<tr bgcolor="#f1f1f1" height="20" align="middle">
		<td WIDTH="10%">
			<b>Cantidad</b>
		</td>
		<td WIDTH="20%">
			<b>Monto</b>
		</td>
		<td WIDTH="15%">
			<b>Tarifa Cliente</b>
		</td>
		<td WIDTH="15%">
			<b>Comisión</b>
		</td>
		<td bgcolor="#fbfbfb" >&nbsp;</td>
		<td bgcolor="#fbfbfb" >&nbsp;</td>
	</tr>
	<% rs.MoveNext %>
	<!-- Giros Anulados -->
	<tr bgcolor="#fbfbfb" height="20" align="right" style="FONT-SIZE: 8pt; COLOR: #303030">
		<td>
			<b><%=FormatNumber(cCur(0 & rs("Cantidad")), 0)%></b>
		</td>
		<td>
			<b><%=FormatNumber(cCur(0 & rs("Monto")), 2)%></b>
		</td>
		<td>
			<b><%=FormatNumber(cCur(0 & rs("Tarifa")), 2)%></b>
		</td>
		<td>
			<b><%=FormatNumber(cCur(0 & rs("Comision_Captador")), 2)%></b>
		</td>
		<td>&nbsp;
			<%If ccur(0 & rs("monto_afex")) <> 0 Then %>
			<% nSaldo = nSaldo - ccur(0 & rs("monto_afex")) %>
				<b><FONT color="red"><%=FormatNumber(cCur(0 & rs("monto_afex")), 2)%></FONT></b>
			<%End If %>
		</td>
		<td>&nbsp;
			<%If ccur(0 & rs("monto_agente")) <> 0 Then %>
			<% nSaldo = nSaldo + ccur(0 & rs("monto_agente")) %>
				<b><FONT color=mediumblue><%=FormatNumber(cCur(0 & rs("monto_agente")), 2)%></FONT></b>
			<%End If %>
		</td>
	</tr>
	
	
	
	<tr height="20" bgcolor="white" >
		<td colspan="4"></td>
		<td bgcolor="#fbfbfb" ></td>
		<td bgcolor="#fbfbfb" ></td>		
	<tr>
	<tr>
		<td colspan="4" style="FONT-SIZE: 8pt">
			<b><FONT color=black size=2>Giros Recibidos</FONT> </b>
		</td>
		<td bgcolor="#fbfbfb" >&nbsp;</td>
		<td bgcolor="#fbfbfb" >&nbsp;</td>
	</tr>
	<tr bgcolor="#f1f1f1" height="20" align="middle">
		<td>
			<b>Cantidad</b>
		</td>
		<td>
			<b>Monto</b>
		</td>
		<td>&nbsp;</td>
		<td>
			<b>Comisión</b>
		</td>
		<td bgcolor="#fbfbfb" >&nbsp;</td>
		<td bgcolor="#fbfbfb" >&nbsp;</td>
	</tr>
	<% rs.MoveNext %>
	<!-- Giros Recibidos -->
	<tr bgcolor="#fbfbfb" height="20" align="right" style="FONT-SIZE: 8pt; COLOR: #303030">
		<td>
			<b><%=FormatNumber(cCur(0 & rs("Cantidad")), 0)%></b>
		</td>
		<td>
			<b><%=FormatNumber(cCur(0 & rs("monto")), 2)%></b>
		</td>
		<td>&nbsp;</td>
		<td>
			<b><%=FormatNumber(cCur(0 & rs("comision_pagador")), 2)%></b>
		</td>
		<td>&nbsp;
			<%If ccur(0 & rs("monto_afex")) <> 0 Then %>
			<% nSaldo = nSaldo - ccur(0 & rs("monto_afex")) %>
				<b><FONT color="red"><%=FormatNumber(cCur(0 & rs("monto_afex")), 2)%></FONT></b>
			<%End If %>
		</td>
		<td>&nbsp;
			<%If ccur(0 & rs("monto_agente")) <> 0 Then %>
			<% nSaldo = nSaldo + ccur(0 & rs("monto_agente")) %>
				<b><FONT color=mediumblue><%=FormatNumber(cCur(0 & rs("monto_agente")), 2)%></FONT></b>
			<%End If %>
		</td>
	</tr>
	<tr height="20" bgcolor="white" >
		<td colspan="4"></td>
		<td bgcolor="#fbfbfb" ></td>
		<td bgcolor="#fbfbfb" ></td>		
	<tr>
	<tr bgcolor="#fbfbfb" height="20" align="right" style="FONT-SIZE: 8pt; COLOR: #303030">
		<td colspan="4">
			<b>Saldo Actual</b>
		</td>
		<td>&nbsp;
			<%If nSaldo < 0 Then %>
				<b><FONT color="red"><%=FormatNumber(nSaldo * -1, 2)%></FONT></b>
			<%End If %>
		</td>
		<td>&nbsp;
			<%If nSaldo >= 0 Then %>
				<b><FONT color=mediumblue><%=FormatNumber(nSaldo, 2)%></FONT></b>
			<%End If %>
		</td>
	</tr>
	<tr height="40" bgcolor="white" align="center">
		<td colspan="6">
			<img id="imgAceptar" src="../images/BotonAceptar.jpg" style="CURSOR: hand" WIDTH="70" HEIGHT="20">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			<img Id="imgImprimir" src="../images/BotonImprimir.jpg" style="CURSOR: hand" WIDTH="70" HEIGHT="20">
		</td>
	</tr>
	</table>
</td></tr>
<tr height="10"><td></td></tr>
<tr id="trPie" style="display: none; font-size: 10pt" ><td>AFEX Ltda.<br>Santiago - Chile</td></tr>
</table>
</div>
<input type="hidden" value="<%=nSaldo%>" name="SaldoActual">
</form>
</body>
<!--#INCLUDE virtual="/Compartido/Rutinas.htm" -->
<script>
	Set rs = Nothing
	
	Sub objConsulta_OnScriptletEvent(strEventName, varEventData)
		Dim sAgtCaptador, sAgtPagador
		
	   Select Case strEventName
			
			Case "Aceptar"
				
						window.navigate "CuentaCorriente.asp?Desde=" & CtaCte.objConsulta.Desde & _
						"&Hasta=" & CtaCte.objConsulta.Hasta 
		End Select
		
	End Sub
	
	Sub document_ondblclick()
		msgbox document.body.innerHTML 
	End Sub
</script>

</html>
