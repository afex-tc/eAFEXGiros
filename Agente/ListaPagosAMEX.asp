<%@ Language=VBScript %>
<!-- #INCLUDE virtual="/Compartido/Rutinas.asp" -->
<!-- #INCLUDE virtual="/Compartido/Errores.asp" -->
<%
	dim sSQL, rs, sFechaDesde, sFechaHasta, sEncabezadoFondo, sEncabezadoTitulo
	
	sFechaDesde = Request.Form("txtFechaDesde")
	sFechaHasta = Request.Form("txtFechaHasta")
	if sFechaDesde = "" then sFechaDesde = Date
	if sFechaHasta = "" then sFechaHasta = Date
	
	sSQL = "exec mostrarpagoamex " & evaluarstr(sFechaDesde) & ", " & evaluarstr(sFechaHasta) & ", " & evaluarstr(Session("CodigoAgente"))
	set rs = ejecutarsqlcliente(Session("afxCnxAFEXchangeMCF"), sSQL)
	if err.number <> 0 then
		mostrarerrorms "Buscar Pagos. 1"
	end if	

	sEncabezadoFondo = "Consultas"
	sEncabezadoTitulo = "Lista de Pagos AMEX"

	Response.Expires = 0
%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<title></title>
<link rel="stylesheet" type="text/css" href="../Estilos/Reportes.css">
<link rel="stylesheet" type="text/css" href="../Estilos/Cliente.css">
</HEAD>
<BODY>
	
	<script language="vbscript">
	<!--
		sub window_onload()
			frmLista.objConsulta.Desde = Replace("<%=sFechaDesde%>", "/", "-")
			frmLista.objConsulta.Hasta = Replace("<%=sFechaHasta%>", "/", "-")
			frmLista.objConsulta.Tipo = 1			
		end sub
		
		Sub objConsulta_OnScriptletEvent(strEventName, varEventData)
			Dim sAgtCaptador, sAgtPagador
			
			<% If nTipo = afxGirosRecibidos Then %>
					sAgtPagador = "<%=sPagador%>"
					sAgtCaptador = Trim(objConsulta.Captador) 
			<% ElseIf nTipo = afxGirosEnviados Then %>
					sAgtCaptador = "<%=sCaptador%>"
					sAgtPagador = Trim(objConsulta.Pagador)
			<% Else %>
					sAgtCaptador = "<%=sCaptador%>"
					sAgtPagador = "<%=sPagador%>"
			<% End If %>			

		   Select Case strEventName
				
				Case "Aceptar"
					frmLista.txtFechaDesde.value = frmLista.objConsulta.Desde
					frmLista.txtFechaHasta.value = frmLista.objConsulta.Hasta
				
					frmLista.action = "ListaPagosAMEX.asp"
					frmLista.submit()
					frmLista.action = ""							
			End Select
			
		End Sub
	-->
	</script>

	<form name="frmLista" method="post" action="">	
		<table align="center">
			<tr>
				<td align="left">
					<!--#INCLUDE virtual="/Compartido/Encabezado.asp" -->				
					<OBJECT id=objConsulta style="HEIGHT: 240px; LEFT: 0px; TOP: 0px; WIDTH: 544px" type=text/x-scriptlet width=544 VIEWASTEXT>
						<PARAM NAME="Scrollbar" VALUE="0">
						<PARAM NAME="URL" VALUE="http:ConfiguracionConsulta.asp"></OBJECT>			
					
					<input type="hidden" name="txtFechaDesde">
					<input type="hidden" name="txtFechaHasta">
					
					<table STYLE="COLOR: #505050; FONT-FAMILY: Verdana; FONT-SIZE: 10px; POSITION: relative; TOP: 0px">				
						<tr class="Encabezado">			
							<td><b>Fecha</b></td>
							<td><b>Tarjeta</b></td>
							<td><b>Monto Dólares</b></td>
							<td><b>T/C</b></td>
							<td><b>Monto Pesos</b></td>
							<td><b>Nombre Titular</b></td>
							<td><b>Estado</b></td>
						</tr>
						
						<%do while not rs.eof%>
							<a href="agregarpagoamex.asp?Accion=1&Pago=<%=rs("codigo")%>">
								<tr style="HEIGHT: 25px;" onmouseover="javascript:this.bgColor='#a4dded';" bgColor="#dbf7ff" onmouseout="javascript:this.bgColor='#DAF6FF';" style="cursor: hand">
									<td><b><%=rs("fecha")%></b></td>
									<td><b><%=rs("numerotarjeta")%></b></td>
									<td align="right"><b><%=formatnumber(rs("montodolares"),2)%></b></td>
									<td align="right"><b><%=formatnumber(rs("tipocambio"),2)%></b></td>
									<td align="right"><b><%=formatnumber(rs("montopesos"),0)%></b></td>
									<td><b><%=rs("nombretitular")%></b></td>
									<td><b><%=rs("nombreestado")%></b></td>					
								</tr>
							</a>
							<%rs.MoveNext%>
						<%loop%>
					</table>		
				</td>
			</tr>
		</div>	
	</form>
</BODY>
</HTML>
<%set rs = nothing%>
