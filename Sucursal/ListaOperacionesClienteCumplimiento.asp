<%@ Language=VBScript %>
<!-- #Include virtual="/Compartido/Rutinas.asp"-->
<!-- #Include virtual="/Compartido/Errores.asp"-->
<%
	
	dim sSQL
	dim rs
	dim sFechaDesde, sFechaHasta

	sFechaDesde = Request.Form("txtFechaDesde")
	sFechaHasta = Request.Form("txtFechaHasta")
	if sFechaDesde = "" then sFechaDesde = date
	if sFechaHasta = "" then sFechaHasta = date
	
	' consulta el listado de operaciones para el día seleccionado
	sSQL = " exec cumplimiento.mostraroperacionclienteautorizacion " & evaluarstr(formatofechasql(sFechaDesde)) & ", " & evaluarstr(formatofechasql(sFechaHasta))
	set rs = ejecutarsqlcliente(Session("afxCnxCorporativa"), ssql)
	if err.number <> 0 then
		MostrarErrorMS "Mostrar Operaciones"
	end if

	Response.Expires = 0
%>

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link rel="stylesheet" type="text/css" href="../Estilos/Cliente.css">
<link rel="stylesheet" type="text/css" href="../Estilos/AFX2004.css">
</HEAD>

<script language="vbscript">
<!--
	sub cmdBuscar_onClick()
		frm.action = "ListaOperacionesClienteCumplimiento.asp"
		frm.submit() 
		frm.action = ""
	end sub
-->
</script>

<BODY>

	
	<form name="frm" method="post" action="">
		<table class="Borde" id="" BORDER="0" cellpadding="0" cellspacing="0" style="HEIGHT: 150px; width:100%; background-color: #f4f4f4">
			<tr height="40" style="background-color: #ffeeaa; #ffdd77; #e1e1e1">
				<td colspan="3" style="font-size: 16pt">&nbsp;&nbsp;Lista Operaciones Cumplimiento</td>
			</tr>
			<tr height="15"><td colspan="3" ></td></tr>
			<tr>
				<td align="center">
					<table class="BordeSombra" cellpadding="0" cellspacing="0" border="0" width="210" bgcolor="#f1f1f1" style="background-color: white; color: ">						
						<tr height="22" bgcolor="#ccddee" style="font-size: 10pt; font-weight: bold">
							<td colspan="5">&nbsp;&nbsp;Periodo de Búsqueda</td>
						</tr>
						<tr><td>&nbsp;</td></tr>
						<tr>
							<td align="center">				
								<input size="10" type="text" name="txtFechaDesde" value="<%=sFechaDesde%>">&nbsp;-&nbsp;
								<input size="10" type="text" name="txtFechaHasta" value="<%=sFechaHasta%>">
							</td>
						</tr>
						<tr><td>&nbsp;</td></tr>
						<tr>
							<td align="center">
								<input type="button" name="cmdBuscar" value="Buscar">
							</td>			
						</tr>
						<tr><td>&nbsp;</td></tr>
					</table>
				</td>
			</tr>
			<tr><td>&nbsp;</td></tr>
			<tr><td>&nbsp;</td></tr>
			<tr>
				<td>	
					<table  cellspacing="1" cellpadding="1" ID="tbReporte" border="0" ALIGN="center" STYLE="background-color: silver; COLOR: #505050; FONT-FAMILY: Verdana; FONT-SIZE: 10px; POSITION: relative; TOP: 0px">
						<tr CLASS="Encabezado" style="background-color: #e1e1e1; height: 25px" align="center">
							<td><b>Fecha</b></td>
							<td><b>Monto Dolates</b></td>
							<td><b>Monto FC Dolares</b></td>
							<td><b>Identificación Cliente</b></td>
							<td><b>Usuario Autoriza</b></td>
							<td><b>Sucursal</b></td>
						</tr>
						
						<%do while not rs.eof%>
							<tr  style="HEIGHT: 25px" language="javascript" onmouseover="javascript:this.bgColor='#f1f1f1';" onmouseout="javascript:this.bgColor='white'; window.status=''" bgColor="white" style="cursor: hand">
								<td><%=rs("fecha")%></b></td>
								<td align="right"><%=formatnumber(rs("MontoDolares"),2)%></b></td>
								<td align="right"><%=formatnumber(rs("MontoFCDolares"),2)%></b></td>
								<td><%=rs("IdentificacionCliente") & " " & rs("PaisIdentificacionCliente") %></b></td>
								<td><%=rs("UsuarioAutoriza")%></b></td>
								<td><%=rs("codigosucursal")%></b></td>
							</tr>
							<%rs.movenext%>
						<%loop%>
					</table>
				</td>
			</tr>
		</table>
	</form>
	
</BODY>
</HTML>
<% set rs = nothing%>