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
	'Variables para encabezado	
	Dim sEncabezadoFondo
	Dim sEncabezadoTitulo
	
	Dim sSQL
	Dim rs
		
	sSQL = " exec MostrarTipoCambioVentaAgentesInternacionales "	
	SET rs = EjecutarSQLCliente(Session("afxCnxAfexpress"), sSQL)
	if err.number <> 0 then
		set rs = nothing
		mostrarerrorms "T/C Agentes Internacionales"
	end if
	
	sEncabezadoFondo = "Consultas"
	sEncabezadoTitulo = "T/C Agentes Internacionales"
	
%>
<!--#INCLUDE virtual="/compartido/Rutinas.asp" -->
<!--#INCLUDE virtual="/compartido/Errores.asp" -->
<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link rel="stylesheet" type="text/css" href="../Estilos/Principal.css">
<link rel="stylesheet" type="text/css" href="../Estilos/Cliente.css">
<link rel="stylesheet" type="text/css" href="../Estilos/Reportes.css">
</head>

<body>
<form name="frmTC" method="post" action="">
	<!--#INCLUDE virtual="/Compartido/Encabezado.asp" -->

	<br />
	<br />
	<center>
		<table width="50%">
			<tr CLASS="Encabezado">
				<td width="20%" align="Center"><b>Moneda</b></td>
				<td width="15%" align="Center"><b>Valor</b></td>
				<td width="75%" align="Center"><b>Agente</b></td>			
			</tr>
			<%Do While not rs.eof%>
				<tr bgColor="#dbf7ff" style="cursor: hand">
					<td align="Center"><%=rs("codigo_moneda")%></td>
					<td align="Right"><%=formatnumber(rs("valor"),2)%></td>
					<td align="Left"><%=rs("nombre_agente")%></td>			
				</tr>
				<%rs.movenext%>	
			<%Loop%>
		</table>
	</center>
</form>	
</body>
</html>
<%set rs = nothing%>