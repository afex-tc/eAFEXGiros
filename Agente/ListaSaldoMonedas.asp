<%@ Language=VBScript %>
<%
	'Se asegura que la página no se almacene en la memoria cache
	Response.Expires = 0
	If Session("CodigoCliente") = "" Then
		response.Redirect "../Compartido/TimeOut.htm"
		response.end
	End If
%>
<!-- #INCLUDE virtual="/Compartido/Constantes.asp" -->
<!-- #INCLUDE virtual="/Compartido/Errores.asp" -->
<!-- #INCLUDE virtual="/Compartido/Rutinas.asp" -->
<%

	'Variables
	Dim rsMoneda, afxWeb

	On Error Resume Next	

	'Set afxWeb = Server.CreateObject("AFEXWebXp.Web")	
	'Set rsMoneda = afxWeb.ObtenerMonedasSaldo(Session("afxCnxAFEXweb"), Session("CodigoCaja"))
	Set rsMoneda = ObtenerMonedasSaldo(Session("afxCnxAFEXweb"), Session("CodigoCaja"))
	
	If Err.number <> 0 Then
		Set rsMoneda = Nothing
		'Set afxWeb = Nothing
		MostrarErrorMS "Moneda 1"
	End If
	'If afxWeb.ErrNumber <> 0 Then			
	'	Set rsMoneda = Nothing
	'	MostrarErrorAFEX afxWeb, "Moneda 2"
	'End If
	
	'set afxWeb = Nothing
	
	Function ObtenerMonedasSaldo(ByVal Conexion, _
								 ByVal Caja)
		Dim sSQL
		      
		On Error Resume Next
		
		Set ObtenerMonedasSaldo = Nothing
		   
		sSQL = "SELECT DISTINCT alias_moneda, " & _
		       "       CASE WHEN codigo_internacional IS NULL " & _
		       "            THEN pm.codigo_moneda " & _
		       "            ELSE pm.codigo_moneda " & _
		       "       END AS codigo_moneda, " & _
		       "       pm.codigo_producto, " & _
		       "       tp.nombre_tipo AS nombre_producto, " & _
		       "       pm.saldo_anterior_extranjera, " & _
		       "       pm.ingresos_extranjera, " & _
		       "       pm.egresos_extranjera, " & _
		       "       pm.saldo_extranjera, " & _
		       "       pm.saldo_anterior_nacional, " & _
		       "       pm.ingresos_nacional, " & _
		       "       pm.egresos_nacional, " & _
		       "       pm.saldo_nacional " & _
		       "FROM   Moneda " & _
		       "JOIN   Plan_Moneda PM ON pm.codigo_moneda=moneda.codigo_moneda " & _
		       "JOIN   Tipo tp ON tp.codigo_tipo = pm.codigo_producto AND tp.nombre_campo = 'PRODUCTO' " & _
		       "WHERE  pm.codigo_caja = '" & Caja & "' " & _
		       "       AND (pm.saldo_nacional <> 0 OR pm.saldo_extranjera <> 0) " & _
		       "ORDER  BY alias_moneda "
		   
		'Asigna al metodo el resultado de la consulta
		Set ObtenerMonedasSaldo = EjecutarSQLCliente(Conexion, sSQL)
	   
		If Err.Number <> 0 Then
			MostrarErrorMS "Obtener Saldo Monedas"
		End If
	End Function


%>
<html>
<style TYPE="text/css">	
</style> 
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link rel="stylesheet" type="text/css" href="../Estilos/Cliente.css">
<link rel="stylesheet" type="text/css" href="../Estilos/Reportes.css">


</head>
<script LANGUAGE="VBScript">
<!--
	'Variables de módulo
	'Variables para encabezado
	Dim sEncabezadoFondo, sEncabezadoTitulo
	Dim rsMoneda, afxWeb, sColor1, sColor2
	sColor1 = "#F0F0F0"
	sColor2 = "#F6F6F6"
	
	sEncabezadoFondo = "Consultas"
	sEncabezadoTitulo = "Saldo de Monedas"
	
	Sub window_onload()
		On Error Resume Next
	End Sub
	
//-->
</script>

<body>
<!-- #INCLUDE virtual="/Compartido/Encabezado.htm" -->
<table ID="tbReporte"  cellpadding="1" cellspacing="1" border="0" ALIGN="center" STYLE          ="COLOR: #505050; FONT-FAMILY: Verdana; FONT-SIZE: 11px; POSITION: relative; TOP: 0px">
	<tr style="COLOR: white; HEIGHT: 30px" class="Encabezado">
		<td WIDTH="100" ALIGN="middle" >
			<b>Moneda</b>
		</td>
		<td WIDTH="70" ALIGN="middle" >
			<b>Producto</b>
		</td>
		<td WIDTH="100" ALIGN="middle" >
			<b>Saldo<br>Anterior</b>
		</td>
		<td WIDTH="100" ALIGN="middle" >
			<b>Ingresos</b>
		</td>
		<td WIDTH="100" ALIGN="middle" >
			<b>Egresos</b>
		</td>
		<td WIDTH="100" ALIGN="middle" >
			<b>Saldo<br>Actual</b>
		</td>
	</tr>
	<%
		Do Until rsMoneda.EOF 
	%>			
			<tr style="HEIGHT: 25px" language="javascript" onmouseover="javascript:this.bgColor='#a4dded'; window.status='<%=rsMoneda("alias_moneda")%>'" onmouseout="javascript:this.bgColor='#DAF6FF'; window.status=''" bgColor="#dbf7ff">
				<td ALIGN="left" ><%=rsMoneda("alias_moneda")%></td>
				<td ALIGN="left" ><%=rsMoneda("nombre_producto")%></td>
				<% If rsMoneda("codigo_moneda") = Session("MonedaNacional") Then %>
						<td ALIGN="right" ><%=FormatNumber(rsMoneda("saldo_anterior_nacional"), 2)%></td>
						<td ALIGN="right" ><%=FormatNumber(rsMoneda("ingresos_nacional"), 2)%></td>
						<td ALIGN="right" ><%=FormatNumber(rsMoneda("egresos_nacional"), 2)%></td>
						<td ALIGN="right" ><%=FormatNumber(rsMoneda("saldo_nacional"), 2, , -1)%></td>
				<% Else %>
						<td ALIGN="right" ><%=FormatNumber(rsMoneda("saldo_anterior_extranjera"), 2)%></td>
						<td ALIGN="right" ><%=FormatNumber(rsMoneda("ingresos_extranjera"), 2)%></td>
						<td ALIGN="right" ><%=FormatNumber(rsMoneda("egresos_extranjera"), 2)%></td>
						<td ALIGN="right" ><%=FormatNumber(rsMoneda("saldo_extranjera"), 2, , -1)%></td>				
				<% End If %>
			</tr>
	<%
			rsMoneda.MoveNext
		Loop
		Set rsMoneda = Nothing
	%>
</table>
</body>
</html>
