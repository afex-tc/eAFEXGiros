<%@ Language=VBScript %>
<!-- #INCLUDE virtual="/Compartido/Constantes.asp" -->
<!-- #INCLUDE virtual="/Compartido/Errores.asp" -->
<!--#INCLUDE virtual="/Compartido/Rutinas.asp" -->
<%
	'Se asegura que la página no se almacene en la memoria cache
	'Response.Expires = 0
	'If session("CodigoCliente") = "" Then
	'	response.Redirect "../Compartido/TimeOut.htm"
	'	response.end
	'End If

	'Variables
	Dim rsMoneda, afxWeb, sCnxTC

	On Error Resume Next	
	'Set afxWeb = Server.CreateObject("AFEXWebXP.Web")
	'Set rsMoneda = afxWeb.ObtenerMonedasParidad(Session("afxCnxTipoCambio"))
	Set rsMoneda = ObtenerMonedasParidad
	
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

	Function ObtenerMonedasParidad()
		Dim sSQL
	      
		On Error Resume Next
	   
		Set ObtenerMonedasParidad = Nothing
	   
		sSQL = "SELECT DISTINCT alias_moneda,   CASE WHEN codigo_internacional IS NULL " & _
	           "                       THEN pm.codigo_moneda " & _
	           "                       ELSE pm.codigo_moneda " & _
	           "                       END AS codigo_moneda, " & _
	           "       pm.paridad_transferencia AS ParidadTransfer, " & _
	           "       pm.tipo_cambio_compra AS Compra, " & _
	           "       pm.tipo_cambio_venta AS Venta, " & _
	           "       pm.tipo_cambio_paridad AS Paridad, " & _
	           "       moneda.codigo_pais " & _
	           "FROM   Moneda " & _
	           "JOIN   Plan_Moneda PM ON pm.codigo_moneda=moneda.codigo_moneda " & _
	           "WHERE  pm.codigo_producto = 1 AND " & _
	           "       pm.codigo_caja = '0000' " & _
	           "ORDER  BY alias_moneda "
	   
		Set ObtenerMonedasParidad = EjecutarSQLCliente(Session("afxCnxTipoCambio"), sSQL)
	   
		If Err.Number <> 0 Then 
			MostrarErrorMS "Obtener Monedas Paridad"
		End If
	End Function

%>
<html>
<style TYPE="text/css">
TD.td1	{	BACKGROUND-COLOR: #f0f0f0}
TD.td2	{	BACKGROUND-COLOR: #F6F6F6}

A:active
{
    background-color: gray
}
A:link
{
    background-color: gray
}
A:visited
{
    background-color: gray
}
A:hover
{
    background-color: blue
}
	
</style> 
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<% 
	Select Case cInt(Request("Tipo"))
	Case afxPrincipal
%>
		<link rel="stylesheet" type="text/css" href="../Estilos/Principal.css">
		<link rel="stylesheet" type="text/css" href="../Estilos/Reportes.css">
<% Case Else %>
		<link rel="stylesheet" type="text/css" href="../Estilos/Cliente.css">
		<link rel="stylesheet" type="text/css" href="../Estilos/Reportes.css">
<% End Select %>


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
	sEncabezadoTitulo = "Tipos de Cambio"
	
	Sub window_onload()
		On Error Resume Next
	End Sub
	
//-->
</script>

<body>
<!-- #INCLUDE virtual="/Compartido/Encabezado.htm" -->
<table ID="tbReporte"  cellpadding="1" cellspacing="1" border="0" ALIGN="center" STYLE          ="COLOR: #505050; FONT-FAMILY: Verdana; FONT-SIZE: 11px; POSITION: relative; TOP: 0px">
	<tr style="COLOR: ; HEIGHT: 30px" bgcolor="#F0E4C0" sclass="Encabezado" sbgcolor="#ffcc66">
		<td WIDTH="200" ALIGN="middle" sbgcolor="#d0d0d0" >
			<b>Moneda</b>
		</td>
		<td WIDTH="100" ALIGN="middle" sbgcolor="#e0e0e0" >
			<b>Compra</b><br>
		</td>
		<td WIDTH="100" ALIGN="middle" sbgcolor="#d0d0d0" >
			<b>Venta</b><br>
		</td>
	</tr>
	<%
		Do Until rsMoneda.EOF 
			If ccur(rsMoneda("compra")) > 0 Or ccur(rsMoneda("venta")) > 0 Then
				If sDetalle <> "Detalle1" Then
					sDetalle = "Detalle1"
				Else
					sDetalle = "Detalle2"
				End If
	%>			
				<!--<tr style="height: 25" CLASS="<%=sDetalle%>" language="javascript" onmouseover="javascript:this.bgColor='#e0f8e0' ">-->
				<tr style="HEIGHT: 25px" language="javascript" onmouseover="javascript:this.bgColor='#F8F2e5'" onmouseout="javascript:this.bgColor='#FBF7F0'" bgColor ="#FBF7F0">			
					<td ALIGN="left" sclass="td1" >&nbsp;&nbsp;<%=rsMoneda("alias_moneda")%></td>
					<% If ccur(rsMoneda("compra")) > 0 Then %>
						<td ALIGN="right" sclass="td2"><%=FormatNumber(rsMoneda("compra"), 2)%></td>
					<% Else %>
						<td ALIGN="right" sclass="td2">0,00</td>
					<% End If
					   If ccur(rsMoneda("venta")) > 0 Then %>
						<td ALIGN="right" sclass="td1"><%=FormatNumber(rsMoneda("venta"), 2)%></td>
					<% Else %>
						<td ALIGN="right" sclass="td1">0,00</td>
					<% End If
		    End If %>		
				</tr>
	<%
			rsMoneda.MoveNext
		Loop
		Set rsMoneda = Nothing
	%>
	<tr>
		<td COLSPAN="4" style="font-size: 8pt">*Precios referenciales sujetos a cambio sin notificación.<br>*Precios y paridades no válidas para cantidades superiores o equivalentes a U$5.000.<br>*Consulte en AFEX Casa Matriz o en su sucursal AFEX más cercana.</td>
	</tr>
</table>
</body>
</html>
