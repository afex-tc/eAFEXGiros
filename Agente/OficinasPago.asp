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
	Dim rsOficina, afxWeb

	On Error Resume Next	

	'Set afxWeb = Server.CreateObject("AFEXWebXp.Web")	
	'Set rsMoneda = afxWeb.ObtenerMonedasSaldo(Session("afxCnxAFEXweb"), Session("CodigoCaja"))
	Set rsOficina = ObtenerOficinasPago(Session("afxCnxAFEXpress"), request("CodPag"), request("PaisB"),request("CiudadB"))
	
	If Err.number <> 0 Then
		Set rsOficina = Nothing
		'Set afxWeb = Nothing
		MostrarErrorMS "Oficina Pago 1"
	End If
	'If afxWeb.ErrNumber <> 0 Then			
	'	Set rsMoneda = Nothing
	'	MostrarErrorAFEX afxWeb, "Moneda 2"
	'End If
	
	'set afxWeb = Nothing
Function ObtenerOficinasPago (ByVal Conexion, _
							  ByVal AgentePagador, ByVal Pais, ByVal Ciudad)
		Dim sSQL
		      
		On Error Resume Next
		
		Set ObtenerOficinasPago = Nothing
	if AgentePagador="" and ciudad="" then
		sSQL = "SELECT sector, pagador, " & _
		       "direccion, telefono, horario_atencion as horario, moneda " & _
			   "FROM   Manual " & _
			   "WHERE  codigo_pais = '" & pais & "' " & _
               "AND  activo = 1 " & _
		       "ORDER  BY sector "
	elseif AgentePagador="" and ciudad<>"" then	
		sSQL = "SELECT sector, pagador, " & _
		       "direccion, telefono, horario_atencion as horario, moneda " & _
			   "FROM   Manual " & _
			   "WHERE  codigo_pais = '" & pais & "' " & _
		       "       AND codigo_ciudad = '" & ciudad & "' " & _
               "AND  activo = 1 " & _
		       "ORDER  BY direccion "
	else 
		sSQL = "SELECT sector, pagador, " & _
		       "direccion, telefono, horario_atencion as horario, moneda " & _
			   "FROM   Manual " & _
			   "WHERE  codigo_agente = '" & AgentePagador & "' " & _
		       "       AND codigo_pais = '" & pais & "' " & _
		       "       AND codigo_ciudad = '" & ciudad & "' " & _
               "AND  activo = 1 " & _
		       "ORDER  BY direccion "
	end if	   
		'Asigna al metodo el resultado de la consulta
		Set ObtenerOficinasPago = EjecutarSQLCliente(Conexion, sSQL)
	   
		If Err.Number <> 0 Then
			MostrarErrorMS "Obtener Oficinas de Pago"
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


<TITLE>Oficinas de Pago</TITLE>
</head>
<script LANGUAGE="VBScript">
<!--
	'Variables de módulo
	'Variables para encabezado
	Dim sEncabezadoFondo, sEncabezadoTitulo
	Dim rsOficina, afxWeb, sColor1, sColor2
	sColor1 = "#F0F0F0"
	sColor2 = "#F6F6F6"
	
	sEncabezadoFondo = "Consultas"
	sEncabezadoTitulo = "Oficinas de Pago"
	
	Sub window_onload()
		On Error Resume Next
	End Sub
	
//-->
</script>
<BODY>
<!-- #INCLUDE virtual="/Compartido/Encabezado.htm" -->
<table ID="tbReporte"  cellpadding="1" cellspacing="1" border="0" ALIGN="center" STYLE="COLOR: #505050; FONT-FAMILY: Verdana; FONT-SIZE: 9px; POSITION: relative; TOP: 0px; LEFT: 0px">
	<tr ID="TrFondoPrint" Style="Font-Family:Verdana;Font-Size:30pt;Font-Weight:Bold; Display: none"> <td colspan=6> Oficinas de Pago </td> </tr>
	<tr style="COLOR: white; HEIGHT: 30px" class="Encabezado">
	    
		<td WIDTH="180" ALIGN="middle" >
			<b>Sector</b>
		</td>
		
		<td WIDTH="190" ALIGN="middle" >
			<b>Pagador</b>
		</td>
		
		<td WIDTH="240" ALIGN="middle" >
			<b>Dirección</b>
		</td>
		
		<td WIDTH="150" ALIGN="middle" >
			<b>Teléfono</b>
		</td>
		
		<td WIDTH="210" ALIGN="middle" >
			<b> Horario </b>
		</td>
		
		<td WIDTH="90" ALIGN="middle" >
			<b>  Moneda<BR>de<BR>Pago  </b>
		</td>
	</tr>
    <%
		Do Until rsOficina.EOF 
	%>			
			<tr style="HEIGHT: 25px" language="javascript" onmouseout="javascript:this.bgColor='#DAF6FF'; window.status=''" bgColor="#dbf7ff">
			 
						   
						<td ALIGN="Left" ><%=rsOficina("sector")%></td>
						<td ALIGN="Left" ><%=rsOficina("pagador")%></td>
						<td ALIGN="Left" ><%=rsOficina("direccion")%></td>
						<td ALIGN="right" ><%=rsOficina("telefono")%></td>
						<td WIDTH="210" ALIGN="Left" ><%=rsOficina("horario")%></td>
						<td ALIGN="center" ><%=rsOficina("moneda")%></td>
			
			</tr>
	<%
			rsOficina.MoveNext
		Loop
		Set rsOficina = Nothing
	%>
</table>

<table ALIGN="center" >
	<tr height=50>
	</tr>
	<tr height=50>
	<td height=50>
		<IMG id=imgImprimir style="CURSOR: hand" height=20 src="../images/BotonImprimir.jpg" width=70 ><br>
	</td>
	</tr>
</table>
<script language=vbScript>
Sub imgImprimir_OnClick()
        tdfondo.style.display = "none"
		trfondoPrint.style.display = ""
		tbReporte.border="1px solid gray"  
		imgimprimir.style.display = "none"
		window.print()
		imgimprimir.style.display = ""
		tdfondo.style.display = ""
		trfondoPrint.style.display = "none"
		tbReporte.border="0px solid gray"
End Sub
</script>

</BODY>
</HTML>
