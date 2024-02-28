<%@ Language=VBScript %>
<%
	'Se asegura que la página no se almacene en la memoria cache
	Response.Expires = 0
	If Session("CodigoCliente") = "" Then
		response.Redirect "http:../Compartido/TimeOut.htm"
		response.end
	End If
%>
<%

	'Variables de módulo
	'Variables para encabezado
	Dim sEncabezadoFondo
	Dim sEncabezadoTitulo
	sEncabezadoFondo = ""
	sEncabezadoTitulo = "Agentes"

	'Objetivo:     Obtiene los agentes pagadores de una ciudad
	'Parámetros:   Conexion, conexión a la base
	'              Pais, código del país de pago
	'              Ciudad, código de la ciudad de pago
	Function ObtenerAgentePagador()
	   Dim rsAgente
	   Dim sSQL, Pais, Ciudad
	   Pais = Request("pa")
	   Ciudad = Request("ci")
	   
	   Set ObtenerAgentePagador = Nothing
	   
	   On Error Resume Next
	   
	   'Crea la consulta
	   sSQL = "SELECT    DISTINCT ag.* " & _
	          "FROM      agente AG " & _
	          "JOIN      comision CM ON ag.codigo_agente = cm.codigo_agente " & _
	          "WHERE     (cm.pais = '" & Pais & "' " & _
	          "  OR      cm.pais = '**') " & _
	          " AND      (cm.ciudad = '" & Ciudad & "' " & _
	          "  OR      cm.ciudad = '***') " & _
	          " AND      cm.sentido = " & afxSentido.afxEnviado & _
	          " AND      ag.estado_agente <> 0 " & _
	          " AND      cm.fecha_termino is null "
	   
	   'Ejecuta la consulta
	   Set rsAgente = EjecutarSQLCliente(Session("afxCnxAFEXpress"), sSQL)
	   
	   'Si se produjeron errores en la consulta
	   If Err.Number <> 0 Then 
	   
	   End If
	      
	   'Si no existe el agente
	   If rsAgente.EOF Then 
	   
	   End If
	   
	   Set ObtenerAgentePagador = rsAgente
	   Set rsAgente = Nothing
	End Function
	
%> 
<!--#INCLUDE virtual="/Compartido/Errores.asp" -->
<!--#INCLUDE virtual="/Compartido/Rutinas.asp" -->
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
	

	Sub window_onload()
	End Sub
	
	
//-->
</script>
<body>
<!--#INCLUDE virtual="/Compartido/Encabezado.asp" -->
<table border="0" cellspacing="0" cellpadding="0" style="LEFT: 20px; POSITION: relative; TOP: 30px">
<tr height="10"><td>
	<table cellspacing="1" cellpadding="1" ID="tbReporte" border="0" ALIGN="center" STYLE="COLOR: #505050; FONT-FAMILY: Verdana; FONT-SIZE: 10px; POSITION: relative; TOP: 0px">
	<tr CLASS="Encabezado">
		<td WIDTH="150">
			<b>Agente</b>
		</td>
		<td WIDTH=150">
			<b>Direccion</b>
		</td>
		<td WIDTH="80">
			<b>Telefono</b>
		</td>
		<td WIDTH="80">
			<b>Fax</b>
		</td>
		<td WIDTH="100">
			<b>Horario</b>
		</td>
	</tr>
<%
		Dim rs, nCantidad
		nCantidad = 0
		Set rs = ObtenerAgentePagador()

		Do Until rs.EOF 
			
%>		
			<a onclick="">
<!--			<tr CLASS="<%=sDetalle%>" sonmouseover="Seleccionar()" sonmouseout="QuitarSeleccion()" >-->
				<!-- <td onmouseover="window.status ='=rsCheque("nombre_completo")'">=rsCheque("nombre_completo")</td> -->
			<tr style="HEIGHT: 25px" language="javascript" onmouseover="javascript:this.bgColor='#a4dded'" onmouseout="javascript:this.bgColor='#DAF6FF'" bgColor="#dbf7ff" style="cursor: hand">
				<td><%=rs("nombre_agente")%></td>
				<td><%=rs("direccion_agente")%></td>
				<td><%=rs("fono_agente")%></td>
				<td><%=rs("fax_agente")%></td>
				<td><%=rs("horario_atencion")%></td>
				<!--<td><%=sNegocio%></td>-->
			</tr>
			</a>
<%
			nCantidad = nCantidad + 1
			rs.MoveNext
		Loop
%>
		<tr style="height: 20px" CLASS="Encabezado">
			<td colspan="1"  ALIGN="left"><b><%=FormatNumber(nCantidad, 0)%>&nbsp;Agentes</b></td>			
		</tr>
<%
		Set rs = Nothing
%>
</table>
</td></tr>
</table>
</body>
</html>
