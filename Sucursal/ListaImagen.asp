<%@ Language=VBScript %>
<%
	'Se asegura que la página no se almacene en la memoria cache
	Response.expires = 0

	If Session("CodigoCliente") = "" Then
		response.Redirect "../Compartido/TimeOut.htm"
		response.end
	End If
%>
<!--#INCLUDE virtual="/Compartido/Constantes.asp" -->
<%

	'Variables de módulo
	'Variables para encabezado
	Dim sEncabezadoFondo
	Dim sEncabezadoTitulo
	
	Dim sTitulo, nTipo
	Dim nCampo, sArgumento, sArgumento2, sArgumento3, rs, nAccion
		
	sTitulo = Request("Titulo")
	nAccion = cInt(0 & Request("Accion"))
	
	If Trim(sTitulo) = "" Then sTitulo = "Imágenes"
	sEncabezadoFondo = "Documentos"
	sEncabezadoTitulo = sTitulo
		
	Set rs = ObtenerCCP()	
	If rs.EOF Then
		Set rs = Nothing
	End If
	
	
	Function ObtenerCCP()
	   Dim rsATC
	   Dim sSQL
	   Dim Condicion
	   Dim Rut
	   Dim Raya
	   Dim sComa

	   Set ObtenerCCP = Nothing

	   On Error Resume Next
		sSQL = "SELECT dc.tipo, CASE WHEN dc.numero = 0 THEN '' ELSE dc.numero END AS Numero, td.nombre AS Nombre_Documento, " & _
				 "			'http:../intraimg/documentos/' + RTRIM(ISNULL(cl.rut, '')) + RTRIM(ISNULL(cl.pasaporte, '')) + '_' + CAST(dc.tipo AS varchar) + '_' + dc.numero + '.jpg' AS nombre_archivo " & _
				 "FROM Documento_Cliente dc " & _
				 "LEFT OUTER JOIN Cliente cl ON cl.codigo=dc.codigo_cliente " & _
				 "LEFT OUTER JOIN Tipo_Documento td ON dc.tipo = td.codigo " & _
				 "WHERE codigo_cliente=" & Request("cc") & " AND dc.tipo IN (" & Request("td") & ") " & _
				 "ORDER BY dc.tipo, dc.numero "
	   'sSQL = "SELECT nombre_archivo FROM documento_cliente WHERE codigo_cliente=" & Request("cc") & " AND tipo_documento=" & Request("td")
	   sComa = ""
	   'mostrarerrorms ssql
	   
	   	   
	   Set rsATC = EjecutarSQLCliente(Session("afxCnxCorporativa"), sSQL)

	   If Err.Number <> 0 Then 
			Set rsATC = Nothing
			MostrarErrorMS "Obtener CCP"
		End If
	   
	   Set rsATC.ActiveConnection = Nothing
	   Set ObtenerCCP = rsATC

	   Set rsATC = Nothing
	End Function

	'Response.Redirect "../Compartido/Error.asp?Titulo=" & rs.eof	

%>
<!--#INCLUDE virtual="/Compartido/Rutinas.asp" -->
<!--#INCLUDE virtual="/Compartido/Errores.asp" -->

<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<title>AFEX</title>
<link rel="stylesheet" type="text/css" href="../Estilos/Cliente.css">
<link rel="stylesheet" type="text/css" href="../Estilos/Reportes.css">
</head>
<script LANGUAGE="VBScript">
<!--
	Dim sOldClass

	Sub window_onload()
	End Sub		

	Sub Seleccionar()

		'msgbox window.event.srcElement.tagname
		sOldClass = window.event.srcElement.className 
		window.event.srcElement.className  = "Seleccionado"

	End Sub
	
	Sub QuitarSeleccion()
	
		'msgbox window.event.srcElement.tagname
		window.event.srcElement.className  = sOldClass
		
	End Sub
		
//-->
</script>
<body>
<!--#INCLUDE virtual="/Compartido/Encabezado.asp" -->

<table border="0" cellspacing="0" cellpadding="0" style="LEFT: 0px; POSITION: relative; TOP: 0px">
<tr height="10"><td>
	<table cellspacing="1" cellpadding="1" ID="tbReporte" border="0" ALIGN="center" STYLE="COLOR: #505050; FONT-FAMILY: Verdana; FONT-SIZE: 10px; POSITION: relative; TOP: 0px"  width="780">
	<%
		Dim i, nCodigo, sDetalle, sNegocio, sPagina
				
		Select Case nAccion
		Case afxAccionIngresarMG
		Case Else
		End Select
		'mostrarerrorms nAccion & ", " & sPagina
		i = 0
		If rs Is Nothing Then
			%>
				<tr><td style="font-size: 14pt"><br><br>&nbsp;&nbsp;&nbsp;Este cliente no tiene imágenes disponibles</td></tr>
			<%
		Else
			Do Until rs.EOF 
					i = ccur(i) + 1		
				%>		
					<!--<a href="<%=sPagina%>?Accion=<%=afxAccionBuscar%>&Campo=<%=nCodigo%>&Argumento=<%=sCodigoCliente%>">-->
					<!--<tr CLASS="<%=sDetalle%>" sonmouseover="Seleccionar()" sonmouseout="QuitarSeleccion()" >-->
					<tr align="left" CLASS="Encabezado" style="height: 10; sbackground-color: skyblue">
						<td WIDTH="250">
							<b><%=rs("nombre_documento")%>&nbsp;<%=rs("numero")%></b>
						</td>
					</tr>
					<tr>
						<td align="center"><img src="<%=rs("nombre_archivo")%>" border="0" width="90%"></a></td>
					</tr>
					<tr height="40"><td></td></tr>
					<!--</a>-->
				<%
				rs.MoveNext
			Loop
		End If
		%>
	</table>
</td></tr>
</table>
</body>
</html>
<%
	Set rs = Nothing
%>