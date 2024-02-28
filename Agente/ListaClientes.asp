<%@ Language=VBScript %>
<%
	'Se asegura que la página no se almacene en la memoria cache
	Response.expires = 0
	'Response.expiresabsolute = Now() - 1
	'Response.addHeader "pragma", "no-cache"
	'Response.addHeader "cache-control", "private"
	'Response.CacheControl = "no-cache"

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
	
	Dim sTipoAgente
	
	sTipoAgente = Request("ag")
		
	sTitulo = Request("Titulo")
	nAccion = cInt(0 & Request("Accion"))
	
	If Trim(sTitulo) = "" Then sTitulo = "Giros"
	sEncabezadoFondo = "Consultas"
	sEncabezadoTitulo = sTitulo

	nCampo = cInt(0 & Request("Campo"))
	sArgumento = request("Argumento")
	sArgumento2 = request("Argumento2")
	sArgumento3 = request("Argumento3")		
		
	Set rs = BuscarCliente(nCampo, sArgumento, sArgumento2, sArgumento3)	
	If rs.EOF Then
		rs.Close
		Set rs = Nothing
	End If
				
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
	<table cellspacing="1" cellpadding="1" ID="tbReporte" border="0" ALIGN="center" STYLE="COLOR: #505050; FONT-FAMILY: Verdana; FONT-SIZE: 10px; POSITION: relative; TOP: 0px">
	<tr CLASS="Encabezado">
		<td WIDTH="300">
			<b>Nombre</b>
		</td>
		<td WIDTH="100">
			<b>Rut</b>
		</td>
		<td WIDTH="100">
			<b>Pasaporte</b>
		</td>
		<td WIDTH="100">
			<b>Ciudad</b>
		</td>
		<td WIDTH="100">
			<b>Negocio</b>
		</td>
        <td WIDTH="100">
			<b>Código Corporativa</b>
		</td>
	</tr>
	<%
		Dim i, nCodigo, sDetalle, sNegocio, sPagina
				
		Select Case nAccion
		Case afxAccionIngresarMG
			sPagina = "IngresarGiroMG.asp"
			
		Case Else
			sPagina = "AtencionClientes.asp"
		End Select
		'mostrarerrorms nAccion & ", " & sPagina
		i = 0
		Do Until rs.EOF 
			If Trim(rs("codigo_pais")) = Trim(Session("PaisCliente")) Then
				i = ccur(i) + 1		
				sNegocio = ""	
				If Not IsNull(rs("Express")) Then
					nCodigo = 6
					sCodigoCliente = rs("Express")
					sNegocio = "Giros"
				End If
				If Not IsNull(rs("Exchange")) Then
					nCodigo = 5
					sCodigoCliente = rs("Exchange")
					If sNegocio <> "" Then
						sNegocio = sNegocio & ", "
					End If
					sNegocio = sNegocio & "Cambios"
				End If
			%>		
				<a href="<%=sPagina%>?Accion=<%=afxAccionBuscar%>&Campo=<%=nCodigo%>&Argumento=<%=sCodigoCliente%>&ag=<%=sTipoAgente%>">
				<!--<tr CLASS="<%=sDetalle%>" sonmouseover="Seleccionar()" sonmouseout="QuitarSeleccion()" >-->
				<tr style="HEIGHT: 25px" language="javascript" onmouseover="javascript:this.bgColor='#a4dded'; window.status='<%=rs("nombre_completo")%>'" onmouseout="javascript:this.bgColor='#DAF6FF'; window.status=''" bgColor="#dbf7ff" style="cursor: hand">			
					<td><%=rs("nombre_completo")%></td>
					<td><%=FormatoRut(rs("rut"))%></td>
					<td><%=rs("pasaporte")%></td>
					<td><%=rs("ciudad")%></td>
					<td><%=sNegocio%></td>				
                    <td><%=rs("Exchange")%></td>				
				</tr>
				</a>
			<%
			End If
			rs.MoveNext
		Loop
		%>
	</table>
</td></tr>
</table>
</body>
<script>

	Sub objConsulta_OnScriptletEvent(strEventName, varEventData)
		
	   Select Case strEventName
	   
			Case "Aceptar"
				msgbox strEventName & ", " & objConsulta.Desde
				
		End Select
		
	End Sub
	
	
</script>

</html>
<%
	rs.Close
	Set rs = Nothing
%>