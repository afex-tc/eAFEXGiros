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
	Dim sSucursal
	Dim nCodigoCliente
	
	Dim sTitulo, nTipo
	Dim nCampo, sArgumento, sArgumento2, sArgumento3, rs, nAccion
		
	sTitulo = Request("Titulo")
	nAccion = cInt(0 & Request("Accion"))
	nCodigoCliente = cCur(0 & Request("cc"))
	nDiasRetencion = cInt(0 & Request("dr"))
	
	If Trim(sTitulo) = "" Then sTitulo = "Detalle de Cheques"
	sEncabezadoFondo = "Consultas"
	sEncabezadoTitulo = sTitulo

	nCampo = cInt(0 & Request("Campo"))
	sArgumento = request("Argumento")
	sArgumento2 = request("Argumento2")
	sArgumento3 = request("Argumento3")		
	sSucursal = Request("sc")
	
	If nCodigoCliente <> 0 Then
		Set rs = ObtenerCheques(nCodigoCliente, nDiasRetencion)	
		If rs.EOF Then
			Set rs = Nothing
		End If
	Else
		Set rs = Nothing
	End If
	
	Function ObtenerCheques(ByVal CodigoCliente, ByVal DiasRetencion)
	   Dim rsCheques
	   Dim sSQL
	   Dim Condicion
	   Dim Rut
	   Dim Raya
	   Dim sComa

	   Set ObtenerCheques = Nothing

	   On Error Resume Next

	   sSQL = "SELECT * FROM detalle_solicitud " & _
			  "WHERE codigo_cliente_corporativa = " & CodigoCliente & " " & _
			  "and   fecha >= '" & CalcularFecha(Date, DiasRetencion, 0) & "' " & _
		      "and   codigo_producto = 2 " & _
			  "and   estado = 1"

	   Set rsCheques = EjecutarSQLCliente(Session("afxCnxCorporativa"), sSQL)

	   If Err.Number <> 0 Then 
			Set rsCheques = Nothing
			MostrarErrorMS "Obtener Cheques"
		End If
	   
	   Set rsCheques.ActiveConnection = Nothing
	   Set ObtenerCheques = rsCheques

	   Set rsCheques = Nothing
	End Function

	'Response.Redirect "../Compartido/Error.asp?Titulo=" & rs.eof	

%>
<!--#INCLUDE virtual="/Compartido/Rutinas.asp" -->
<!--#INCLUDE virtual="/sucursal/Rutinas.asp" -->
<!--#INCLUDE virtual="/Compartido/Errores.asp" -->

<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<title>AFEX</title>
<link rel="stylesheet" type="text/css" href="../Estilos/Cliente.css">
<link rel="stylesheet" type="text/css" href="../Estilos/AFX2004.css">
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

	'Sub cmdAceptar_onClick()
	'	window.navigate "ListaClientes.asp?sc=" & cbxSucursal.value
	'End Sub
		
//-->
</script>
<body id="bb" border="0" style="margin: 2 2 2 2" >
	<table class="Borde" id="" BORDER="0" cellpadding="0" cellspacing="0" style="HEIGHT: 150px; width:100%; background-color: #f4f4f4">	
	<tr height="40" style="background-color: #ffeeaa; #ffdd77; #e1e1e1"><td colspan="3" style="font-size: 16pt">&nbsp;&nbsp;Detalle de Cheques</td></tr>
	<tr height="1" style="background-color: silver "><td colspan="3" ></td></tr>
	<tr height="4"><td colspan="3" ></td></tr>
	
	<!--<table border="0" cellspacing="0" cellpadding="0" style="LEFT: 0px; POSITION: relative; TOP: 0px">-->

<tr height="10"><td>
	<table cellspacing="1" cellpadding="1" ID="tbReporte" border="0" ALIGN="center" STYLE="background-color: silver; COLOR: #505050; FONT-FAMILY: Verdana; FONT-SIZE: 10px; POSITION: relative; TOP: 0px">
	<tr CLASS="Encabezado" style="background-color: #e1e1e1; height: 25px" align="center">
		<td WIDTH="80">
			<b>Fecha</b>
		</td>
		<td WIDTH="80">
			<b>Nº Cheque</b>
		</td>
		<td WIDTH="100">
			<b>Código Moneda</b>
		</td>
		<td align="center" WIDTH="150">
			<b>Monto</b>
		</td>
		<td WIDTH="100">
			<b>Ejecutivo</b>
		</td>
	</tr>
	<%
		Dim i, nTotal
				
		'mostrarerrorms nAccion & ", " & sPagina
		i = 0
		nTotal = 0
		If rs Is Nothing Then
		Else		
			Do Until rs.EOF 
					i = ccur(i) + 1		
				%>		
					<!--<a href="<%=sPagina%>?Accion=<%=afxAccionBuscar%>&Campo=<%=nCodigo%>&Argumento=<%=sCodigoCliente%>">-->
					<!--<tr CLASS="<%=sDetalle%>" sonmouseover="Seleccionar()" sonmouseout="QuitarSeleccion()" >-->
					<tr style="HEIGHT: 25px" language="javascript" onmouseover="javascript:this.bgColor='#f1f1f1'; window.status=''" onmouseout="javascript:this.bgColor='white'; window.status=''" bgColor="white" style="cursor: hand">
							<td align="center"><%=rs("fecha")%></td>
							<td align="right"><%=rs("numero_producto")%></td>
							<td align="center"><%=rs("codigo_moneda")%></td>
							<td align="right"><%=FormatNumber(EvaluarVar(rs("monto"), 0), 2)%></td>
							<td><%=rs("codigo_usuario")%></td>
					</tr>
					<!--</a>-->
				<%
				nTotal = nTotal + cCur(0 & EvaluarVar(rs("monto"), 0))
				rs.MoveNext
			Loop
		End If
		%>
		<tr>
			<td></td>
			<td></td> 
			<td align="right">Total</td> 
			<td align="right"><%=FormatNumber(nTotal, 2)%></td> 
			<td></td>
		</tr>
	</table>
	<br>
</td></tr>
</table>
</body>
</html>
<%
	Set rs = Nothing
%>