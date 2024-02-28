<%@ LANGUAGE = VBScript %>
<%
	'Expiracion
	'Se asegura que la página no se almacene en la memoria cache
	Response.Expires = 0
	If Session("CodigoCliente") = "" Then
		response.Redirect "TimeOut.htm"
		response.end
	End If

	'Variables
	Dim sVigencia, rMoneda
		
	'Obtener Paridades
	Set rMoneda = ObtenerMonedasTransfer()
	
	'Obtener vigencia
	sVigencia = ValidarVigencia()
	
	
%>
<!--#INCLUDE virtual="/compartido/Transferencias.asp" -->
<!--#INCLUDE virtual="/compartido/Rutinas.asp" -->
<!--#INCLUDE virtual="/compartido/Errores.asp" -->
<html>
<head>
<title>Paridades</title>
<link rel="stylesheet" type="text/css" href="../Estilos/Cliente.css">
</head>

<script language="VBScript">
<!--
	Sub txtRecargoOpr_onChange()		
		' si está vacío le asigna 0
		If txtRecargoOpr.value = Empty Then txtRecargoOpr.value = "0"
		
		' calcula el recargo operacional
		CalcularRecargo "PO", "RO", txtRecargoOpr.value
				
		' convierte los valores a pesos
		ConvertirParidades
	End Sub
	
	Sub txtRecargoTrf_onChange()		
		' si está vacío le asigna 0
		If txtRecargoTrf.value = Empty Then txtRecargoTrf.value = "0"
		
		' calcula el recargo transferencia
		CalcularRecargo "PT", "RT", txtRecargoTrf.value
		
		' convierte los valores a pesos
		ConvertirParidades
	End Sub
	
	Sub cmdPesos_onClick()
		' si el valor dolar es Empty o 0 se sale
		If txtValorDolar.value = Empty OR txtValorDolar.value = "0" Then Exit Sub
		
		' muestra la tabla de valores en pesos
		If document.all.item("tbParidadesPesos").style.display <> "none" Then		
			document.all.item("tbParidadesPesos").style.display = "none"
		Else
			document.all.item("tbParidadesPesos").style.display = ""
			window.tbParidadesPesos.focus()			
		End If
		
		' convierte los valores a pesos
		ConvertirParidades
	End Sub
	
	Sub txtValorDolar_onChange()
		' oculta la tabla de paridades en pesos
		document.all.item("tbParidadesPesos").style.display = "none"
	End Sub
	
	Sub cmdImprimir_onClick()
		Dim sUsuario
		
		sUsuario = "<%=session("CodigoCliente")%>"
		
		' muestra el reporte		
		window.open  "../Reportes/Paridades.rpt?init=actx" & _
						"&prompt0=" & sUsuario & _
						"&prompt1=" & txtRecargoOpr.value & _
						"&prompt2=" & txtRecargoTrf.value & _
						"&prompt3=" & txtValorDolar.value & _
						"&user0=Cambios&password0=Cambios", _
						"", "dialogHeight= 250pxl; dialogWidth= 250pxl; " & _
					    "dialogTop= 0; dialogLeft= 0; " & _
						"status=no; scrollbars=no"		
	End Sub
	
	'**********************************Procedimientos***************************************
	'Objetivo:		calcular el recargo de una de las columnas de paridad
	'Parámetros:	Paridad, código de la columna de paridad a la que se le aplica el recargo	
	'				Recargo, código de la columna de recar en donde se asigna el resultado
	'				Porcentaje, porcentaje de recargo
	Private Sub CalcularRecargo(ByVal Paridad, ByVal Recargo, Byval Porcentaje)
		Dim i
		Dim cRecargo
		Dim cParidad
		
		' calcula el recargo para cada fila de la tabla
		For i = 1 To window.tbParidades.rows.length - 1		
			' calcula el % de recargo
			cParidad = cdbl(window.tbParidades.cells(Paridad & i).innertext)
			cRecargo = (cdbl(cParidad) * cdbl(Porcentaje)) / 100
			
			' asigna el monto original + el recargo
			window.tbParidades.cells(Recargo & i).innertext = Round(cdbl(cParidad) - cdbl(cRecargo), 4)
			window.tbParidades.cells(Recargo & i).innertext = window.tbParidades.cells(Recargo & i).innertext _
															  & "   "
		Next
	End Sub
	
	'Objetivo:	Convertir las paridades de Dolar a Pesos con el txtValorDolar
	Private Sub ConvertirParidades
		Dim i
		Dim cParidad
		Dim cRecargo
		
		' verfiica que la tabla de paridades pesoso esté visible
		If document.all.item("tbPAridadesPesos").style.display = "none" Then Exit Sub
		
		' traspasa los valores en dolares a pesos
		For i = 1 To window.tbParidades.rows.length - 1
			'**Paridad Operacional**
			' convierte la paridad a dolar y luego a pesos
			cParidad = cdbl(window.tbParidades.cells("PO" & i).innertext)
			If cdbl(cParidad) > 0 Then 						
				cParidad = cdbl(1) / cdbl(cParidad)
				cParidad = Round(cdbl(cParidad) * ccur(txtValorDolar.value), 4)
			End If
			
			' asigna el monto en pesos
			window.tbParidadesPesos.cells("PO" & i).innertext = cParidad & "   "
			'**Recargo Operacional
			' convierte el recargo a dolar y luego a pesos
			cRecargo = cdbl(window.tbParidades.cells("RO" & i).innertext)
			If cdbl(cRecargo) > 0 Then				
				cRecargo = cdbl(1) / cdbl(cRecargo)
				cRecargo = Round(cdbl(cRecargo) * ccur(txtValorDolar.value), 4)
			End If
			
			' asigna el monto en pesos
			window.tbParidadesPesos.cells("RO" & i).innertext = cRecargo & "   "
			'**Fin**
			
			'**Paridad Transferencia**			
			' convierte la paridad a dolar y luego a pesos
			cParidad = cdbl(window.tbParidades.cells("PT" & i).innertext)
			If cdbl(cParidad) > 0 Then				
				cParidad = cdbl(1) / cdbl(cParidad)
				cParidad = Round(cdbl(cParidad) * ccur(txtValorDolar.value), 4)			
			End If
			
			' asigna el monto en pesos
			window.tbParidadesPesos.cells("PT" & i).innertext = cParidad & "   "
			'**Recargo Transferencia			
			' convierte el recargo a dolar y luego a pesos
			cRecargo = cdbl(window.tbParidades.cells("RT" & i).innertext)
			If cdbl(cRecargo) > 0 Then				
				cRecargo = cdbl(1) / cdbl(cRecargo)
				cRecargo = Round(cdbl(cRecargo) * ccur(txtValorDolar.value), 4)
			End If
			
			' asigna el monto en pesos
			window.tbParidadesPesos.cells("RT" & i).innertext = cRecargo & "   "
			'**Fin**
		Next		
	End Sub
	'***************************************Fin*********************************************
-->
</script>

	<body background="../agente/imagenes/Giros_FondoVentana.jpg" bgcolor="#FFFFFF" text="#008080" link="#0000FF" vlink="#000080">
<script LANGUAGE="VBScript">
<!--
	Const sEncabezadoFondo = "Información"
	Const sEncabezadoTitulo = "Paridades"
	Const sClass = "TituloPrincipal"
-->
</script>
	
<!--#INCLUDE virtual="/compartido/Encabezado.htm" -->		

		<font FACE="Verdana" SIZE="3">
			<div ALIGN="CENTER"><%=sVigencia%></div>
		</font>	
 
		<br>
		<table width="100%" border="0">
			<tr>
				<td style="font-family: Verdana; font-size: 14">
					Tipo Cambio<br>
					<input type="Text" style="text-align:right" name="txtValorDolar" value="0" size="6" maxlength="6">
					<input type="Button" name="cmdPesos" value=" $ ">
				</td>
				<td align="Right" style="font-family: Verdana; font-size: 14">
					&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;% Recargo&nbsp;&nbsp;&nbsp;
					&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;% Recargo<br>
					<input type="Text" style="text-align:right" name="txtRecargoOpr" value="0" size="6" maxlength="6">
					&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
					&nbsp;&nbsp;					
					<input type="Text" style="text-align:right" name="txtRecargoTrf" value="0" size="6" maxlength="6">
				</td>
			</tr>
		</table>
		<table ID="tbParidades" width="100%" border="0.5" ALIGN="CENTER" STYLE="font-family: Verdana; font-size: 13; color:#505050">
			<tr BORDER="1">
				<td WIDTH="25%" ALIGN="CENTER" BGCOLOR="#EAD094" STYLE="font-family: Verdana; font-size: 14">
					<b>Moneda</b>
				</td>
				<td WIDTH="20%" ALIGN="CENTER" BGCOLOR="#FCF6E9" STYLE="font-family: Verdana; font-size: 14">
					<b>Paridad Operacional</b>
				</td>
				<td WIDTH="20%" ALIGN="CENTER" BGCOLOR="#FCF6E9" STYLE="font-family: Verdana; font-size: 14">
					<b>Paridad Transferencia</b>
				</td>
				<td WIDTH="18%" ALIGN="CENTER" BGCOLOR="#FCF6E9" STYLE="font-family: Verdana; font-size: 14">
					<b>Recargo Operacional</b>
				</td>
				<td WIDTH="17%" ALIGN="CENTER" BGCOLOR="#FCF6E9" STYLE="font-family: Verdana; font-size: 14">
					<b>Recargo Transferencia</b>
				</td>
			</tr>
			<%
				i = 0
				Do Until rMoneda.EOF 
					i = ccur(i) + 1
			%>
					<tr>
						<td colSpan="0" rowSpan="0" BGCOLOR="#F0E4C0">&nbsp;&nbsp;<%=rMoneda("Alias_Moneda")%></td>
						<td ID="<%="PO" & i%>" BGCOLOR="#FBF7F0" ALIGN="RIGHT"><%=rMoneda("Paridad")%>&nbsp;&nbsp;</td>
						<% If cDbl(rMoneda("ParidadTransfer")) <> 0 Then %>
							<td ID="<%="PT" & i%>" BGCOLOR="#FBF7F0" ALIGN="RIGHT"><%=Round(cdbl(1) / cdbl(rMoneda("ParidadTransfer")), 7)%>&nbsp;&nbsp;</td>
						<% Else %>
							<td ID="<%="PT" & i%>" BGCOLOR="#FBF7F0" ALIGN="RIGHT"><%=rMoneda("ParidadTransfer")%>&nbsp;&nbsp;</td>
						<% End If %>
						<td ID="<%="RO" & i%>" BGCOLOR="#FBF7F0" ALIGN="RIGHT">0&nbsp;&nbsp;</td>
						<td ID="<%="RT" & i%>" BGCOLOR="#FBF7F0" ALIGN="RIGHT">0&nbsp;&nbsp;</td>
					</tr>
				<%
					rMoneda.MoveNext
				Loop
				%>
		</table>
		
		<!-- Tabla de Paridades en Pesos -->
		<br>		
		<table ID="tbParidadesPesos" width="100%" border="0.5" ALIGN="CENTER" STYLE="font-family: Verdana; font-size: 13; color:#505050; display:none">
			<font face="Verdana"><b>Valores en Pesos</b></font>			
			<tr BORDER="1">
				<td WIDTH="25%" ALIGN="CENTER" BGCOLOR="#EAD094" STYLE="font-family: Verdana; font-size: 14">
					<b>Moneda</b>
				</td>
				<td WIDTH="20%" ALIGN="CENTER" BGCOLOR="#FCF6E9" STYLE="font-family: Verdana; font-size: 14">
					<b>Paridad Operacional</b>
				</td>
				<td WIDTH="20%" ALIGN="CENTER" BGCOLOR="#FCF6E9" STYLE="font-family: Verdana; font-size: 14">
					<b>Paridad Transferencia</b>
				</td>
				<td WIDTH="18%" ALIGN="CENTER" BGCOLOR="#FCF6E9" STYLE="font-family: Verdana; font-size: 14">
					<b>Recargo Operacional</b>
				</td>
				<td WIDTH="17%" ALIGN="CENTER" BGCOLOR="#FCF6E9" STYLE="font-family: Verdana; font-size: 14">
					<b>Recargo Transferencia</b>
				</td>
			</tr>
			<%
				i = 0
				rMoneda.MoveFirst
				Do Until rMoneda.EOF 
					i = ccur(i) + 1
			%>
					<tr>
						<td colSpan="0" rowSpan="0" BGCOLOR="#F0E4C0">&nbsp;&nbsp;<%=rMoneda("Alias_Moneda")%></td>
						<td ID="<%="PO" & i%>" BGCOLOR="#FBF7F0" ALIGN="RIGHT"></td>						
						<td ID="<%="PT" & i%>" BGCOLOR="#FBF7F0" ALIGN="RIGHT"></td>						
						<td ID="<%="RO" & i%>" BGCOLOR="#FBF7F0" ALIGN="RIGHT"></td>
						<td ID="<%="RT" & i%>" BGCOLOR="#FBF7F0" ALIGN="RIGHT"></td>
					</tr>
				<%
					rMoneda.MoveNext
				Loop
				%>
		</table>
		<!-- Fin -->
		
		<%	
		Set rMoneda = Nothing
		%>
		
		<!-- se muestran los botones -->
		<center>
			<!--<a href="javascript:history.back()"><img src="../agente/imagenes/GIROS_Anterior.jpg" border="0" 			alt="Volver a la página Anterior" WIDTH="99" HEIGHT="80"></img></a>-->
			<input TYPE="image" NAME="cmdImprimir" SRC="../images/BotonImprimir.jpg" border="0" alt="Imprimir las Paridades" WIDTH="70" HEIGHT="20">
		</center>
	</body>
</html>