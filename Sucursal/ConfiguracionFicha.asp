<%@ Language=VBScript %>
<%
	Response.Expires = 0
	If Session("CodigoCliente") = "" Then
		response.Redirect "../Compartido/TimeOut.htm"
		response.end
	End If
%>
<% 
	response.Expires = 0
	response.Buffer = True
	response.Clear 
%>
<!--#INCLUDE virtual="/Compartido/TimeOut.asp" -->
<!--#INCLUDE virtual="/Compartido/Rutinas.asp" -->
<!--#INCLUDE virtual="/Compartido/Errores.asp" -->
<!--#INCLUDE virtual="/sucursal/Rutinas.asp" -->
<%

	Dim rsFicha, rsProducto, rsProductoFicha
	
	Set rsFicha = ObtenerFichaCliente()
	Set rsProducto = ObtenerProducto()
	Set rsProductoFicha = ObtenerProductoFicha()
	

	Sub AgregarEncabezadoProducto	
		If rsProducto Is Nothing Then
		ElseIf rsProducto.EOF Then
		Else
			rsProducto.MoveFirst
			Response.Write "<td style=""background-color: #d1d1d1; font-size: 10pt;font-weight: bold"" width=""70px"">" & rsFicha("nombre_grupo") & "</td>"
			Do Until rsProducto.EOF
				Response.Write "<td style=""background-color: #e1e1e1;"" width=""70px"">" & rsProducto("nombre") & "</td>"
				rsProducto.MoveNext
			Loop
			rsProducto.MoveFirst
		End If
	End Sub

	Sub AgregarGrillaFicha
		Dim nCampo, sChecked, nGrupo
					
		If rsFicha Is Nothing Then
		ElseIf rsFicha.EOF Then
		Else
			nGrupo = 0
			Do Until rsFicha.EOF
				If nGrupo <> rsFicha("grupo") Then
					'response.Write "<tr style=""height: 10px; background-color: #d1d1d1""><td colspan=""8"" style=""font-weight: bold; color: white; font-size: 10pt"">" & rsFicha("nombre_grupo") & "</td></tr>"
					AgregarEncabezadoProducto	
					nGrupo = rsFicha("grupo")
				End If
				response.Write "<tr>"
				response.write  "<td width=150px style=""background-color: #e1e1e1;"">" & rsFicha("descripcion") & "</td>"
				If Not rsProductoFicha.EOF Then
					nCampo = cInt(rsProductoFicha("campo"))
					Do Until rsProductoFicha.EOF
						If rsProductoFicha("estado") = 1 Then
							sChecked = "checked"
						Else
							sChecked = ""
						End If
						Response.Write "<td align=""left"" style=""background-color: white"" width=""70px""><input type=""checkbox"" name=" & rsProductoFicha("campo") & "_" & rsProductoFicha("producto") & " " & sChecked & "><input name=txt" & rsProductoFicha("campo") & "_" & rsProductoFicha("producto") & " style='width: 50px; text-align: right' onblur=""txt" & rsProductoFicha("campo") & "_" & rsProductoFicha("producto") & ".value=   FormatNumber(txt" & rsProductoFicha("campo") & "_" & rsProductoFicha("producto") & ".value, 0)                "" onkeypress=""IngresarTexto(1)"" value=" & formatnumber(cCur(0 & EvaluarVar(rsProductoFicha("monto_desde"), 0)), 0) & " ></td>"
						rsProductoFicha.MoveNext
						If Not rsProductoFicha.EOF Then
							If nCampo <> cInt(rsProductoFicha("campo")) Then Exit Do
						End If
					Loop
				End If
				Response.Write "</tr>"
				rsFicha.MoveNext
			Loop
		End If
		Set rsFicha = Nothing
		Set rsProductoFicha = Nothing					
		Set rsProducto = Nothing
	End Sub
	
%>
<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<title>Configuración Ficha Cliente</title>
<link rel="stylesheet" type="text/css" href="../Estilos/Cliente.css">
<link rel="stylesheet" type="text/css" href="../Estilos/AFX2004.css">
</head>

<script LANGUAGE="VBScript">
<!--
	On Error Resume Next

	Sub window_onload()
	End Sub

	Sub GuardarCambios_onclick()
		frmFicha.action = "http:GuardarFichaProducto.asp"
		frmFicha.submit 
		frmFicha.action = ""
	End Sub
		
//-->
</script>
<!--#INCLUDE virtual="/Compartido/Rutinas.htm" -->
<body id="bb" border="0" style="margin: 2 2 2 2" >
<form id="frmFicha" method="post">
	<table class="Borde" id="" BORDER="0" cellpadding="0" cellspacing="0" style="HEIGHT: 150px; width:100%; background-color: #f4f4f4">	
	<tr height="40" style="background-color: #ffeeaa; #ffdd77; #e1e1e1"><td colspan="3" style="font-size: 16pt">&nbsp;&nbsp;Configuración Ficha Cliente</td></tr>
	<tr height="1" style="background-color: silver "><td colspan="3" ></td></tr>
	<tr height="4"><td colspan="3" ></td></tr>
	<tr>
		<td colspan="3">
		<table width="100%">
		<tr>
			<td>
				<table style="font-size: 8pt">
					<tr>
					</tr>
				</table>
			</td>
			<td align="right">
				<table>
				<tr>
					<td colspan="2">
						<!--#INCLUDE virtual="/sucursal/MenuConfiguracionFicha.htm" -->
					</td>
					</div>
				</tr>
				</table>
			</td>
		</tr>
		</table>
		</td>
	</tr>

	<tr height="10"><td></td></tr>

	<tr>
		<td colspan="3" align="center">
			<table cellpadding="6" cellspacing="1" style="background-color: silver;" >			
			<!--
			<tr  align="center">
				<td width="150px" style="background-color: #e1e1e1;" ></td>
				<%	'AgregarEncabezadoProducto %>
			</tr>
			-->
			<%	AgregarGrillaFicha %>
			</table>
		</td>
	</tr>


	<tr height="10"><td></td></tr>
	</table>
</form>
</body>
</html>