<%@ Language=VBScript %>
<%
	Response.Expires = 0
	Response.buffer = true	
	If Not Session("SesionActiva") Then
'		response.Redirect "../Compartido/TimeOut.htm"
		response.end
	End If

%>
<%
	
%>
<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<title>Apertura de Puntas <%=Request("cmn")%></title>
<link rel="stylesheet" type="text/css" href="../Compartido/VisorAP.css">
</head>
<!--#INCLUDE virtual="/afexvisorap/visorap/Rutinas.htm" -->
<!--#INCLUDE virtual="/afexvisorap/Compartido/Boton.htm" -->
<script LANGUAGE="VBScript">
<!--
	Dim nTCCompra, nTCVenta, nTCTrf
	
	nTCCompra = cCur(0 & "<%=Request("tcc")%>")
	nTCVenta = cCur(0 & "<%=Request("tcv")%>")
	nTCUSDVenta = cCur(0 & "<%=Request("tcusv")%>")
	nTCTrf = cCur(0 & "<%=Request("tct")%>")

	Sub window_onload()
		Cargar
		CalcularTodas 
	End Sub
	
	Sub window_onunload()
	
	End Sub
	
	Sub Cargar()
	   Dim rs, sSQL, sHTML, sChecked
			   
	   On Error Resume Next

		sSQL = "APAperturaPunta " & EvaluarStr("<%=Request("mn")%>")
		
		Set rs = EjecutarSQLCliente("<%=Session("cnxVisorAP")%>", sSQL)			   
		If Err.Number <> 0 Then 
		
			Set rs = Nothing
			Exit Sub
		End If	   

		sHTML = "<table id=tbAP border=0 cellspacing=1 cellpadding=1 style=""display: ; border: 1px solid '#EEEEEE'"">"
		sHTML =  sHTML & "<tr class=""titulo"" align=center class=titulo style=""height: 26px"">" & _
					 "<td width=180px>" & rs.Fields(0).name & "</td>" & _
					 "<td width=80px>" & rs.Fields(1).name & "</td>" & _
					 "<td width=80px>" & rs.Fields(2).name & "</td>" & _
					 "<td width=80px>" & rs.Fields(3).name & "</td>" & _
					 "<td width=80px>" & "Compra" & "</td>" & _					 
					 "<td width=80px>" & "Venta" & "</td>" & _					 
					 "</tr>"			      
	   Do Until rs.EOF
			If rs.Fields(3) = 1 Then
				sChecked = "checked"
			Else
				sChecked = ""
			End If 
			sHTML = sHTML & _
						 "<tr style=""height: 22px; cursor: hand; "" color=#BBBBBB bgcolor=#F6F6F6 onMouseOver=""javascript:this.bgColor='#EEEEEE'"" onMouseOut=""javascript:this.bgColor='#F6F6F6'"" >" & _
						 "<td align=left id=tdPr" & rs.Fields(4) & ">" & rs.Fields(0) & "</td>" & _
						 "<td align=center style=""color: blue""><input id=txtCompra" & rs.Fields(4) & " value=" & FormatNumber(rs.Fields(1), 2) & " style=""width: 40px"" onKeyPress=""IngresarNumero"" onBlur=""CalcularCompra " & rs.Fields(4) & ", txtCompra" & rs.Fields(4) & ".value,  tdCompra" & rs.Fields(4) & " ""></td>" & _
						 "<td align=center style=""color: red""><input id=txtVenta" & rs.Fields(4) & " value=" & FormatNumber(rs.Fields(2), 2) & " style=""width: 40px"" onKeyPress=""IngresarNumero"" onBlur=""CalcularVenta " & rs.Fields(4) & ", txtVenta" & rs.Fields(4) & ".value,  tdVenta" & rs.Fields(4) & " ""></td>" & _
						 "<td align=center ><input type=checkbox id=chkVisor" & rs.Fields(4) & " " & sChecked & " style=""width: 40px""></td>" & _
						 "<td align=center id=tdCompra" & rs.Fields(4) & " style=""color: blue""></td>" & _					 
						 "<td align=center id=tdVenta" & rs.Fields(4) & " style=""color: red""></td>" & _					 
						 "</tr>"
			   
	      rs.MoveNext
	   Loop
	   rs.Close		   
		sHTML = sHTML & "</table>"
		dvMonedas.innerHTML = sHTML
	   Set rs = Nothing
			   
	End Sub
	
	Sub CalcularTodas
		CalcularCompra 0, txtCompra0.value, tdCompra0
		CalcularCompra 1, txtCompra1.value, tdCompra1
		CalcularCompra 2, txtCompra2.value, tdCompra2
		CalcularCompra 3, txtCompra3.value, tdCompra3
		CalcularCompra 4, txtCompra4.value, tdCompra4
		CalcularCompra 5, txtCompra5.value, tdCompra5
		CalcularCompra 6, txtCompra6.value, tdCompra6
		CalcularCompra 7, txtCompra7.value, tdCompra7
		CalcularCompra 8, txtCompra8.value, tdCompra8

		CalcularVenta 0, txtVenta0.value, tdVenta0
		CalcularVenta 1, txtVenta1.value, tdVenta1
		CalcularVenta 2, txtVenta2.value, tdVenta2
		CalcularVenta 3, txtVenta3.value, tdVenta3
		CalcularVenta 4, txtVenta4.value, tdVenta4
		CalcularVenta 5, txtVenta5.value, tdVenta5
		CalcularVenta 6, txtVenta6.value, tdVenta6
		CalcularVenta 7, txtVenta7.value, tdVenta7
		CalcularVenta 8, txtVenta8.value, tdVenta8
	End Sub

	Sub CalcularCompra(ByVal Producto, ByVal Paridad, ByRef Objeto)
		Dim nValor
		If Paridad = "" Then Paridad = 0
		
		nValor = Round(CDbl(0 & nTCCompra) - (CDbl(0 & nTCCompra) * Paridad ) / 100, 2)
		Objeto.innertext = FormatNumber(nValor, 2)
		
	End Sub

	Sub CalcularVenta(ByVal Producto, ByVal Paridad, ByRef Objeto)
		Dim nValor
		If Paridad = "" Then Paridad = 0
				
		Select Case Producto
		Case 2, 3, 4
				'nValor = CDbl(0 & nTCTrf) + (CDbl(0 & nTCTrf) * Paridad) / 100
				nValor = Round(CDbl(0 & (nTCUSDVenta / (1 / cCur("<%=Request("ptr")%>")))) + (CDbl(0 & (nTCUSDVenta / (1 / cCur("<%=Request("ptr")%>")))) * Paridad) / 100, 2)
		Case Else
				nValor = Round(CDbl(0 & nTCVenta) + (CDbl(0 & nTCVenta) * Paridad) / 100, 2)		
		End Select
		Objeto.innertext = FormatNumber(nValor, 2)		
		
	End Sub

	Sub Actualizar
		If MsgBox("Está seguro que desea guardar los cambios?", vbYesNo+vbQuestion) <> vbYes Then Exit Sub

		If Not ActualizarPR(0, tdPr0.innerText, txtCompra0.value, txtVenta0.value, chkVisor0.checked) Then Exit Sub
		If Not ActualizarPR(1, tdPr1.innerText, txtCompra1.value, txtVenta1.value, chkVisor1.checked) Then Exit Sub
		If Not ActualizarPR(2, tdPr2.innerText, txtCompra2.value, txtVenta2.value, chkVisor2.checked) Then Exit Sub
		If Not ActualizarPR(3, tdPr3.innerText, txtCompra3.value, txtVenta3.value, chkVisor3.checked) Then Exit Sub
		If Not ActualizarPR(4, tdPr4.innerText, txtCompra4.value, txtVenta4.value, chkVisor4.checked) Then Exit Sub
		If Not ActualizarPR(5, tdPr5.innerText, txtCompra5.value, txtVenta5.value, chkVisor5.checked) Then Exit Sub
		If Not ActualizarPR(6, tdPr6.innerText, txtCompra6.value, txtVenta6.value, chkVisor6.checked) Then Exit Sub
		If Not ActualizarPR(7, tdPr7.innerText, txtCompra7.value, txtVenta7.value, chkVisor7.checked) Then Exit Sub
		If Not ActualizarPR(8, tdPr8.innerText, txtCompra8.value, txtVenta8.value, chkVisor8.checked) Then Exit Sub
		
		window.close 	
		
	End Sub

	Function ActualizarPR(ByVal Producto, ByVal Alias, ByVal Compra, ByVal Venta, ByVal Visor)
		Dim bOk
		Dim nVisor

	   If Visor Then
			nVisor = 1
	   Else
			nVisor = 0
	   End If
		
		bOk = ActualizarAPPro("<%=Session("cnxVisorAP")%>", "<%=Request("mn")%>", Producto, _
							 Compra, Venta, nVisor)
							 
		If Not bOk Then 
			MsgBox "Se produjo un error al intentar actualizar el producto " & Alias, vbCritical
		End If
		
		ActualizarPR = bOk
	End Function
	

	Function ActualizarAPPro(ByVal Conexion, ByVal Moneda, ByVal Producto, _
									 ByVal Compra, ByVal Venta, ByVal Visor)
	   Dim sSQL
	   Dim BD
	   
	   On Error Resume Next
	   ActualizarAPPro = False
	   
	   sSQL = "DELETE APApertura_Punta " & _
	          "WHERE moneda = " & EvaluarStr(Moneda) & " AND producto = " & Producto
	          
	   sSQL = sSQL & vbCrLf & " " & _
	          "INSERT APApertura_Punta " & _
	          "    ( moneda, producto, compra, venta, visor ) " & _
	          "VALUES " & _
	          "    (" & EvaluarStr(Moneda) & ", " & Producto & ", " & FormatoNumeroSQL(Compra) & ", " & FormatoNumeroSQL(Venta) & ", " & Visor & ") "
	   
		'Conexion
		Set BD = CreateObject("ADODB.Connection")
		BD.Open Conexion                          'Abre la conexion
		If Err.Number <> 0 Then
			Set BD = Nothing
			Exit Function
		End If
		   
		'Consulta
		BD.BeginTrans
		BD.Execute sSQL                           'Ejecuta la consulta
		If Err.Number <> 0 Then 
			BD.RollbackTrans 
			BD.Close 
			Set BD = Nothing
			Exit Function
		End If
		   
		ActualizarAPPro = True
		BD.CommitTrans
		BD.Close
		Set BD = Nothing		
	End Function
//-->
</script>
<body style="margin-left: 10px; margin-top: 10px">
<div id="dvMonedas" style>
	espere...
</div>
<center>
<table>
<tr height="20px"><td></td></tr>
<tr align="center" style="height: 20px;">
	<td class="boton" style="width: 100px;" onClick="Actualizar" onMouseOver="MouseOver()" onMouseOut="MouseOut()">Aceptar</td>
	<td class="boton" style="width: 100px;" onClick="window.close" onMouseOver="MouseOver()" onMouseOut="MouseOut()">Salir</td>
</tr>
</table>
</center>
</body>
</html>
