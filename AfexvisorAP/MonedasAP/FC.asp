<%@ Language=VBScript %>
<%
	'Response.Expires = 0
	'If Session("CodigoCliente") = "" Then
	'	response.Redirect "../Compartido/TimeOut.htm"
	'	response.end
	'End If
%>
<%

	Session("cnxVisorAP") = "Provider=SQLOLEDB.1;Password=cambios;User ID=cambios;Initial Catalog=cambios;Data Source=canelo;"
'	If Not Then
'		Response.Redirect "http:../Compartido/Error.asp?Titulo=VisorAP&Description=..."
'		Response.End 
'	End If

	Sub CargarVisor
		On Error Resume Next		
			Dim rs, sSQL, sHTML
			   
			On Error Resume Next

			sSQL = "SELECT    codigo_internacional AS Moneda, pm_compra AS Compra, pm_venta AS Venta, paridad_compra AS 'Paridad<br>Compra', paridad_venta AS 'Paridad<br>Venta', paridad_transferencia AS 'Paridad<br>Transfer' " & _
			       "FROM      VAPPrecio_referencia WITH(nolock) " & _
			       "WHERE     estado_ap = 1 AND estado = 1 AND producto = 1 " & _
			       "ORDER BY  orden_ap"
			Set rs = EjecutarSQLCliente(Session("cnxVisorAP"), sSQL)			   
			If Err.Number <> 0 Then 
				Set rs = Nothing
				Exit Sub
			End If	   
			
			
			Response.Write "<table id=tbMonedas border=0 cellspacing=1 cellpadding=1 style=""display: ; border: 1px solid '#EEEEEE'"">"
			
			Response.Write "<tr class=""titulo"" align=center class=titulo style=""height: 10px"">" & _
						 "<td width=80px>" & rs.Fields(0).name & "</td>" & _
						 "<td width=80px>" & rs.Fields(1).name & "</td>" & _
						 "<td width=80px>" & rs.Fields(2).name & "</td>" & _
						 "<td width=80px>" & rs.Fields(3).name & "</td>" & _
						 "<td width=80px>" & rs.Fields(4).name & "</td>" & _
						 "<td width=80px>" & rs.Fields(5).name & "</td>" & _
						 "</tr>"
			Response.Write  err.Description 
			Do Until rs.EOF
				Response.Write vbCrlf & _
							 "<tr style=""height: 10px; cursor: hand; "" color=#BBBBBB bgcolor=#F6F6F6 onMouseOver=""javascript:this.bgColor='#EEEEEE'"" onMouseOut=""javascript:this.bgColor='#F6F6F6'"" >" & _
							 "<td align=left >" & rs.Fields(0) & "</td>" & _
							 "<td align=right style=""color: blue"">" & FormatNumber(rs.Fields(1), 2) & "</td>" & _
							 "<td align=right style=""color: red"">" & FormatNumber(rs.Fields(2), 2) & "</td>" & _
							 "<td align=right style=""color: blue"">" & FormatNumber(rs.Fields(3), 4) & "</td>" & _
							 "<td align=right style=""color: red"">" & FormatNumber(rs.Fields(4), 4) & "</td>" & _
							 "<td align=right style=""color: red"">" & FormatNumber(rs.Fields(5), 7) & "</td>" & _
							 "</tr>"
			   
			   rs.MoveNext
			Loop
			rs.Close		   
			Response.Write "</table>"
			Set rs = Nothing			   
	End Sub

	Function EjecutarSQLCliente(ByVal Conexion, ByVal SQL)
	   Dim rsESQL
	   Const adUseClient = 2
	   Const adOpenStatic = 3
	   Const adLockBatchOptimistic = 4

		On Error Resume Next
	   Set EjecutarSQLCliente = Nothing
	   Set rsESQL = CreateObject("ADODB.Recordset")
	   rsESQL.CursorLocation = 3
	   rsESQL.Open SQL, Conexion, 3, 4
		If Err.number <> 0 Then
			Set rsESQL = Nothing
		End If
		'If rsESQL Is Nothing Then Exit Function
	   Set rsESQL.ActiveConnection = Nothing
	   Set EjecutarSQLCliente = rsESQL
	   Set rsESQL = Nothing
   
	End Function	

%>

<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<title>MonedasAP</title>
<link rel="stylesheet" type="text/css" href="../Compartido/VisorAP.css">
</head>

<script LANGUAGE="VBScript">
	On Error Resume Next
		
	Sub window_onload()
		window.setInterval "CargarVisor",  10000, "vbscript"
	End Sub

	Sub CargarVisor
		window.location.reload 
	End Sub				      
</script>
<body border="0" style="margin: 2 2 2 2" sbackground="../images/clouds.jpg">
<form id="frmVisorAP" method="post">
<%	CargarVisor %>
</form>
</body>
</html>