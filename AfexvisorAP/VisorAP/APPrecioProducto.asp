<%@ Language=VBScript %>
<%
'	Response.Expires = 0
'	Response.buffer = true	
'	If Session("CodigoCliente") = "" Then
'		response.Redirect "../Compartido/TimeOut.htm"
'		response.end
'	End If
%>
<%
	
%>
<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<title>Precios <%=Request("cmn")%></title>
<link rel="stylesheet" type="text/css" href="../Compartido/VisorAP.css">
</head>
<!--#INCLUDE virtual="/afexvisorap/visorap/Rutinas.htm" -->
<script LANGUAGE="VBScript">
<!--
	
	Sub window_onload()
		Cargar
	End Sub
	
	Sub Cargar()
	   Dim rs, sSQL, sHTML
			   
	   On Error Resume Next

		sSQL = "SELECT nombre_producto AS Producto, pr_Compra AS Compra, pr_Venta AS Venta " & _
		       "FROM   VAPPrecio_Referencia PR " & _
		       "WHERE  codigo = " & EvaluarStr("<%=Request("mn")%>") & " AND estado = 1 AND pr.visor = 1 " & _
		       "ORDER BY producto "
		Set rs = EjecutarSQLCliente("<%=Session("cnxVisorAP")%>", sSQL)			   
		If Err.Number <> 0 Then 
			Set rs = Nothing
			Exit Sub
		End If	   

		sHTML = "<table id=tbPrecioProducto border=0 cellspacing=1 cellpadding=1 style=""display: ; border: 1px solid '#EEEEEE'"">"
		sHTML =  sHTML & "<tr class=""titulo"" align=center class=titulo style=""height: 26px"">" & _
					 "<td width=180px>" & rs.Fields(0).name & "</td>" & _
					 "<td width=80px>" & rs.Fields(1).name & "</td>" & _
					 "<td width=80px>" & rs.Fields(2).name & "</td>" & _
					 "</tr>"			      
	   Do Until rs.EOF
			sHTML = sHTML & _
						 "<tr style=""height: 22px; cursor: hand; "" color=#BBBBBB bgcolor=#F6F6F6 onMouseOver=""javascript:this.bgColor='#EEEEEE'"" onMouseOut=""javascript:this.bgColor='#F6F6F6'"" >" & _
						 "<td align=left >" & rs.Fields(0) & "</td>" & _
						 "<td align=right style=""color: blue"">" & FormatNumber(rs.Fields(1), 2) & "</td>" & _
						 "<td align=right style=""color: red"">" & FormatNumber(rs.Fields(2), 2) & "</td></tr>"
			   
	      rs.MoveNext
	   Loop
	   rs.Close		   
		sHTML = sHTML & "</table>"
		dvMonedas.innerHTML = sHTML
	   Set rs = Nothing
			   
	End Sub


//-->
</script>
<body style="margin-left: 10px; margin-top: 10px">
<div id="dvMonedas" style>
	espere...
</div>
</body>
</html>
