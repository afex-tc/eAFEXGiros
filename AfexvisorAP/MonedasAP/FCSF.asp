<%@ Language=VBScript %>
<%

	Session("cnxVisorAP") = "Provider=SQLOLEDB.1;Password=38Ki40p67;User ID=cambios;Initial Catalog=cambios;Data Source=canelo;"
	
	
	'Session("cnxVisorAP") = "Provider=SQLOLEDB.1;Password=cambios;User ID=cambios;Initial Catalog=cambios;Data Source=alerce;"

	CargarVisor	

	Sub CargarVisor
		On Error Resume Next		
			Dim rs, sSQL, sHTML
			   
			On Error Resume Next

			sSQL = "SELECT    codigo_moneda AS Moneda, alias_moneda AS Alias, pm_compra AS Compra, pm_venta AS Venta, paridad_compra AS 'P.Compra', paridad_venta AS 'P.Venta', paridad_transferencia " & _
			       "FROM      VAPPrecio_referencia WITH(nolock) " & _
			       "WHERE     estado_ap = 1 AND estado = 1 AND producto = 1 " & _
			       "ORDER BY  orden_ap"
			Set rs = EjecutarSQLCliente(Session("cnxVisorAP"), sSQL)			   
			If Err.Number <> 0 Then 
				Set rs = Nothing
				Exit Sub
			End If	   
			Response.Write "MN:Monedas:0,000:0,000:0,000:0,000:0,000000"
						
			Do Until rs.EOF
				Response.Write _
							 rs.Fields(0) & ":" & _
							 rs.Fields(1) & ":" & _
							 FormatNumber(rs.Fields(2), 2) & ":" & _
							 FormatNumber(rs.Fields(3), 2) & ":" & _
							 FormatNumber(rs.Fields(4), 4) & ":" & _
							 FormatNumber(rs.Fields(5), 4) & ":" & _
							 FormatNumber(rs.Fields(6), 7) & vbCrlf
			   
			   rs.MoveNext
			Loop
			rs.Close		   
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

