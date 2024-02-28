<!-- Rutinas.asp -->
<%
	
	Function FormatoFechaSQL(Byval Fecha)
		FormatoFechaSQL = "'" & Year(Fecha) & Right("0" & Month(Fecha), 2) & Right("0" & Day(Fecha), 2) & "'"
	End Function

	Function FormatoNumeroSQL(ByVal Numero)
	   Dim sFormato
	   
	   sFormato = CStr(Numero)
	   sFormato = Replace(Numero, ",", ".")
	   FormatoNumeroSQL = sFormato
	End Function

	Function EjecutarSQLCliente(ByVal Conexion, ByVal SQL)
	   Dim rsESQL, cnx
	   
	   Const adUseClient = 2
	   Const adOpenStatic = 3
	   Const adLockBatchOptimistic = 4

		On Error Resume Next
	   Set EjecutarSQLCliente = Nothing
      Set cnx = CreateObject("ADODB.Connection")
	   'cnx.CommandTimeout = TimeOut
	   cnx.Open Conexion
		If Err.number <> 0 Then
			Set cnx = Nothing
			'MsgBox Err.Description 
		End If

	   Set rsESQL = CreateObject("ADODB.Recordset")
	   rsESQL.CursorLocation = 3
	   rsESQL.Open SQL, cnx, 3, 4
		If Err.number <> 0 Then
			Set cnx = Nothing
			Set rsESQL = Nothing
			'MsgBox Err.Description 
		End If
		'If rsESQL Is Nothing Then Exit Function

	   Set rsESQL.ActiveConnection = Nothing
	   Set EjecutarSQLCliente = rsESQL
	   Set rsESQL = Nothing
		Set cnx = Nothing
   
	End Function	

	Function EvaluarVar(ByVal Valor, ByVal Devuelve)
		
		If Valor="" Then 
			EvaluarVar = Devuelve
		Else
			EvaluarVar = Valor
		End If

	End Function

	
	Function EvaluarStr(ByVal Valor)
		
		If Valor="" Then 
			EvaluarStr = "Null"	
		Else
			EvaluarStr = "'" & Valor & "'"
		End If

	End Function


%>
