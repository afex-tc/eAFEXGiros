<!-- Rutinas.htm -->
<%
'************* Funciones y Procedimientos *******************
		
	Function ObtenerVisorAP(Byval Conexion, ByVal Moneda)
	   Dim rs
		Dim sSQL

	   Set ObtenerVisorAP = Nothing

	   On Error Resume Next

		sSQL = "APObtenerPRActualPro '" & (Moneda) & "' "
	   	   
	   Set rs = EjecutarSQLCliente(Conexion, sSQL)

	   If Err.Number <> 0 Then 
			Set rs = Nothing
			'Msgbox err.Description 			
		End If
	   
	   Set ObtenerVisorAP = rs

	   Set rs = Nothing
	End Function


	Function FormatoFechaSQL(Byval Fecha)
		FormatoFechaSQL = Year(Fecha) & Right("0" & Month(Fecha), 2) & Right("0" & Day(Fecha), 2)
	End Function


	Function ObtenerRSAPVPosicion(ByVal Conexion, ByVal Moneda, ByVal Fecha)
	   Dim rs
	   Dim sSQL
	      
	   On Error Resume Next
	   Set ObtenerRSAPVPosicion = Nothing
	   
	   sSQL = "APVisorPosicion '" & Moneda & "', '" & FormatoFechaSQL(Fecha) & "' "
	   Set rs = EjecutarSQLCliente(Conexion, sSQL)
	   If Err.Number <> 0 Then 
			Set rs = Nothing
			'Msgbox " ..." & err.Description 			
		End If	   
		
	   Set ObtenerRSAPVPosicion = rs	   
	   Set rs = Nothing
	End Function


	Public Function DiferenciaHora(ByVal Inicio, ByVal Termino)
		If Inicio = Empty Then
			DiferenciaHora = "00:00:00"
			Exit Function
		End If
		DiferenciaHora = TimeSerial(CInt(Left(Termino, 2)) - CInt(Left(Inicio, 2)), _
                               CInt(Mid(Termino, 4, 2)) - CInt(Mid(Inicio, 4, 2)), _
                               CInt(Right(Termino, 2)) - CInt(Right(Inicio, 2)))
	End Function


	Public Function EjecutarSQLCliente(ByVal Conexion, ByVal SQL)
	   Dim rsESQL
	   Const adUseClient = 2
	   Const adOpenStatic = 3
	   Const adLockBatchOptimistic = 4

	   Set EjecutarSQLCliente = Nothing
	   Set rsESQL = CreateObject("ADODB.Recordset")
	   rsESQL.CursorLocation = 3
	   rsESQL.Open SQL, Conexion, 3, 4
		If Err.number <> 0 Then
			Set rsESQL = Nothing
			MsgBox Err.Description 
		End If
		'If rsESQL Is Nothing Then Exit Function
	   Set rsESQL.ActiveConnection = Nothing
	   Set EjecutarSQLCliente = rsESQL
	   Set rsESQL = Nothing
   
	End Function	

	Function EvaluarStr(ByVal Valor)
		Dim Devuelve
		
		If Valor="" Then 
			EvaluarStr = "Null"	
		Else
			EvaluarStr = "'" & Valor & "'"
		End If

	End Function

%>
