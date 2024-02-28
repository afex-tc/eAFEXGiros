<!-- Giros.asp -->
<%

Function ObtenerAgenteLista(ByVal Conexion, _
									   ByVal CodigoPais, _
										ByVal CodigoCiudad, _
										ByVal CodigoRegion, _
										ByVal Categoria)
   Dim sSQL
	
   'Manejo de errores
   On Error Resume Next
	
   'Crea la consulta
   sSQL = "SELECT    * " & _
          "FROM      vAgente " & _
          "WHERE     mostrar_web = 1 "
	
   If CodigoPais <> "" Then
      sSQL = sSQL & _
             " and pais_agente = '" & CodigoPais & "' "
   End If
   If CodigoCiudad <> "" Then
      sSQL = sSQL & _
             " and ciudad_agente = '" & CodigoCiudad & "' "
   End If
   If CodigoRegion <> 0 Then
      sSQL = sSQL & _
             " and codigo_region = " & CodigoRegion & " "
   End If
   If Categoria > -1 Then
      If Categoria = Sucursales Then
         sSQL = sSQL & _
                " and categoria_agente in (0, 1, 2) "
      End If
      If Categoria = Agentes Then
         sSQL = sSQL & _
                " and categoria_agente in (3, 4) "
      End If
   End If
   sSQL= sSQL & "order by nombre_agente"
   '"http:../compartido/informacion.asp?detalle=Temporalmente fuera de servicio.<br><br>" & sSQL
   'MostrarErrorMS sSQL
   'Asigna al metodo el resultado de la consulta
   Set ObtenerAgenteLista = EjecutarSQLCliente(Conexion, sSQL)
	
   'Si se produjeron errores en la consulta
   If Err.Number <> 0 Then
		Set ObtenerAgenteLista = Nothing
		MostrarErrorMS "Direcciones en Chile"
	End If
	
End Function


Function ValidarInvoice(ByVal Conexion, _
                        ByVal Invoice, _
                        ByVal Agente)
   Dim sSQL
   Dim rsGiro
	
   ValidarInvoice = False
	   
   'Manejo de errores
   On Error Resume Next
	
   'Crea la consulta
   sSQL = "SELECT    COUNT(codigo_giro) AS Giros " & _
          "FROM      Giro WITH(nolock) " & _
          "WHERE     invoice = '" & Invoice & "' " & _
          "And       agente_captador = '" & Agente & "' "
	
   'Set rsGiro = New ADODB.Recordset
   Set rsGiro = EjecutarSQLCliente(Conexion, sSQL)
	
   'Si se produjeron errores en la consulta
   If Err.Number <> 0 Then
		Set rsGiro = Nothing
		Exit Function
   End If
   If rsGiro Is Nothing Then
		Exit Function
   End If
	
   If rsGiro("Giros") > 0 Then
      ValidarInvoice = True
   End If
	
   Set rsGiro = Nothing
	
End Function

Function ObtenerGiroXP(ByVal Codigo)
	Dim rsO
	
	On Error Resume Next
	Set rsO = BuscarGiroXP(Session("afxCnxAFEXpress"), Codigo)
	'mostrarerrorms rso("codigo_giro")
	If Err.number <> 0 Then
		Set rsO = Nothing
		MostrarErrorMS "Obtener Giro 1"
	End If
	
	Set ObtenerGiroXP = rsO
	Set rsO = Nothing
End Function

Function BuscarGiroXP(ByVal Conexion, _
							  ByVal Codigo)
   Dim sSQL
	
   'Manejo de errores
  On Error Resume Next
	
   'Crea la consulta
   ' JFMG 31-05-2011 se comenta para que utilice un procedimiento almacenado
   'sSQL = "SELECT    * " & _
   '       "FROM      VGiro " & _
   '       "WHERE     codigo_giro = '" & Trim(Codigo) & "' "
   sSQL = " exec MostrarDatosGiroCodigo '" & Trim(Codigo) & "'"
   ' FIN JFMG 31-05-2011
		
   'Asigna al metodo el resultado de la consulta
   Set BuscarGiroXP = EjecutarSQLCliente(Conexion, sSQL)
	
   'Si se produjeron errores en la consulta
   If Err.Number <> 0 Then
		MostrarErrorMS "Buscar Giro " 
   End If
	
End Function

Function ValidarCredito(ByVal Agente, ByVal Moneda, ByVal Monto)
   Dim sSQL
   Dim rsCrdt, nDec
   dim Saldo ,Total, saldo1, credito, TotalUsado
	
   ValidarCredito = False
	
   'Manejo de errores
   On Error Resume Next
	
	
	sSQL = "ValidarCreditoAgente '" & Trim(Agente)	& "' , '" & trim(Moneda) & "'" 
	
	If Moneda = Session("MonedaNacional") Then
		'Crea la consulta
		'sSQL = "SELECT    ISNULL(ag.credito_nacional, 0) AS Credito, ISNULL(pc.saldo_nacional, 0) AS Saldo, mn.prefijo_moneda, mn.cantidad_decimales " & _
		'       "FROM      Agente AG " & _
		'       "JOIN		Plan_Cuentas PC ON pc.codigo_agente=ag.codigo_agente " & _
		'       "JOIN		Moneda MN ON mn.codigo_moneda='" & TRIM(Moneda) & "' " & _
		'       "WHERE     ag.codigo_agente='" & TRIM(Agente) & "' " & _
		'       "				AND pc.uso_cuenta IN (2, 3, 4) " & _
		'       "				AND pc.codigo_moneda='" & TRIM(Moneda) & "' "	
		nDec = 0
	Else
		'Crea la consulta
		'sSQL = "SELECT    ISNULL(ag.credito, 0) AS Credito, ISNULL(pc.saldo_extranjera, 0) AS Saldo, mn.prefijo_moneda, mn.cantidad_decimales " & _
		'       "FROM      Agente AG " & _
		'       "JOIN		Plan_Cuentas PC ON pc.codigo_agente=ag.codigo_agente " & _
		'       "JOIN		Moneda MN ON mn.codigo_moneda='" & TRIM(Moneda) & "' " & _
		'       "WHERE     ag.codigo_agente='" & TRIM(Agente) & "' " & _
		'       "				AND pc.uso_cuenta IN (2, 3, 4) " & _
		'       "				AND pc.codigo_moneda='" & TRIM(Moneda) & "' "
		nDec = 2
	End If
	'Response.Write ssql
	'Response.End 

   'Set rsGiro = New ADODB.Recordset
   Set rsCrdt = EjecutarSQLCliente(Session("afxCnxAFEXpress"), sSQL)
   'Si se produjeron errores en la consulta
   If Err.Number <> 0 Then
		Set rsCrdt = Nothing
		Exit Function
   End If
   
   If rsCrdt Is Nothing Then
		Err.Raise 100001, "Validar Credito", "No se pudo ejecutar la consulta de credito", 0, 0
	   Set rsCrdt = Nothing
		MostrarErrorMS "Validar Credito 1"
	End If
	
   If rsCrdt.EOF Then 
		Err.Raise 100001, "Validar Credito", "No se encontro el agente y sus saldos", 0, 0
	   Set rsCrdt = Nothing
		MostrarErrorMS "Validar Credito 2"
	End If
	
	'**NUEVO***************PSS 23-12-2008****************************************

	If Ccur(0 & rsCrdt("credito")) = cCur(0) Then
		ValidarCredito = true
	ElseIf (  cCur(0 & Monto)+ cCur(rsCrdt("saldo1")) ) > cCur(0 & rsCrdt("Credito"))+ cCur(rsCrdt("saldo")) Then
		Err.Raise 1001, "Validar Credito", "Su credito disponible es de " & rsCrdt("prefijo_moneda") & " " & formatNumber(ccur(0 & rsCrdt("Credito")) + formatNumber(cCur(rsCrdt("saldo")) - cCur(rsCrdt("saldo1"))) , rsCrdt("cantidad_decimales")) & ". No se puede enviar el giro, no tiene saldo a favor.", 0, 0 
		Set rsCrdt = Nothing
		MostrarErrorMS "Validar Credito 3"
	Else
		saldo = cCur(0 & rsCrdt("saldo"))
		total =cCur(rsCrdt("Credito"))+ cCur(Saldo)
		Credito = cCur(0 & (rsCrdt("Credito")))
		If (ccur(rsCrdt("Saldo"))+ cCur(rsCrdt("Credito")))< 0 then
			Err.Raise 1001, "Validar Credito", "Su credito disponible es de " & rsCrdt("prefijo_moneda") & " " & formatNumber(ccur(0 & rsCrdt("Credito")) + formatNumber(cCur(rsCrdt("saldo")) - cCur(rsCrdt("saldo1"))) , rsCrdt("cantidad_decimales")) & ". No se puede enviar el giro, no tiene saldo a favor.", 0, 0 
			Set rsCrdt = Nothing
			MostrarErrorMS "Validar Credito 4"
		Else
			If ccur(rsCrdt("Saldo"))+ CCUR(CREDITO) - ccur(rsCrdt("Saldo1")) < 0   then
				TotalUsado = ccur(rsCrdt("Saldo"))+CCUR(CREDITO)- ccur(rsCrdt("Saldo1"))
			End If
		End If
	End If	

	IF cCur(TotalUsado) > cCur(Credito) Then
			Err.Raise 1001, "Validar Credito", "Su credito disponible es de " & rsCrdt("prefijo_moneda") & " " & formatNumber(ccur(0 & rsCrdt("Credito")) + formatNumber(cCur(rsCrdt("saldo")) - cCur(rsCrdt("saldo1"))) , rsCrdt("cantidad_decimales")) & ". No se puede enviar el giro, no tiene saldo a favor.", 0, 0 
			Set rsCrdt = Nothing
			MostrarErrorMS "Validar Credito 5"
	End If
	ValidarCredito = True
		
'******************************************************************************************	

	
	'If cCur(0 & rsCrdt("credito")) = cCur(0) Then
    '  ValidarCredito = True
  ' ElseIf (cCur(rsCrdt("saldo")) + cCur(0 & Monto)) > cCur(0 & rsCrdt("Credito")) Then
	'Else
'		If cCur(rsCrdt("saldo")) < 0 Then
'			saldo = cCur(rsCrdt("saldo"))* - 1
'			total =cCur(rsCrdt("Credito"))- cCur(Saldo)
'			If cCur(total) < 0  Then 
 ' 				Err.Raise 1001, "Validar Credito", "Su credito disponible es de " & rsCrdt("prefijo_moneda") & " " & formatNumber(ccur(Total) , rsCrdt("cantidad_decimales")) & ". No se puede enviar el giro, no tiene saldo a favor.", 0, 0 
'				Set rsCrdt = Nothing
'				MostrarErrorMS "Validar Credito 3"
'			End If
'			ValidarCredito = True
'		
'		' JFMG 05-12-2008 valida que el crédito más el saldo sean suficientes para el monto del giro
'		else
'			saldo = cCur(rsCrdt("saldo"))
'			total = cCur(rsCrdt("Credito")) + cCur(Saldo)
'			If cCur(total) < ccur(Monto) Then 
 ' 				Err.Raise 1001, "Validar Credito", "Su credito disponible es de " & rsCrdt("prefijo_moneda") & " " & formatNumber(ccur(Total) , rsCrdt("cantidad_decimales")) & ". No se puede enviar el giro, no tiene saldo a favor.", 0, 0 
'				Set rsCrdt = Nothing
'				MostrarErrorMS "Validar Credito 4"
'			End If
'			ValidarCredito = True
'		' *************************** FIN ***************************
		
'		End IF
'	End If
		
   Set rsCrdt = Nothing
	
End Function


%>
