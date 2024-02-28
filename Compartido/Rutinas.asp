<!-- Rutinas.asp -->
<%
'***************************************FUNCIONES VARIAS********************************************
	'Objetivo:		verificar si la tecla que se presionó corresponde a un número o letra y dejarla 
	'				seguir o eliminar la acción según el llamado del procedimiento
	'Parámetros:	Tipo, indica que quiere verificar "1=Número; 2=Letra; 3=Número y Letra"
	Public Sub IngresarTexto(byval Tipo, ByVal Tecla)
		' verifica el tipo de dato que se quiere verificar
		
		Select Case Tipo
			Case 1		' verificar si es un número o coma
				If Tecla=Asc(".") Then
					IngresarTexto=  Asc(",")
					Exit Sub
				End If
				If (Tecla < 48 OR Tecla > 57) And _
					Tecla <> Asc(",") Then
					IngresarTexto=  0
				End If

			Case 2		' verificar si es una letra
				If (Tecla < 65 OR Tecla > 90) AND _
				   (Tecla < 97 OR Tecla > 122) AND _
				   Tecla <> 46 AND Tecla <> 39 AND _
				   Tecla <> 180 AND Tecla <> 209 AND _
				   Tecla <> 241 AND Tecla <> 32 Then
					' no es una letra, se cancela la acción
					IngresarTexto=  0
				ElseIf (Tecla >= 97 AND Tecla <= 122 ) Then
					' es una minuscula, se transforma a mayúscula
					'IngresarTexto=  (Tecla - 32)
				ElseIf Tecla = 241 Then
				    '
					IngresarTexto=  209
				End If
			
			Case 3		' verificar si es una letra o número
				If (Tecla < 65 OR Tecla > 90  ) AND _
				   (Tecla < 97 OR Tecla > 122 ) AND _
				   Tecla <> 46 AND Tecla <> 39 AND _
				   Tecla <> 180 AND Tecla <> 209 AND _
				   Tecla <> 241 AND Tecla <> 32 AND _
				   (Tecla < 48 OR Tecla > 57  ) AND _
				   Tecla <> 35 Then
					IngresarTexto=  0
				ElseIf ( Tecla >= 97 AND Tecla <= 122 ) Then
					'IngresarTexto=  ( Tecla - 32 )
				ElseIf Tecla = 241 Then
					IngresarTexto=  209
				End If					
		End Select
	End Sub

	'Objetivo:  Devuelve una cadena de caracteres transformados a mayuscula y minuscula
	Public Function MayMin(ByVal Cadena)
	   Dim aCadena
	   Dim i 
	   Dim sCadena 
	   Dim STRING_MAYMIN 
	   
		MayMin = UCase(Cadena)
		Exit Function	   
	   If IsNull(Cadena) Then Exit Function
	   If Len(Trim(Cadena)) <= 0 Then Exit Function
	   STRING_MAYMIN = ";es;una;uno;de;por;para;que;si;y;"
	   aCadena = Split(Cadena, " ")
	   sCadena = ""
	   For i = 0 To UBound(aCadena)
	      If InStr(UCase(STRING_MAYMIN), ";" & UCase(aCadena(i)) & ";") = 0 Or i = 0 Then
	         sCadena = sCadena & PMayuscula(aCadena(i)) & " "
	      Else
	         sCadena = sCadena & LCase(aCadena(i)) & " "
	      End If
	   Next
	   sCadena = Left(sCadena, Len(sCadena) - 1)
	   MayMin = sCadena
	   
	End Function
	
	Public Function PMayuscula(ByVal Cadena)
	   
	   PMayuscula = UCase(Left(Cadena, 1)) & LCase(Mid(Cadena, 2))
	   
	End Function
	

	Public Function LimpiarCampoTxt(ByVal Cadena)
	        
	        Cadena = Replace(LCase(Cadena),"select","")
	        Cadena = Replace(LCase(Cadena),"drop","")
	        Cadena = Replace(LCase(Cadena),"delete","")
	        Cadena = Replace(LCase(Cadena),"update","")
	        Cadena = Replace(LCase(Cadena),"exec","")
	        Cadena = Replace(LCase(Cadena),"having","")
	        Cadena = Replace(LCase(Cadena),"truncate","")
	        Cadena = Replace(LCase(Cadena),"alter","")
	        Cadena = Replace(LCase(Cadena),"create","")
	        Cadena = Replace(LCase(Cadena),"grant","")
	        Cadena = Replace(LCase(Cadena),"insert","")
	        Cadena = Replace(LCase(Cadena),"<","")
	        Cadena = Replace(LCase(Cadena),">","")
	        Cadena = Replace(LCase(Cadena),"'","")
	        Cadena = Replace(LCase(Cadena),"--","")	   
	        
	        LimpiarCampoTxt = Cadena
	End Function

	'Objetivo:  Devuelve una cadena de caracteres transformados a mayuscula y minuscula
	Public Function MayusculaMinuscula(ByVal Cadena)
	   Dim aCadena
	   Dim i 
	   Dim sCadena 
	   Dim STRING_MAYMIN 
	   
	   If IsNull(Cadena) Then Exit Function
	   If Len(Trim(Cadena)) <= 0 Then Exit Function
	   STRING_MAYMIN = ";es;una;uno;de;por;para;que;si;y;"
	   aCadena = Split(Cadena, " ")
	   sCadena = ""
	   For i = 0 To UBound(aCadena)
	      If InStr(UCase(STRING_MAYMIN), ";" & UCase(aCadena(i)) & ";") = 0 Or i = 0 Then
	         sCadena = sCadena & PMayuscula(aCadena(i)) & " "
	      Else
	         sCadena = sCadena & LCase(aCadena(i)) & " "
	      End If
	   Next
	   sCadena = Left(sCadena, Len(sCadena) - 1)
	   MayusculaMinuscula = sCadena
	   
	End Function

    'AMP 23/05/2016
    'Objetivo   : Obtener un listado de comunas segun codigo de ciudad SP: [dbo.ObtenerComunasCiudad @CodigoCiudad]
    'JIRA : CUM-590
    Sub CargarComunaCiudad(ByVal CodigoCiudad, ByVal Seleccionado)
    
        Dim sConexion
        Dim rs
        Dim sMayMin
        Dim sSelect
        On Error Resume Next

        sConexion = Session("afxCnxCorporativa")
        sSQL = "EXECUTE dbo.ObtenerComunasCiudad @CodigoCiudad = '" & CodigoCiudad & "'"
        Set rs = EjecutarSQLCliente(sConexion, sSQL)
        If Err.number <> 0 Then 
			MostrarErrorMS "Cargar CargarComunaCiudad"
		End If
		
		Response.Write "<option value=></option>"
		If  Not rs.EOF Then
			Do Until rs.eof	
				If UCASE(Trim(rs("codigo"))) = Ucase(Trim(Seleccionado)) Then 					
					sSelect = "SELECTED"				
				Else
					sSelect = ""
				End If
				sMayMin = trim(rs("nombre"))
				sMayMin = MayMin(sMayMin) 
				Response.write "<option " & sSelect & " value=" & _
				trim(rs("codigo")) & ">" & _
				sMayMin & _
				" </option> "
				If err.number <> 0 Then 
					MostrarErrorMS "Cargar CargarComunaCiudad"
				End If																
				rs.MoveNext
			Loop
		End If
    End Sub

    'AMP 23/05/2016
    'Objetivo   : Obtener un listado de Ciudades segun codigo Pais SP: [dbo.ObtenerCiudades @CodPais]
    'JIRA : CUM-590
    Sub CargarCiudadesPais(ByVal CodigoPais, ByVal Seleccionado)
    
        Dim sConexion
        Dim rs
        Dim sMayMin
        Dim sSelect
        On Error Resume Next

        sConexion = Session("afxCnxCorporativa")
        sSQL = "EXECUTE dbo.ObtenerCiudades @CodPais = '" & CodigoPais & "'"
        Set rs = EjecutarSQLCliente(sConexion, sSQL)
        If Err.number <> 0 Then 
			MostrarErrorMS "Cargar CargarCiudadesPais"
		End If
		
		Response.Write "<option value=></option>"
		If  Not rs.EOF Then
			Do Until rs.eof	
				If UCASE(Trim(rs("codigo"))) = Ucase(Trim(Seleccionado)) Then 					
					sSelect = "SELECTED"				
				Else
					sSelect = ""
				End If
				sMayMin = trim(rs("nombre"))
				sMayMin = MayMin(sMayMin) 
				Response.write "<option " & sSelect & " value=" & _
				trim(rs("codigo")) & ">" & _
				sMayMin & _
				" </option> "
				If err.number <> 0 Then 
					MostrarErrorMS "Cargar CargarCiudadesPais"
				End If																
				rs.MoveNext
			Loop
		End If
    End Sub

	Sub CargarUbicacion(ByVal Tipo, ByVal Codigo, ByRef Seleccionado)
		Dim sSelect
		Dim afxCOM
		Dim rs
		Dim sConexion
		Dim sMayMin
		Dim sCodigo
		
		On Error Resume Next
		Set rs = Server.CreateObject("ADODB.Recordset")
		If Err.number <> 0 Then response.Redirect "Error.asp?Titulo=Error en HágaseCliente&Number=" & Err.number   & "&Source=" & Err.Source & "&Description=" & Err.description		
		sConexion = Session("afxCnxCorporativa") '"DSN=AFEXcorporativa;UID=corporativa;PWD=afxsqlcor;"
		sSelect = Empty

		Select Case Tipo
			Case 1	'Pais
				
				Set afxCOM = Server.CreateObject("AfexCorporativo.Pais")				
				Set rs = afxCOM.Buscar(sConexion)
				sCodigo = Seleccionado
				
			Case 2	'Ciudad
				Set afxCOM = Server.CreateObject("AfexCorporativo.Ciudad")				
				Set rs = afxCOM.Buscar(sConexion, , Codigo)
				sCodigo = Seleccionado

            Case 3	'Comuna
				Set afxCOM = Server.CreateObject("AfexCorporativo.Comuna")
				Set rs = afxCOM.Buscar(sConexion, , Codigo)
				sCodigo = Seleccionado

		End Select		
	
		If Err.number <> 0 Then 
			'response.Redirect "../Compartido/Error.asp?Titulo=Errores en HágaseCliente&Number=" & Err.number   & "&Source=" & Err.Source & "&Description=" & Err.description
			MostrarErrorMS "Cargar Ubicación 1"
		End If
				
		If afxCOM.ErrNumber <> 0 Then
			If afxCOM.ErrNumber = 10527 Or afxCOM.ErrNumber = 10532 Then	
				Set afxCOM = Nothing	
				Exit Sub
			Else
				'response.Redirect "../Compartido/Error.asp?Titulo=Error en Cargar Ubicacion&Number=" & afxCOM.ErrNumber  & "&Source=" & afxCOM.ErrSource & "&Description=" & replace(afxCOM.ErrDescription, vbCrLf , "^")	
				MostrarErrorAFEX afxCOM, "Cargar Ubicación 2 "  & codigo
			End If
		End If		
		Response.Write "<option value=></option>"
		If  Not rs.EOF Then
			Do Until rs.eof	
				If UCASE(Trim(rs("codigo"))) = Ucase(Trim(sCodigo)) Then 					
					sSelect = "SELECTED"				
				Else
					sSelect = ""
				End If
				sMayMin = trim(rs("nombre"))
				sMayMin = MayMin(sMayMin) 
				Response.write "<option " & sSelect & " value=" & _
				trim(rs("codigo")) & ">" & _
				sMayMin & _
				" </option> "
				If err.number <> 0 Then 
					'response.Redirect "../Compartido/Error.asp?Titulo=Error en HágaseCliente&Number=" & Err.Number  & "&Source=" & Err.source  & "&Description=" & Err.description	
					MostrarErrorMS "Cargar Unicación 3"
				End If																
				rs.MoveNext
			Loop
		End If
			
		Set afxCOM = Nothing	
	End Sub

	' Objetivo   : Realizar el formateo de los numeros para colocarlos en una consulta sql
	' Parametros : Numero    Número a formatear
	' Devuelve   : El número formateado como string con decimales para SQL
	Function FormatoNumeroSQL(ByVal Numero)
	   Dim sFormato
	   
	   sFormato = CStr(Numero)
	   sFormato = Replace(Numero, ",", ".")
	   FormatoNumeroSQL = sFormato

	End Function

    
	
	
	
	Function ObtenerMenuCliente(ByVal Codigo, ByRef Nombre)
		Dim afxCliente, rs, sMenu
		
		Set afxCliente = Server.CreateObject("AfexCorporativo.Cliente")
		Set rs = Server.CreateObject("ADODB.recordset")
		'sCodigo = Request("Codigo")
		
		On Error Resume Next
		Set rs = afxCliente.Buscar(Session("afxCnxCorporativa"), 3, Codigo)
		If afxCliente.ErrNumber <> 0 Then			
			Set rs = Nothing
			MostrarErrorAFEX afxCliente, ""
		End If
		On Error Goto 0
		
		sMenu = "10;20;30;31;"
		If rs("envio_giro")=1 Then sMenu = sMenu & "11;21;"
		If rs("recepcion_giro")=1 Then sMenu = sMenu & "12;"
		If rs("giro_pendiente_aviso")=1 Then sMenu = sMenu & "15;"
		If rs("envio_transferencia")=1 Then sMenu = sMenu & "13;22;"
		If rs("compra_venta")=1 Then sMenu = sMenu & "14;23;"
		If rs("paridades")=1 Then sMenu = sMenu & "16;"
		If rs("agregar_cliente")=1 Then sMenu = sMenu & "24;"
		If rs("informes_cliente")=1 Then sMenu = sMenu & "17;"
		If rs("giro_enviado_pendiente")=1 Then sMenu = sMenu & "40;"
		'If rs("alarma")=1 Then sMenu = sMenu & "33;"
		Nombre = MayMin(rs("nombre"))
		Set afxCliente = Nothing
		Set rs = Nothing
		ObtenerMenuCliente = sMenu
				
	End Function

	Function ObtenerDDI(ByVal Tipo, ByVal Codigo)
		Dim afxCOM
		Dim ddi
		
		ObtenerDDI = 0
		If Trim(Codigo) = "" Then Exit Function
		
		Select Case Tipo
			Case 1
				Set afxCOM = Server.CreateObject("AfexCorporativo.Pais")
			
			Case 2
				Set afxCOM = Server.CreateObject("AfexCorporativo.Ciudad")

			Case 3
				Set afxCOM = Server.CreateObject("AfexCorporativo.Comuna")
		End Select		
		ddi = afxCOM.BuscarDDI(Session("afxCnxCorporativa"), Codigo)
		If afxCOM.ErrNumber <> 0 Then 
			Set afxCOM = Nothing
			Exit Function
		End If
		ObtenerDDI = ddi
		Set afxCOM = Nothing		
	End Function

	'Objetivo   :  Evaluar el contenido de un paràmetro
	'Paràmetro  :  Valor     Valor a evaluar
	'Devuelve   :  "Null"      Si Valor es Empty
	'              Valor     Si Valor es distinto de EMpty
	Function EvaluarVar(ByVal Valor, _
							  ByVal Devuelve)
		If Devuelve = Empty Then Devuelve = ""		
	   If IsNull(Valor) Then
	      EvaluarVar = Devuelve
	   Else
	      EvaluarVar = Valor
	   End If

	End Function

	'Objetivo   :  Da formato tipo Rut
	Function FormatoRut(ByVal Rut)

	   FormatoRut = Empty
	   
	   If Rut = Empty Or IsNull(Rut) Then
	      Exit Function
	   End If
	   Rut = Trim(Replace(Rut, "-", ""))
	   Rut = Replace(Rut, ".", "")
	   FormatoRut = FormatNumber(Left(Rut, Len(Rut) - 1), 0) & "-" & Right(Rut, 1)
	                              
	End Function

		
	Function BuscarCliente(ByVal Campo, Byval Argumento, ByVal Argumento2, ByVal Argumento3)
		Dim afxClienteATC, rs
		
		On Error Resume Next
		'Set afxClienteATC = Server.CreateObject("AfexCorporativo.Cliente")
		If Session("ModoPrueba") Then
			Set rs = ObtenerATC(Session("afxCnxCorporativa"), Campo, Argumento, Argumento2, Argumento3)
		Else
			'Set rs = afxClienteATC.ObtenerATC(Session("afxCnxCorporativa"), Campo, Argumento, Argumento2, Argumento3)
			Set rs = ObtenerATC(Session("afxCnxCorporativa"), Campo, Argumento, Argumento2, Argumento3)
		End If
		If Err.number <> 0 Then
			Set rs = Nothing
			'Set afxClienteATC = Nothing
			MostrarErrorMS "Buscar Cliente 1"
		End If
		'If afxClienteATC.ErrNumber <> 0 Then
		'	Set rs = Nothing
		'	MostrarErrorAFEX afxClienteATC, "Buscar Cliente 2"
		'End If
		
		Set BuscarCliente = rs
		Set rs = Nothing
		'Set afxClienteATC = Nothing
	End Function
	
	
	Function ObtenerATC(ByVal Conexion, _
                       ByVal Campo, _
                       ByVal Argumento1, _
                       ByVal Argumento2, _
                       ByVal Argumento3)
	   Dim rsATC
	   Dim sSQL
	   Dim Condicion
	   Dim Rut
	   Dim Raya
	   Dim sComa

	   Set ObtenerATC = Nothing

	   On Error Resume Next

	   sComa = ""
	   
	      ' verifica el campo por el que se busca
	   Select Case Campo
	      Case 1 'afxCampoBusqueda.afxRut
				sSQL = "EXECUTE ObtenerClienteRut "
				Argumento1 = replace(Argumento1, ".", "")
				Argumento1 = replace(Argumento1, "-", "")
				Argumento1 = right("00000" & Argumento1, 9)

				'Condicion = " Rut in ('" & Argumento1 & "', '" & Rut & Raya & "')"
				Condicion = " @Rut = '" & Argumento1 & "' "

	      Case 2 'afxCampoBusqueda.afxPasaporte
				sSQL = "EXECUTE ObtenerClientePasaporte "
				Condicion = " @pasaporte = '" & Argumento1 & "'"
	                  
	      'Case afxCampoBusqueda.afxCodigo
	      '   Condicion = " Codigo = " & Argumento1
	         
	      Case 4 'afxCampoBusqueda.afxNombres
				sSQL = "EXECUTE ObtenerClienteNew "
				Condicion = " @nombre = '" & Argumento1 & "'"
				Condicion = Condicion & ", @paterno = '" & Trim(Argumento2) & "'"
				Condicion = Condicion & ", @materno = '" & Trim(Argumento3) & "'"
	         
	      Case 5 'afxCampoBusqueda.afxCodigoExchange
				sSQL = "EXECUTE ObtenerClienteNew "
				Condicion = " @exchange = '" & Argumento1 & "'"

	      Case 6 'afxCampoBusqueda.afxCodigoExpress
				sSQL = "EXECUTE ObtenerClienteNew "
				Condicion = " @express = '" & Argumento1 & "'"
	      
	      Case 7 'afxCampoBusqueda.afxTelefono
				sSQL = "EXECUTE ObtenerClienteTelefono "
				Condicion = " @telefono = " & Argumento1

	      Case 8 'afxCampoBusqueda.afxTelefono2
				sSQL = "EXECUTE ObtenerClienteTelefono "
				Condicion = " @telefono = " & Argumento1
				
		  Case 9 ' tarjeta
				sSQL = "EXECUTE ObtenerClienteTarjeta "
				Condicion = evaluarstr(Argumento1)

	   End Select
	   sSQL = sSQL & Condicion
	   
	  
	   
	   Set rsATC = EjecutarSQLCliente(Conexion, sSQL)

	   If Err.Number <> 0 Then 
			Set rsATC = Nothing
			MostrarErrorMS "Obtener ATC"
		End If
	   
	   Set rsATC.ActiveConnection = Nothing
	   Set ObtenerATC = rsATC

	   Set rsATC = Nothing
		'MostrarErrorMS "ok"
	End Function

	

	Function ObtenerCliente(ByVal Campo, Byval Argumento)
		Dim afxCliente2, rs
		
		On Error Resume Next
		Set afxCliente2 = Server.CreateObject("AfexCorporativo.Cliente")
		Set rs = afxCliente2.Buscar(Session("afxCnxCorporativa"), Campo, Argumento)
		If Err.number <> 0 Then
			Set rs = Nothing
			Set afxCliente2 = Nothing
			MostrarErrorMS ""
		End If
		If afxCliente2.ErrNumber <> 0 Then			
			Set rs = Nothing
			MostrarErrorAFEX afxCliente2, ""
		End If
		
		Set ObtenerCliente = rs
		Set rs = Nothing
		Set afxCliente2 = Nothing
	End Function
	
	'Objetivo:	Cargar en un combo los agentes pagadores para giros
	Function CargarAgentePagador(ByVal Pais, ByVal Ciudad, ByVal Seleccionado, ByVal Mail)
		Dim afxGiro, rs
		
		On Error Resume Next
		Set afxGiro = Server.CreateObject("AFEXGiro.Giro")
		Set rs = afxGiro.ObtenerAgentePagador(Session("afxCnxAFEXpress"), Pais, Ciudad)
		If Err.number <> 0 Then
			Set rs = Nothing
			Set afxGiro = Nothing
			MostrarErrorMS ""
		End If
		If afxGiro.ErrNumber <> 0 Then			
			Set rs = Nothing
			MostrarErrorAFEX afxGiro, ""
		End If
		'response.Redirect "../compartido/error.asp?description=" & Ciudad & ", Pagador:" & Seleccionado

		Response.Write "<option value=></option>"
		If  Not rs.EOF Then
			Do Until rs.eof
				If trim(rs("Codigo_agente"))<> "ME" then			
					If (UCASE(Trim(rs("codigo_agente"))) = "AF" Or UCASE(Trim(rs("codigo_agente"))) = "ME") And _
						Mail = 1 Then
				
					Else
						If Session("Categoria")<>4 Or (Session("Categoria")=4 And rs("codigo_agente")=Session("CodigoMatriz")) Then
							If UCASE(Trim(rs("codigo_agente"))) = Ucase(Trim(Seleccionado)) Then 					
								sSelect = "SELECTED"				
							Else
								sSelect = ""
							End If
							sMayMin = trim(rs("nombre_agente"))
							sMayMin = MayMin(sMayMin) 
							Response.write "<option " & sSelect & " value=" & _
							trim(rs("codigo_agente")) & ">" & _
							sMayMin & _
							" </option> "
							If Err.number <> 0 Then
								Set rs = Nothing
								Set afxGiro = Nothing
								MostrarErrorMS ""
							End If				
						End If
					End If
				End If
				rs.MoveNext
			Loop
		End If
		
		Set rs = Nothing
		Set afxGiro = Nothing
	End Function
		
	'Objetivo:	Cargar en un combo las monedas de pago para giros	
	Sub CargarMonedaGiro(ByVal Agente, ByVal Pais, Byval Ciudad, ByVal Seleccionado)
		Dim afxMn, rs
		
		On Error Resume Next
		If Seleccionado = "CLP" Then sSelect = "SELECTED" Else sSelect = ""
		If Pais = Session("PaisMatriz") or agente=Session("CodigoMGEnvio") Then
			Response.write "<option " & sSelect & " value=CLP>PESOS CHILENOS</option> "
		End If
		IF AGENTE<>Session("CodigoMGEnvio") THEN
			If Seleccionado = "USD" Then sSelect = "SELECTED" Else sSelect = "" 
			Response.write "<option " & sSelect & " value=USD>DOLAR AMERICANO</option> "
		END IF 
	End Sub
	
	Sub CargarMonedaDeposito( Byval seleccionado)
			Response.Write "<option value=></option>"
		If Seleccionado ="USD" Then sSelect = "SELECTED" Else sSelect = ""
			Response.Write "<option " & sSelect & " value=USD>DOLAR AMERICANO</option> "
		If Seleccionado ="MNL" Then sSelect = "SELECTED" else sSelect = ""
			Response.Write "<option " & sSelect & " value=MNL>MONEDA LOCAL</option> "
		
	End Sub
	
	'INTERNO-8479 MS 11-11-2016
	Function PagadorDeposita (byval Agente , byval conexion, byval pais)
		dim sSQL
		set PagadorDeposita = nothing 
		
		sSQL =" SELECT codigo_Agente " & _
		    " FROM comision with(nolock) " & _
		    " WHERE forma_pago = 1 " & _
		    " and sentido = 1 " & _
		    " and codigo_agente = '" & Agente & "'" & _
            " and pais = '" & pais & "'"  & _
		    " and (fecha_termino is null or fecha_termino >= '" & date & "')"
		
		set PagadorDeposita  = EjecutarSQLCliente (Session("afxCnxAFEXpress"), sSQL)
		
		If err.number <> 0 Then
			MostrarErrorMS "Obtener Agente deposito"
		End If	
	End Function
	
    Function PagaHomeDelivery (byval Agente , byval conexion, byval pais)
		dim sSQL
		set PagaHomeDelivery = nothing 
		
		sSQL =" SELECT codigo_Agente " & _
		    " FROM comision with(nolock) " & _
		    " WHERE lugar_pago = 0 " & _
		    " and sentido = 1 " & _
		    " and codigo_agente = '" & Agente & "'" & _
            " and pais = '" & pais & "'"  & _
		    " and (fecha_termino is null or fecha_termino >= '" & date & "')"
		
		set PagaHomeDelivery  = EjecutarSQLCliente (Session("afxCnxAFEXpress"), sSQL)
		
		If err.number <> 0 Then
			MostrarErrorMS "Obtener Agente deposito"
		End If	
	End Function
	
	' *********** JFMG 24-05-2012 *********************
	Function PagadorDiferentesMonedas(byval Agente , byval conexion, byval pais)
		dim sSQL
		set PagadorDiferentesMonedas = nothing 
		
		sSQL =" SELECT distinct c.moneda_pago, m.nombre_moneda " & _
            " FROM comision c with(nolock) " & _
	        " inner join moneda m with(nolock) on m.codigo_moneda = c.moneda_pago " & _
            " WHERE sentido = 1 " & _
	            " and codigo_agente = '" & Agente & "' " & _
	            " and fecha_inicio <= '" & date & "' " & _
	            " and (fecha_termino is null or fecha_termino >= '" & date & "') " 

        if Agente = "SW" then
              sSQL = sSQL + " and pais = '" & pais & "'" 
        else
              sSQL = sSQL + " and c.moneda_pago NOT In('USD', 'CLP') and pais = '" & pais & "'" 
        end if	            
		
		set PagadorDiferentesMonedas = EjecutarSQLCliente(Session("afxCnxAFEXpress"), sSQL)
		
		If err.number <> 0 Then
			MostrarErrorMS "Obtener Agente diferentes moneda "
		End If	
	End Function	
	' *********** FIN JFMG 24-05-2012 *****************
	
	Sub CargarBancoBCP (Byval Seleccionado)
		dim sSelect
		'Response.Write "<option value=> </option>"
	If Seleccionado = "BC" Then sSelect ="SELECTED" else sSelect = ""		
		Response.Write "<option" & sSelect & " value=BC>BANCO DE CREDITO PERU</OPTION>"		
	End Sub
	
    'miki SMC-9 MM 2015-11-30
    Function BuscaSucursalesPago (Byval sPagador, Byval sFormaPago, byval sPais)
		dim sSql
		set BuscaSucursalesPago = nothing 
        'if sFormaPago = "0" then 'efectivo
        'If sPagador = "SW" and sPais = "COL" then
		sSql =" SELECT distinct s.IdSucursalPago, s.NombreSucursalPago FROM SucursalPago s inner join TipoPagoSucursalPago t on (s.IdSucursalPago = t.IdSucursalPago) " & _
	          " WHERE s.codigo_agente = '" & sPagador & "' and  t.FormaPago = " & sFormaPago & " and CodigoPais = '" & sPais & "' order by 1"
        'else
        '    sSql =" SELECT distinct t.CodigoSucursalPago, s.IdSucursalPago, t.NombreAgenciaBanco FROM SucursalPago s inner join TipoPagoSucursalPago t on (s.IdSucursalPago = t.IdSucursalPago) " & _
	    '            " WHERE s.codigo_agente = '" & sPagador & "' and  t.FormaPago = " & sFormaPago & " and CodigoPais = '" & sPais & "' order by 1"
        'End if
        'else 'deposito
        '    sSql =" SELECT s.IdSucursalPago,s.NombreSucursalPago FROM SucursalPago s " & _
	    '            " WHERE s.codigo_agente = '" & sPagador & "' and CodigoPais = '" & sPais & "' order by 1"
        'end if
		set BuscaSucursalesPago = EjecutarSqlCLiente (Session("afxCnxAFEXpress") ,sSql )
		If Err.number <> 0 Then
			MostrarErrorMS "Obtener Sucursal de Pago"
		End IF
	End Function

    'miki SMC-9 MM 2015-11-30
    Sub CargarBancoSW (Byval sPagador, Byval sFormaPago, byval sPais)
		dim rsSucursalPago
        dim sTiraComboSucursal
        Set rsSucursalPago = BuscaSucursalesPago(sPagador, sFormaPago, sPais) 
        sTiraComboSucursal = "<option value=> </OPTION>"
        Do While Not rsSucursalPago.EOF
            'If sPagador = "SW" and sPais = "COL" then
		    sTiraComboSucursal = sTiraComboSucursal & "<option value=" & rsSucursalPago("IdSucursalPago") & ">" & rsSucursalPago("NombreSucursalPago") & "</option> "
            'else
            '    sTiraComboSucursal = sTiraComboSucursal & "<option value=" & rsSucursalPago("CodigoSucursalPago") & ">" & rsSucursalPago("NombreAgenciaBanco") & "</option> "
            'End if
		    rsSucursalPago.MoveNext
		Loop
        
        Response.Write sTiraComboSucursal
    End Sub

     Function BuscaAgenciasSucursalesPago (Byval sPagador, Byval sFormaPago, byval sPais)
		dim sSql
		set BuscaAgenciasSucursalesPago = nothing 
		sSql = "SELECT distinct CodigoSucursalPago, NombreAgenciaBanco  FROM SucursalPago s inner join TipoPagoSucursalPago t on (s.IdSucursalPago = t.IdSucursalPago) " & _
	            " WHERE s.codigo_agente = '" & sPagador & "' and  t.FormaPago = " & sFormaPago & " and CodigoPais = '" & sPais & "' order by 1"
        
		set BuscaAgenciasSucursalesPago = EjecutarSqlCLiente (Session("afxCnxAFEXpress") ,sSql )
		If Err.number <> 0 Then
			MostrarErrorMS "Obtener Sucursal de Pago"
		End IF
	End Function

     Sub CargarAgenciaBancoSW (Byval sPagador, Byval sFormaPago, byval sPais)
		dim rsAgSucursalPago
        dim sTiraComboSucursal
        Set rsAgSucursalPago = BuscaAgenciasSucursalesPago(sPagador, sFormaPago, sPais) 
        sTiraComboSucursal = "<option value=> </OPTION>"
        Do While Not rsAgSucursalPago.EOF
		    sTiraComboSucursal = sTiraComboSucursal & "<option value=" & rsAgSucursalPago("CodigoSucursalPago") & ">" & rsAgSucursalPago("NombreAgenciaBanco") & "</option> "
		    rsAgSucursalPago.MoveNext
		Loop
        
        Response.Write sTiraComboSucursal
    End Sub
	'FIN INTERNO-8479 MS 11-11-2016

    'miki SMC-9 MM 2015-11-30
    Function BuscarCuentas (Byval sPagador, Byval sBanco)
		dim sSql
		set BuscarCuentas = nothing 
		sSql =" SELECT DescTipoCuenta, NumeroSucursal FROM SucursalPago WHERE CodigoAgente = '" & sPagador & "' and  FormaPago = 1 and  NombreSucursal = '" & sBanco & "'"
		set BuscarCuentas = EjecutarSqlCLiente (Session("afxCnxAFEXpress") ,sSql )
		If Err.number <> 0 Then
			MostrarErrorMS "Obtener Cuentas"
		End IF
	End Function

    'miki SMC-9 MM 2015-11-30
    Sub CargaTipoCuentaSW
		Response.Write "<option value=> </OPTION><option value=AH>AHORROS</OPTION><option value=CC>CUENTA CORRIENTE</OPTION>"
	End Sub

	Sub CargarBancoIK (Byval Seleccionado)		
		dim sSelect
		' JFMG 25-05-2012 se comenta todo para dejar solo a INETRBANK como Banco posible
		'Response.Write "<option value=> </OPTION>"
	    'If Seleccionado = "1" Then sSelect ="SELECTED" else sSelect = ""		
		    Response.Write "<option" & sSelect & " value=1>BANCO INTERNACIONAL DEL PERU-INTERBANK</OPTION>"
	    'If Seleccionado = "2" Then sSelect ="SELECTED" else sSelect = ""		
		'    Response.Write "<option" & sSelect & " value=2>BANCO CONTINENTAL</OPTION>"
	    'If Seleccionado = "3" Then sSelect ="SELECTED" else sSelect = ""		
		'    Response.Write "<option" & sSelect & " value=3>BANCO DE CREDITO DE PERU/OPTION>"
	    'If Seleccionado = "4" Then sSelect ="SELECTED" else sSelect = ""		
		'    Response.Write "<option" & sSelect & " value=4>SCOTIABANK</OPTION>"
	    'If Seleccionado = "5" Then sSelect ="SELECTED" else sSelect = ""		
		'    Response.Write "<option" & sSelect & " value=5>CITIBANK</OPTION>"
	    'If Seleccionado = "6" Then sSelect ="SELECTED" else sSelect = ""		
		'    Response.Write "<option" & sSelect & " value=6>BANCO INTERAMERICANO DE FINANZAS</OPTION>"
	    'If Seleccionado = "7" Then sSelect ="SELECTED" else sSelect = ""		
		'    Response.Write "<option" & sSelect & " value=7>BANCO FINANCIERO</OPTION>"
	    'If Seleccionado = "8" Then sSelect ="SELECTED" else sSelect = ""		
		'    Response.Write "<option" & sSelect & " value=8>BANCO DE C0MERCIO</OPTION>"
	    'If Seleccionado = "9" Then sSelect ="SELECTED" else sSelect = ""		
		'    Response.Write "<option" & sSelect & " value=9>BANCO DE TRABAJO</OPTION>"
	End Sub
	'INTERNO-3855 MS 26-04-2015
	Sub CargarBancoFPI (Byval Seleccionado)		
		dim sSelect 
		Response.Write "<option value=> </OPTION>"
	    If Seleccionado = "1" Then sSelect ="SELECTED" else sSelect = ""		
		   Response.Write "<option " & sSelect & " value=1>BANCOLOMBIA</OPTION>"
	    If Seleccionado = "2" Then sSelect ="SELECTED" else sSelect = ""		
		   Response.Write "<option " & sSelect & " value=2>BANCO DE BOGOTÁ</OPTION>"
	    If Seleccionado = "3" Then sSelect ="SELECTED" else sSelect = ""		
		   Response.Write "<option" & sSelect & " value=3>BANCO DAVIVIENDA S.A</OPTION>"
	    If Seleccionado = "4" Then sSelect ="SELECTED" else sSelect = ""		
		   Response.Write "<option" & sSelect & " value=4>BBVA</OPTION>"
	    If Seleccionado = "5" Then sSelect ="SELECTED" else sSelect = ""		
		   Response.Write "<option" & sSelect & " value=5>BANCO DE OCCIDENTE</OPTION>"
	    If Seleccionado = "6" Then sSelect ="SELECTED" else sSelect = ""		
		   Response.Write "<option" & sSelect & " value=6>BANCO AGRARIO</OPTION>"
	    If Seleccionado = "7" Then sSelect ="SELECTED" else sSelect = ""		
		   Response.Write "<option" & sSelect & " value=7>BANCO POPULAR</OPTION>"
	    If Seleccionado = "8" Then sSelect ="SELECTED" else sSelect = ""		
		   Response.Write "<option" & sSelect & " value=8>BANCO COLPATRIA</OPTION>"
	    If Seleccionado = "9" Then sSelect ="SELECTED" else sSelect = ""		
		   Response.Write "<option" & sSelect & " value=9>BANCO GNB SUDAMERIS COLOMBIA</OPTION>"
	    If Seleccionado = "10" Then sSelect ="SELECTED" else sSelect = ""		
		   Response.Write "<option" & sSelect & " value=10>HELM BANK</OPTION>"
	    If Seleccionado = "11" Then sSelect ="SELECTED" else sSelect = ""		
		   Response.Write "<option" & sSelect & " value=11>BANCO CAJA SOCIAL</OPTION>"
	    If Seleccionado = "12" Then sSelect ="SELECTED" else sSelect = ""		
	       Response.Write "<option" & sSelect & " value=12>BANCO CITIBANK COLOMBIA</OPTION>"  
	    If Seleccionado = "13" Then sSelect ="SELECTED" else sSelect = ""		
	       Response.Write "<option" & sSelect & " value=13>FINANCIERA PAGOS INTERNACIONALES</OPTION>"
	    If Seleccionado = "14" Then sSelect ="SELECTED" else sSelect = ""		
	       Response.Write "<option" & sSelect & " value=14>COOPERATIVA FINANCIERA DE ANTIOQUIA</OPTION>"
	    If Seleccionado = "15" Then sSelect ="SELECTED" else sSelect = ""		
	       Response.Write "<option" & sSelect & " value=15>GIROS Y FINANZAS</OPTION>"
	    If Seleccionado = "16" Then sSelect ="SELECTED" else sSelect = ""		
	       Response.Write "<option" & sSelect & " value=16>BANCO DE LA REPÚBLICA</OPTION>"
	    If Seleccionado = "17" Then sSelect ="SELECTED" else sSelect = ""		
	       Response.Write "<option" & sSelect & " value=17>BANCO CORPBANCA</OPTION>"
	    If Seleccionado = "18" Then sSelect ="SELECTED" else sSelect = ""		
	       Response.Write "<option" & sSelect & " value=18>BANCOOMEVA</OPTION>"
	    If Seleccionado = "19" Then sSelect ="SELECTED" else sSelect = ""		
	       Response.Write "<option" & sSelect & " value=19>RED MULTIBANCA COLPATRIA S.A.</OPTION>"
	    If Seleccionado = "20" Then sSelect ="SELECTED" else sSelect = ""		
	       Response.Write "<option" & sSelect & " value=20>AV VILLAS</OPTION>"
	    If Seleccionado = "21" Then sSelect ="SELECTED" else sSelect = ""		
	       Response.Write "<option" & sSelect & " value=21>BANCO WWB S.A.</OPTION>"
	    If Seleccionado = "22" Then sSelect ="SELECTED" else sSelect = ""		
	       Response.Write "<option" & sSelect & " value=22>PROCREDIT</OPTION>"
	    If Seleccionado = "23" Then sSelect ="SELECTED" else sSelect = ""		
	       Response.Write "<option" & sSelect & " value=23>BANCAMIA</OPTION>"
		If Seleccionado = "24" Then sSelect ="SELECTED" else sSelect = ""		
	       Response.Write "<option" & sSelect & " value=24>BANCO PICHINCHA S.A.</OPTION>"
	    If Seleccionado = "25" Then sSelect ="SELECTED" else sSelect = ""		
	       Response.Write "<option" & sSelect & " value=25>BANCO FALABELLA S.A.</OPTION>"
	    If Seleccionado = "26" Then sSelect ="SELECTED" else sSelect = ""		
	       Response.Write "<option" & sSelect & " value=26>BANCO FINANDINA S.A.</OPTION>"
	    If Seleccionado = "27" Then sSelect ="SELECTED" else sSelect = ""		
	       Response.Write "<option" & sSelect & " value=27>BANCO MULTIBANK S.A.</OPTION>"
	    If Seleccionado = "28" Then sSelect ="SELECTED" else sSelect = ""		
	       Response.Write "<option" & sSelect & " value=28>BANCO SANTANDER</OPTION>"
	    If Seleccionado = "29" Then sSelect ="SELECTED" else sSelect = ""		
	       Response.Write "<option" & sSelect & " value=29>BANCO COOPERATIVO COOPECENTRAL</OPTION>"   

	End Sub
	'FIN INTERNO-3855 MS 26-04-2015
	
	Function TCambioMG (BYVAL conexion, _
						Byval agente )
		dim sSql
		set TCambioMG = nothing 
		
		sSql =" Select valor from tipo_cambio where codigo_agente ='ME' and fecha_termino is null" 
		
		set TCambioMG = EjecutarSqlCLiente (Session("afxCnxAFEXpress") ,sSql )
		If Err.number <> 0 Then
			MostrarErrorMS "Obtener Tipo Cambio MG"
		End IF
	End Function
	
	' JFMG 24-05-2012
	Function TCambioPagador(BYVAL conexion, _
						    Byval Agente, _
						    Byval Moneda )
		dim sSql
		set TCambioPagador = nothing 
		
		sSql =" SELECT valor " & _
		    " FROM tipo_cambio with(nolock) " & _
		    " WHERE codigo_agente = '" & Agente & "' " & _
		        " and (fecha_termino is null or fecha_termino >= '" & date & "')" & _ 
		        " and sw_tipo = 1 " & _
		        " and codigo_moneda = '" & Moneda & "'"
		set TCambioPagador = EjecutarSqlCLiente (Session("afxCnxAFEXpress") ,sSql )
		If Err.number <> 0 Then
			MostrarErrorMS "Obtener Tipo Cambio Pagador"
		End IF
	End Function	
	' FIN JFMG 24-05-2012
	
	
	Function TCambioOPeracion (BYVAL conexion )
		dim sSql
		set TCambioOPeracion = nothing 
		
		sSql =" Select valor from tipo_cambio where codigo_agente ='AF' and fecha_termino is null and sw_tipo = 1" 
		
		set TCambioOperacion = EjecutarSqlCLiente (Session("afxCnxAFEXpress") ,sSql )
		If Err.number <> 0 Then
			MostrarErrorMS "Obtener Tipo Cambio Operación"
		End IF
	End Function

	Sub CargaTipoCuentaInterbank (Byval Seleccionado)
	Dim sSelect
	
		Response.Write "<option value=> </OPTION>"
	If Seleccionado = "A" Then sSelect ="SELECTED" else sSelect = ""		
		Response.Write "<option" & sSelect & " value=A>ABONO CUENTA</OPTION>"
	If Seleccionado = "T" Then sSelect ="SELECTED" else sSelect = ""		
		Response.Write "<option" & sSelect & " value=T>TARJETA</OPTION>"
	End Sub


	Sub CargaTipoCuentaBCP (Byval Seleccionado)
	dim sSelect 
		Response.Write "<option value=> </OPTION>"
	If Seleccionado = "AH" Then sSelect ="SELECTED" else sSelect = ""		
		Response.Write "<option " & sSelect & " value=AH>AHORROS</OPTION>"
	If Seleccionado = "CC" Then sSelect ="SELECTED" else sSelect = ""		
		Response.Write "<option " & sSelect & " value=CC>CUENTA CORRIENTE</OPTION>"
	If Seleccionado = "CM" Then sSelect ="SELECTED" else sSelect = ""			
		Response.Write "<option " & sSelect & " value=CM>CUENTA MAESTRA</OPTION>"
	End Sub
	
	'INTERNO-3855 MS 26-04-2015
	Sub CargaTipoCuentaFPI (Byval Seleccionado)
	dim sSelect 
		Response.Write "<option value=> </OPTION>"
	If Seleccionado = "AH" Then sSelect ="SELECTED" else sSelect = ""		
		Response.Write "<option " & sSelect & " value=AH>AHORROS</OPTION>"
	If Seleccionado = "CC" Then sSelect ="SELECTED" else sSelect = ""		
		Response.Write "<option " & sSelect & " value=CC>CUENTA CORRIENTE</OPTION>"
	End Sub
	'INTERNO-3855 MS 26-04-2015

	'INTERNO-8479 MS 11-11-2016
	Sub CargarFormaPago (Byval seleccionado, byval pagaHomeDelivery)
		dim sSelect
	 If Seleccionado = 0 Then sSelect = "SELECTED" Else sSelect = ""
 		Response.Write "<option " & sSelect  & " value=0>EFECTIVO</option> "
 	 If Seleccionado = 1 Then sSelect = "SELECTED" Else sSelect = ""
		Response.Write "<option " & sSelect & " value=1>DEPOSITO</option> "

        If pagaHomeDelivery = true then
            If Seleccionado = 2 Then sSelect = "SELECTED" Else sSelect = ""
		        Response.Write "<option " & sSelect & " value=2>HOME DELIVERY</option> "
        End if
	End Sub
	'FIN INTERNO-8479 MS 11-11-2016
	
	Function TipoCuenta(Byval Agente ,Byval Banco, Byval seleccionado)
		dim sSQL , rsCuenta, stipo
		Set TipoDeposito = nothing
		
		sSQL = "select t.nombre as cuenta , a.tipo as tipo from  agencia_deposito a inner join tipo t on t.codigo=a.tipo " & _
				" where codigo_agente=" & EvaluarStr(agente)& "and codigo_agencia= " & EvaluarStr(banco)
		set rsCuenta = EjecutarSqlCLiente(Session("afxCnxAfexpress"),ssql)
		
		iF ERR.number <> 0 Then
			Set rsCuenta = Nothing 
			MostrarErrorMS "Cargar Tipo cuenta 1. "
		End If 

		If Not rsCuenta.EOF Then
			Response.write "<option value=''></option> "
			Do Until rsCuenta.eof	
				If UCASE(Trim(rsCuenta("tipo"))) = Ucase(Trim(Seleccionado)) Then 					
					sSelect = "SELECTED"
				Else
					sSelect = ""
				End If
				sMayMin = trim(rsCuenta("Cuenta"))
				sMayMin = MayMin(sMayMin) 
				Response.write "<option " & sSelect & " value=" & _
				trim(rsCuenta("tipo")) & ">" & _
				sMayMin & _
				" </option> "
				If Err.number <> 0 Then
					Set rsCuenta = Nothing
					MostrarErrorMS "Cargar Tipo Cuenta 2."
				End If
				rsCuenta.MoveNext
			Loop
		End If
		Set rsCuenta = Nothing	
		'-----------**************AQUI QUEDE
	End Function
	
	
	
	Function CargarBancoDeposito (Byval Agente , seleccionado)
		Dim sSQL, rsBanco
		'Response.Write SELECCIONADO
		sSQL = "select distinct(descripcion), codigo_agencia from agencia_deposito where codigo_agente = " & evaluarstr(agente) & " order by descripcion "
		Set rsBanco = ejecutarsqlcliente(Session("afxCnxAFEXpress"), sSQL)

		If Err.number <> 0 Then
			Set rsBanco = Nothing			
			MostrarErrorMS "Cargar Agencias 1."
		End If		
		
		If Not rsBanco.EOF Then
			Response.write "<option value=''></option> "
			Do Until rsBanco.eof	
				If UCASE(Trim(rsBanco("codigo_agencia"))) = Ucase(Trim(Seleccionado)) Then 					
					sSelect = "SELECTED"
				Else
					sSelect = ""
				End If
				sMayMin = trim(rsBanco("descripcion"))
				sMayMin = MayMin(sMayMin) 
				Response.write "<option " & sSelect & " value=" & _
				trim(rsBanco("codigo_agencia")) & ">" & _
				sMayMin & _
				" </option> "
				If Err.number <> 0 Then
					Set rsBanco = Nothing
					MostrarErrorMS "Cargar Agencias 2."
				End If
				rsBanco.MoveNext
			Loop
		End If
		Set rsBanco = Nothing	
		
	End Function	
	
	'Objetivo:	Cargar en un combo las monedas de pago para giros	
	Sub CargarMonedaPago(ByVal Agente, ByVal Pais, Byval Ciudad, ByVal Seleccionado)
		Dim afxMn, rs
		
		On Error Resume Next
		Set afxMn = Server.CreateObject("AFEXGiro.Giro")
		Set rs = afxMn.ObtenerMonedaPago(Session("afxCnxAFEXpress"), Agente, Pais, Ciudad)
		If Err.number <> 0 Then
			Set rs = Nothing
			Set afxMn = Nothing
			MostrarErrorMS ""
		End If
		If afxMn.ErrNumber <> 0 Then			
			Set rs = Nothing
			Exit Sub
			MostrarErrorAFEX afxMn, Agente & ", " & pais & ", " & Ciudad 
		End If		

		'response.Redirect "../compartido/error.asp?description=" & rs("codigo_pago") & ", Moneda:" & sMayMin = trim(rs("nombre_pago"))
		'Response.Write "<option value=></option>"
		If  Not rs.EOF Then			
			Do Until rs.eof	
				If UCASE(Trim(rs("codigo_pago"))) = Ucase(Trim(Seleccionado)) Then 					
					sSelect = "SELECTED"				
				Else
					sSelect = ""
				End If
				sMayMin = trim(rs("nombre_pago"))
				sMayMin = MayMin(sMayMin) 
				Response.write "<option " & sSelect & " value=" & _
				trim(rs("codigo_pago")) & ">" & _
				sMayMin & _
				" </option> "
				
				If Err.number <> 0 Then
					Set rs = Nothing
					Set afxMn = Nothing
					MostrarErrorMS ""
				End If
				rs.MoveNext
			Loop
		End If
		
		Set rs = Nothing
		Set afxMn = Nothing
	End Sub

	'Objetivo:	Devuelve el costo del envio, y actualiza los valores pasados por referencia
	Function ObtenerTarifaGiros(Byval Captador, ByVal Pagador, Byval PaisPago, ByVal CiudadPago, _
								  ByVal MonedaGiro, ByVal MonedaPago, ByVal Monto, _
								  ByRef Tarifa, ByRef GastoTransfer, ByRef ComisionCaptador, _
								  ByRef ComisionPagador, ByRef ComisionMatriz, ByRef AfectoIva, _ 
                                  Byval nLugarPago) 'INTERNO-8479 MS 11-11-2016
		Dim afx, nPrioridad, nFormaPago
		ObtenerTarifaGiros = 0
		
		On Error Resume Next
		Set afx = Server.CreateObject("AFEXGiro.Giro")
		Tarifa = 0
		GastoTransfer = 0
		ComisionCaptador = 0
		ComisionPagador = 0
		ComisionMatriz = 0
		AfectoIva = 0
		
		If Captador = Session("CodigoMGPago") Or Pagador = Session("CodigoMGEnvio") Then
			nPrioridad = afxPrioridadUrgente
		Else
			nPrioridad = afxPrioridadNormal
		End If
		
		If Session("Deposito") = True Then
			nFormaPago = afxPagoDeposito
		Else
			nFormaPago = afxPagoEfectivo
		End If
		'Response.Write monto & " " & tarifa
		'Response.End 
		'INTERNO-8479 MS 11-11-2016
        if Trim(nLugarPago) <> "" then
            if Trim(nLugarPago) = "0" then
                nLugarPago = afxPagoDomicilio
            else
                nLugarPago = afxPagoSucursal
            end if
        else
            nLugarPago = afxPagoSucursal
        end if

		ObtenerTarifaGiros = afx.CalcularCostoEnvio(Session("afxCnxAFEXpress"), Captador, Pagador, PaisPago, CiudadPago, MonedaGiro, MonedaPago, nPrioridad, nFormaPago, nLugarPago,  cCur(Monto), Tarifa, GastoTransfer, ComisionCaptador, ComisionPagador, ComisionMatriz, AfectoIva)
		'FIN INTERNO-8479 MS 11-11-2016
		
		If Err.number <> 0 Then
			Set afx = Nothing
    		MostrarErrorMS "Obtener Tarifa 1"
		End If
		If afx.ErrNumber <> 0 Then
			MostrarErrorAFEX afx, "No se han encontrado los valores de tarifa/comisión para los datos introducidos. Favor contactar al área de Giros." 'miki APPL-23390 MM 2016-05-23
		End If
		
		Set rs = Nothing
		Set afx = Nothing
	End Function

	'Objetivo:	Devuelve la comision, y actualiza los valores pasados por referencia
	Function ObtenerComisionMG(Byval PaisPago, ByVal CiudadPago, ByVal Monto, byval Moneda)
		Dim afx5, nPrioridad, sAgente, nSentido
		ObtenerComisionMG = 0
		
		On Error Resume Next
		Set afx5 = Server.CreateObject("AFEXGiro.Giro")
		
		If Session("CiudadCliente") = Session("CiudadMatriz") Then
			sAgente = Session("CodigoMatriz")
			nSentido = afxSentidoRecibido
		Else
			sAgente = Session("CodigoAgente")
			nSentido = afxSentidoEnviado
		End If
		
		
		'codigo original
		ObtenerComisionMG = afx5.ObtenerComision(Session("afxCnxAFEXpress"), sAgente, Monto, PaisPago, CiudadPago, Moneda, Moneda, nSentido, afxGiroUrgente, afxPagoSucursal, afxPagoEfectivo)
		
		If Err.number <> 0 Then
			Set afx5 = Nothing
			MostrarErrorMS "Obtener Comisión MG ó TX 1"
		End If
		If afx5.ErrNumber <> 0 Then			
			MostrarErrorAFEX afx5, "Obtener Comisión MG ó TX 2"
		End If
		
		Set rs = Nothing
		Set afx5 = Nothing
	End Function

	' Objetivo   : Realizar el formateo de los numeros para colocarlos en una consulta sql
	' Parametros : Numero    Número a formatear
	' Devuelve   : El número formateado como string con decimales para SQL
	Function FormatoNumeroSQL(ByVal Numero)
	   Dim sFormato
	   
	   sFormato = CStr(Numero)
	   sFormato = Replace(Numero, ",", ".")
	   FormatoNumeroSQL = sFormato

	End Function


	'Objetivo	:	Verifica si el usuario está bloqueado debido al cierre en AFEXpress
	'Devuelve	:	Falso si el usuario está bloqueado
	Function VerificarBloqueoUsuario()
		Dim rUsuario		' recordset para verificar si el usuario está bloquedao para grabar
		' crea la conexión con la base de datos
		
		VerificarBloqueoUsuario = True
		Exit Function
		
		set GLB_Base = server.createobject("ADODB.Connection")
		GLB_Base.open "afex_giros","Giros","giros"

		' consulta el usuario
		sSQL = "select bloqueado " & _
			   "from usuario " & _
			   "where codigo_usuario='" & trim(session("Usu_Codigo")) & "'"
		If NOT GFNC_Consulta(rUsuario, sSQL) Then
			Set rUsuario = Nothing
			session("Des_Problema") = "Problemas al verificar el usuario. "
			session("Pag_Anterior") = "javascript:history.back()"
			response.redirect "Giros_Error.asp"
		ElseIf NOT rUsuario.EOF AND rUsuario("bloqueado") = 1 Then 
			Set rUsuario = Nothing
			session("Des_Problema") = "Este usuario se encuentra bloqueado por el momento, " & _
									  "intente el proceso en unos minutos más. "
			session("Pag_Anterior") = "javascript:history.back()"
			response.redirect "Giros_Error.asp"
		End If
		Set rUsuario = Nothing
		GLB_Base.close
		Set GLB_Base = Nothing
		'******************************************
	End Function

	Function AgregarClienteXP
		Dim afxClienteXP, sNombre, sApellido, sTarjeta, sSexo
		Dim sSQL, rsCliente ' APPL-9009
		
		If Trim(Request.Form("txtNombres")) = "" Then
			sNombre = Trim(Request.Form("txtRazonSocial"))
			sApellido = "EMPRESA"
			sSexo = 0 
		Else
			sNombre = Trim(Request.Form("txtNombres"))
			sApellido = Trim(Trim(Request.Form("txtApellidoP")) & " " & Trim(Request.Form("txtApellidoM")))
			sSexo = Request.Form("cbxSexo")
		End If
       
		sTarjeta = Request.Form("txtTarjeta")
	 	if trim(sTarjeta) <> "" then sTarjeta = trim(Request.Form("txtTarjeta1")) & right(trim(Request.Form("txtTarjeta")),6)
	 	
	 	' APPL-9009
		'Set afxClienteXP = Server.CreateObject("AfexClienteXP.Cliente")
		'AgregarClienteXP = afxClienteXP.Agregar(Session("afxCnxAFEXpress"), "", Session("CodigoAgente"), _
		'						 Request.Form("txtRut"), Request.Form("txtPasaporte"), Request.Form("cbxPaisPasaporte"), _
		'						 sNombre, sApellido, Request.Form("txtFechaNacimiento"), request.Form("txtDireccion"), _
		'						 request.Form("cbxComuna"), request.Form("cbxCiudad"), request.Form("cbxPais"), _
	 	'						 CInt(0 & request.Form("txtPaisFono")), CInt(0 & request.Form("txtAreaFono")), CCur(0 & request.Form("txtFono")), _
	 	'						 CInt(0 & request.Form("txtAreaFono2")), CCur(0 & request.Form("txtFono2")), _
	 	'						 sTarjeta, sSexo, Request.Form("cbxNacionalidad"), Request.Form("txtCorreoElectronico"))
	 	Dim sRut, sPasaporte, sPaisPasaporte, sFechaNacimiento

	 	If Request.Form("txtRut") <> "" Then
            sRut = ValorRut(Request.Form("txtRut"))
            sPasaporte = ""
            sPaisPasaporte = ""        
        Else
            sRut = ""
            sPasaporte = Request.Form("txtPasaporte")
            sPaisPasaporte = Request.Form("cbxPaisPasaporte")        
        End If
        
        If Request.Form("txtFechaNacimiento") <> "" then
            sFechaNacimiento = FormatoFechaSQL(Request.Form("txtFechaNacimiento"))
        Else
            sFechaNacimiento = ""
        End If

	 	sSQL = "exec InsertarCliente NULL, " & EvaluarStr(sRut) & ", " & EvaluarStr(sPasaporte) & ", " & _
                               EvaluarStr(sPaisPasaporte) & ", " & EvaluarStr(sNombre) & ", " & _
                               EvaluarStr(sApellido) & ", " & EvaluarStr(sFechaNacimiento) & ", " & _
                               EvaluarStr(request.Form("txtDireccion")) & ", " & EvaluarStr(request.Form("cbxComuna")) & ", " & _
                               EvaluarStr(request.Form("cbxCiudad")) & ", " & EvaluarStr(request.Form("cbxPais")) & ", " & _
                               CInt("0" & request.Form("txtPaisFono")) & ", " & CInt("0" & request.Form("txtAreaFono")) & ", " & _
                               CCur("0" & request.Form("txtFono")) & ", " & CInt("0" & request.Form("txtAreaFono2")) & ", " & _
                               CCur("0" & request.Form("txtFono2")) & ", " & EvaluarStr(sTarjeta) & ", " & EvaluarVar(sSexo, "0") & ", " & _
                               EvaluarStr(Request.Form("cbxNacionalidad")) & ", " & EvaluarStr(Request.Form("txtCorreoElectronico")) & ", " & _
                               " NULL, NULL, NULL, NULL, NULL, NULL, NULL, " & CCur("0" & request.Form("txtNumeroCelular"))
	 	Set rsCliente = EjecutarSqlCliente(Session("afxCnxAfexpress"), sSQL)
		If Err.number <> 0 Then
			'Set afxClienteXP = Nothing
			MostrarErrorMS "Agregar Cliente Giros 1"
		Else
		    AgregarClienteXP = rsCliente("codigo")
		End If
		'If afxClienteXP.ErrNumber <> 0 Then
		'	MostrarErrorAFEX afxClienteXP, "Agregar Cliente Giros 2"
		'End If
		'Set afxClienteXP = Nothing
        ' FIN APPL-9009
	End Function

	Function AgregarClienteXC(ByVal TipoCliente)
		Dim afxClienteXC, sCodigoXC, sRutXC, sTarjeta, s

		sTarjeta = Request.Form("txtTarjeta")
	 	if trim(sTarjeta) <> "" then sTarjeta = trim(Request.Form("txtTarjeta1")) & right(trim(Request.Form("txtTarjeta")),6)
	 							 
		'MostrarErrorMS Request.Form("txtRut") & ", ok"
		sRutXC = Request.Form("txtRut")


			
		's = Session("afxCnxAFEXchange") & ", , " & Session("CodigoAgente") & ", " & _
		' sRutXC & ", " & request.Form("txtPasaporte") & ", " & Request.Form("cbxPaisPasaporte") & ", " & TipoCliente & ", " & _
		' Request.Form("txtApellidoP") & ", " & Request.Form("txtApellidoM") & ", " & Request.Form("txtNombres") & ", ," &  _
		' Request.Form("txtRazonSocial") & ", " & request.Form("txtDireccion") & ", " & request.Form("cbxComuna") & ", " & _
		' request.Form("cbxCiudad") & ", " & CInt(0 & request.Form("txtPaisFono")) & ", " & CInt(0 & request.Form("txtAreaFono")) & ", " & _
		' CCur(0 & request.Form("txtFono")) & ",,,,,,,,,,,,,," & Request.Form("txtCorreoElectronico") & ",,,,,,,,," & sTarjeta & ", " & _
		'Request.Form("cbxNacionalidad") & ", " & Request.Form("cbxSexo") & ", " & Request.Form("txtFechaNacimiento")

		
		'response.write s
		'response.end



		Set afxClienteXC = Server.CreateObject("AfexClienteXC.Cliente")
		sCodigoXC = afxClienteXC.Agregar(Session("afxCnxAFEXchange"), , Session("CodigoAgente"), _
								 sRutXC , request.Form("txtPasaporte"), Request.Form("cbxPaisPasaporte"), TipoCliente, _
								 Request.Form("txtApellidoP"), Request.Form("txtApellidoM"), Request.Form("txtNombres"), , _
								 Request.Form("txtRazonSocial"), request.Form("txtDireccion"), request.Form("cbxComuna"), _
								 request.Form("cbxCiudad"), CInt(0 & request.Form("txtPaisFono")), CInt(0 & request.Form("txtAreaFono")), _
								 CCur(0 & request.Form("txtFono")),,,,,,,,,,,,,,Request.Form("txtCorreoElectronico"),,,,,,,,,,sTarjeta, Request.Form("cbxNacionalidad"), _
								 Request.Form("cbxSexo"), Request.Form("txtFechaNacimiento"))
		If Err.number <> 0 Then
			Set afxClienteXC = Nothing
			MostrarErrorMS "Agregar Cliente Cambios 1"
		End If
		If afxClienteXC.ErrNumber <> 0 Then
			MostrarErrorAFEX afxClienteXC, "Agregar Cliente Cambios 2"
		End If
		Set afxClienteXC = Nothing
		AgregarClienteXC = sCodigoXC

	End Function	

	Function ObtenerObservado()
		Dim afxObs, nObs
		ObtenerObservado = 0
		
		Set afxObs = Server.CreateObject("AFEXGiro.Giro")
		nObs = afxObs.ObtenerObservado(Session("afxCnxAFEXpress"))
		
		If Err.number <> 0 Then
			Set afxObs = Nothing
			MostrarErrorMS "Obtener Observado 1"
		End If
		If afxObs.ErrNumber <> 0 Then
			MostrarErrorAFEX afxObs, "Obtener Observado 2"
		End If
		Set afxObs = Nothing
		ObtenerObservado = nObs
	End Function		

	Function ObtenerMonedas()
		Dim afx1
		Dim rsMoneda1
				
		Set afx1 = Server.CreateObject("AfexWeb.Web")
		Set rsMoneda1 = afx1.ObtenerMonedasTransfer(Session("afxCnxAFEXchange"))
		If Err.number <> 0 Then
			Set rsMoneda1 = Nothing
			Set afx1 = Nothing
			MostrarErrorMS "Obtener Monedas 1"
		End If
		If afx1.ErrNumber <> 0 Then			
			Set rsMoneda1 = Nothing
			MostrarErrorAFEX afx1, "Obtener Monedas 2"
		End If
		
		Set afx1 = Nothing
		Set ObtenerMonedas = rsMoneda1
		Set rsMoneda1 = Nothing
		
	End Function	

	Sub CargarParidades(ByVal Tipo, ByVal Seleccionado)
		Dim rsP, sMayMin, sSelect, sTipo

		Select Case Tipo
		Case afxTCCompra
			sTipo = "Compra"
		Case afxTCVenta
			sTipo = "Venta"
		Case afxTCTransferencia
			sTipo = "ParidadTransfer"
		Case afxTCParidad
			sTipo = "Paridad"
		End Select
		
		Set rsP = ObtenerMonedas()
		If  Not rsP.EOF Then		
			Do Until rsP.eof	
				If UCASE(Trim(rsP("codigo_moneda"))) = Ucase(Trim(Seleccionado)) Then 					
					sSelect = "SELECTED"				
				Else
					sSelect = ""
				End If
				Response.write "<option " & sSelect & " value=" & _
				trim(rsP("codigo_moneda")) & "> " & rsP(sTipo) & " </option> "
				If Err.number <> 0 Then
					Set rsP = Nothing
					MostrarErrorMS "Cargar Paridades 1"
				End If
				rsP.MoveNext
			Loop
		End If
		Set rsP = Nothing	
	End Sub

	Sub CargarMonedas(ByVal Seleccionado)
		Dim rs, sMayMin, sSelect

				
		Set rs = ObtenerMonedas()
		
		If  Not rs.EOF Then
			Do Until rs.eof	
				If UCASE(Trim(rs("codigo_moneda"))) = Ucase(Trim(Seleccionado)) Then 					
					sSelect = "SELECTED"				
				Else
					sSelect = ""
				End If
				sMayMin = trim(rs("alias_moneda"))
				sMayMin = MayMin(sMayMin) 
				Response.write "<option " & sSelect & " value=" & _
				trim(rs("codigo_moneda")) & ">" & _
				sMayMin & _
				" (" & UCase(trim(rs("codigo_moneda"))) & ")" & _
				" </option> "
				If Err.number <> 0 Then
					Set rs = Nothing
					MostrarErrorMS "Cargar Monedas 1"
				End If
				rs.MoveNext
			Loop
		End If
		Set rs = Nothing	
	End Sub

	Function ObtenerGiro(ByVal Codigo)
		Dim afxGiro2, rsO
		
		On Error Resume Next
		Set afxGiro2 = Server.CreateObject("AFEXGiro.Giro")
		Set rsO = afxGiro2.Buscar(Session("afxCnxAFEXpress"), Codigo)
		If Err.number <> 0 Then
			Set rs = Nothing
			Set afxGiro2 = Nothing
			MostrarErrorMS "Obtener Giro 1"
		End If
		If afxGiro2.ErrNumber <> 0 Then			
			Set rs = Nothing
			MostrarErrorAFEX afxGiro2, "Obtener Giro 2"
		End If
		
		Set ObtenerGiro = rsO
		Set rsO = Nothing
		Set afxGiro2 = Nothing
	End Function

	Function CargarGiro(ByVal Codigo)
		Dim afxGiro3, bOk
		
		On Error Resume Next
		Set afxGiro3 = Server.CreateObject("AFEXGiro.Giro")
		bOk = afxGiro3.Cargar(Session("afxCnxAFEXpress"), Codigo)
		If Err.number <> 0 Then
			Set afxGiro3 = Nothing
			MostrarErrorMS "Cargar Giro 1"
		End If
		If afxGiro3.ErrNumber <> 0 Then			
			MostrarErrorAFEX afxGiro3, "Cargar Giro 2"
		End If
		If Not bOk Then
			Set afxGiro3 = Nothing
			Response.Redirect "../Compartido/Error.asp?Titulo=Cargar Giro 3&description=Se produjo un error desconocido al cargar el giro"
		End If
	
		Set CargarGiro = afxGiro3
		Set afxGiro3 = Nothing
	End Function

	Sub CargarReclamo()
		Response.write "<option selected value=""BEN1"">Beneficiario se cambió de casa</option>"
		Response.write "<option value=""BEN2"">Beneficiario incapacitado para cobrar</option>"
		Response.write "<option value=""BEN3"">Giro no recibido</option>"
		Response.write "<option value=""BEN4"">Beneficiario fallecido</option>"
		Response.write "<option value=""BEN5"">Beneficiario no quiere recibir el dinero</option>"
		Response.write "<option value=""BEN6"">Beneficiario se encuentra fuera del pais</option>"
		Response.write "<option value=""BEN7"">Confirmar nombre completo beneficiario</option>"
		Response.write "<option value=""BEN8"">Nombre del beneficiario incompleto</option>"
		Response.write "<option value=""BEN9"">Cambio de beneficiario</option>"
		Response.write "<option value=""CTA1"">Número de Cta. Cte. no existe</option>"
		Response.write "<option value=""CTA2"">Número de Cta. de Ahorro no existe</option>"
		Response.write "<option value=""CTA3"">Banco no corresponde a la cuenta</option>"
		Response.write "<option value=""CTA4"">Número de Cta. incompleto</option>"
		Response.write "<option value=""CTA5"">Indicar tipo de Cta. bancaria Pesos o USD</option>"
		Response.write "<option value=""ENV1"">Cliente reclama demora en entrega</option>"
		Response.write "<option value=""ENV2"">Enviar conf. de pago urgente</option>"
		Response.write "<option value=""ENV3"">Teléfono correcto es:</option>"
		Response.write "<option value=""ENV4"">Número de Cta. correcto es:</option>"
		Response.write "<option value=""ENV5"">Nombre beneficiario correcto es:</option>"
		Response.write "<option value=""ENV6"">Urgente informar estado giro</option>"
		Response.write "<option value=""ENV7"">Enviar Comprobante de Pago</option>"
		Response.write "<option value=""REC1"">Dirección incompleta</option>"
		Response.write "<option value=""REC2"">Confirmar ciudad de destino</option>"
		Response.write "<option value=""REC3"">Confirmar comuna de destino</option>"
		Response.write "<option value=""REC4"">Confirmar monto a pagar</option>"
		Response.write "<option value=""TEL1"">Teléfono no contesta</option>"
		Response.write "<option value=""TEL2"">Teléfono ocupado</option>"
		Response.write "<option value=""TEL3"">Teléfono fuera de servicio</option>"
		Response.write "<option value=""TEL4"">Teléfono no existe</option>"
		Response.write "<option value=""TEL5"">Teléfono no corresponde al beneficiario</option>"
		Response.write "<option value=""TEL6"">Teléfono no pertenece a la ciudad</option>"
		Response.write "<option value=""TEL7"">Enviar otro número de teléfono</option>"
		
		Response.write "<option value=""""></option>"
	End Sub	

	Function ObtenerListaGiros(ByVal Tipo, ByVal Cliente, ByVal Agente, ByVal Registros)
		Dim afxGiros, rsGiros

		On Error Resume Next
		Set afxGiros = Server.CreateObject("AFEXGiro.Giro")
		
		If Err.number <> 0 Then
			Set afxGiros = Nothing
			Set rsGiros = Nothing	
			MostrarErrorMS "Obtener Lista Giros 1"
		End If	
		Select Case Tipo
				
			Case afxGirosRecibidos
'				Response.Redirect "../Compartido/Error.asp?description=4 " & Session("afxCnxAFEXpress") & ", " & Tipo & ", " &  Agente & ", " &  _
'											 Cliente & ",,,,, " &  Registros & ", " & afxGirosRecibidos

				Set rsGiros = afxGiros.Lista(Session("afxCnxAFEXpress"), afxGirosRecibidos, True, Agente, "", _
								   Cliente,,,,,, Registros)

			Case afxGirosEnviados
				Set rsGiros = afxGiros.Lista(Session("afxCnxAFEXpress"), afxGirosEnviados, True, Agente,, _
								   Cliente,,,,,, Registros)
	
		End Select

		If Err.number <> 0 Then
			Set afxGiros = Nothing
			Set rsGiros = Nothing	
			MostrarErrorMS "Obtener Lista Giros 2"
		End If	
		If afxGiros.ErrNumber <> 0 Then
			Set rsGiros = Nothing
			MostrarErrorAFEX afxGiros, "Obtener Lista Giros 3"
		End If
		
		Set ObtenerListaGiros = rsGiros	
		Set afxGiros = Nothing
	End Function
	
	Function ObtenerUltimosGiros(ByVal Tipo, ByVal Cliente, ByVal Agente, ByVal Registros)
		Dim afxGiros, rsGiros

		On Error Resume Next
		Set afxGiros = Server.CreateObject("AFEXGiro.Giro")
		
		If Err.number <> 0 Then
			Set afxGiros = Nothing
			Set rsGiros = Nothing	
			MostrarErrorMS "Obtener Lista Giros 1"
		End If	
		Select Case Tipo
				
			Case afxGirosRecibidos
				Set rsGiros = afxGiros.Lista(Session("afxCnxAFEXpress"), afxGirosRecibidos, True, Agente, "", _
								   Cliente,,,,,, Registros)

			Case afxGirosEnviados
				' Jonathan Miranda G. 16-11-20006
				'Set rsGiros = afxGiros.Lista(Session("afxCnxAFEXpress"), afxGirosEnviados, True, Agente,, _
				'				   Cliente,,,,,, Registros)
				
				sSQL = " execute UltimosGiros " & Registros & ", null, " & EvaluarStr(Cliente) & ", " & _
															EvaluarStr(Agente) & ", null "
				Set rsGiros = EjecutarSQLCliente(Session("afxCnxAFEXpress"), sSQL)								   
				'----------------------------- Fin -------------------------------				
	
		End Select
		Set ObtenerUltimosGiros = Nothing
		
		If Err.number <> 0 Then
			Set afxGiros = Nothing
			Set rsGiros = Nothing	
			Exit Function
		End If	
		If afxGiros.ErrNumber <> 0 Then
			Set rsGiros = Nothing
			Exit Function
		End If
		
		Set ObtenerUltimosGiros = rsGiros	
		Set afxGiros = Nothing
	End Function
	
	Function FormatoFechaSQL(Byval Fecha)
		FormatoFechaSQL = Year(Fecha) & Right("0" & Month(Fecha), 2) & Right("0" & Day(Fecha), 2)
	End Function	
	

	'Objetivo:	Devuelve el costo del envio, y actualiza los valores pasados por referencia
	Function ObtenerTarifaUltimoGiro(Byval Captador, ByVal Pagador, Byval PaisPago, ByVal CiudadPago, _
								  ByVal MonedaGiro, ByVal MonedaPago, ByVal Monto, _
								  ByRef Tarifa, ByRef GastoTransfer, ByRef ComisionCaptador, _
								  ByRef ComisionPagador, ByRef ComisionMatriz, ByRef AfectoIva, _
                                  Byval nLugarPago)'INTERNO-8479 MS 11-11-2016
		Dim afx, nPrioridad
		ObtenerTarifaUltimoGiro = 0
		
		On Error Resume Next
		Set afx = Server.CreateObject("AFEXGiro.Giro")
		Tarifa = 0
		GastoTransfer = 0
		ComisionCaptador = 0
		ComisionPagador = 0
		ComisionMatriz = 0
		AfectoIva = 0
		
		If Captador = Session("CodigoMGPago") Or Pagador = Session("CodigoMGEnvio") Then
			nPrioridad = afxPrioridadUrgente
		Else
			nPrioridad = afxPrioridadNormal
		End If
		'MostrarErrorMS Pagador & ", 3" & nPrioridad & ", " & afxPrioridadUrgente

        'INTERNO-8479 MS 11-11-2016
		if Trim(nLugarPago) <> "" then
            if Trim(nLugarPago) = "0" then
                nLugarPago = afxPagoDomicilio
            else
                nLugarPago = afxPagoSucursal
            end if
        else
            nLugarPago = afxPagoSucursal
        end if
		ObtenerTarifaUltimoGiro = afx.CalcularCostoEnvio(Session("afxCnxAFEXpress"), Captador, Pagador, PaisPago, CiudadPago, MonedaGiro, MonedaPago, nPrioridad, afxPagoEfectivo, nLugarPago,  cCur(Monto), Tarifa, GastoTransfer, ComisionCaptador, ComisionPagador, ComisionMatriz, AfectoIva)
		'FIN INTERNO-8479 MS 11-11-2016
		
		If Err.number <> 0 Then
			Set afx = Nothing
			MostrarErrorMS "Obtener Tarifa 1"
		End If
		
		Set rs = Nothing
		Set afx = Nothing
	End Function

	Function ValidarEstadoBD
		Dim afxBD, bOkBD
		
		Set afxBD = Server.CreateObject("AfexWebXP.web")
		bOKBD = afxBD.ObtenerEstadoBDGiros(Session("afxCnxAFEXPress"))
		If Err.number <> 0 Then
			Set afxBD = Nothing
			MostrarErrorMS "Validar EstadoBD 1"
		End If
		If afxBD.ErrNumber <> 0 Then
			MostrarErrorAFEX  afxBD, "Validar EstadoBD 2"
		End If
		
		ValidarEstadoBD = bOkBD
		Set afxBD = Nothing
		
	End Function

	'Objetivo:	Validar el esatdo de la BD de Giros. Usuarios bloqueados
	'				Cierre todos los objetos antes de ejecutar esta funcion
	Function ValidarBDGiros
		Dim rsBD1, sSQL2, nBloqueado
		
		ValidarBDGiros = True
		'response.Redirect "http:compartido/informacion.asp?Tipo=0&detalle=" & Time()
		
		If Time() < "08:30:00" Or Time() > "10:00:00" Then Exit Function
		sSQL2 = "SELECT COUNT(bloqueado) AS Bloqueado FROM Usuario WHERE bloqueado = 1 "
		Set rsBD1 = EjecutarSQLCliente(Session("afxCnxAFEXpress"), sSQL2)
		If Err.number <> 0 Then
			Set rsBD1 = Nothing
			MostrarErrorMS "Validar EstadoBD 1"
		End If
		If rsBD1 Is Nothing Then Exit Function
		nBloqueado = rsBD1("Bloqueado")
		Set rsBD1 = Nothing
		If nBloqueado > 0 Then ValidarBDGiros = False
	End Function


	Function ValidarGiro(Byval Boleta, Byval Remitente, Byval Monto, Byval NombreBeneficiario, _
						Byval ApellidoBeneficiario)
		Dim bExiste, sSQLg, rsGiro11
		
		ValidarGiro = True
		If Boleta <> 0 Then
            ' Tecnova rperez 29-08-2016 - Sprint 2 Cambio de condicion, comprueba (agente_capturador) y ((la existencia de similitud de beneficiario) o (numero de boleta))
			sSQLg = "SELECT COUNT(codigo_giro) AS Giros " & _
				"FROM Giro WITH(nolock) " & _
				"WHERE ((codigo_remitente = '" & Remitente & "' " & _
				"AND   monto_giro = " & FormatoNumeroSQL(Monto) & " " & _
				"AND   Fecha_captacion = '" & Date & "' " & _
				"AND   estado_giro <> 9 " & _
				"AND   nombre_beneficiario='" & NombreBeneficiario & "' " & _
				"AND   apellido_beneficiario='" &ApellidoBeneficiario & "') " & _
				"OR    (numero_documento = " & Boleta & ")) " & _
				"AND   (agente_captador = '" & Session("CodigoAgente") & "') " & _
				"UNION SELECT 99999 AS Giros ORDER BY Giros"

		Else
            ' Tecnova rperez 29-08-2016 - Sprint 2 Agrega condicion para agente_captador
			sSQLg = "SELECT COUNT(codigo_giro) AS Giros FROM Giro WITH(nolock) WHERE (codigo_remitente = '" & Remitente & "' AND monto_giro = " & FormatoNumeroSQL(Monto) & " AND Fecha_captacion = '" & Date & "' AND estado_giro <> 9 AND nombre_beneficiario='" & NombreBeneficiario & "' AND agente_captador = '" & Session("CodigoAgente") & "') UNION SELECT 99999 AS Giros ORDER BY Giros"		
		End If
		Set rsGiro11 = EjecutarSQLCliente(Session("afxCnxAFEXpress"), sSQLg)
		'If rsGiro11 Is Nothing Then 
		'	ValidarGiro = False
		'	Exit Function
		'End If
		'rsGiro11.MoveNext
		'mostrarerrorms rsGiro11.recordCount
		If rsGiro11("Giros") > 0 Then
			ValidarGiro = True
		Else
			ValidarGiro = False
		End If		
		rsGiro11.close
		Set rsGiro11 = Nothing
		
	End Function

	Public Function ValidarTransfer(Byval Cliente, Byval Beneficiario, Byval Moneda, Byval Monto)
		Dim bExiste, sSQLt, rsGiro12
		
		ValidarTransfer = True
		If Beneficiario <> "" Then
			sSQLt = "SELECT COUNT(correlativo_transferencia) AS Transfers FROM Transferencia WITH(nolock) WHERE (codigo_cliente = '" & Cliente & "' AND nombre_titular_destino = '" & Beneficiario & "' AND monto_transferencia = " & FormatoNumeroSQL(Monto) & " AND Fecha_transferencia = '" & Date & "' AND codigo_moneda = '" & Moneda & "' AND estado_transferencia <> 0) UNION SELECT 99999 AS Transfers ORDER BY Transfers"
		Else
			sSQLt = "SELECT COUNT(correlativo_transferencia) AS Transfers FROM Transferencia WITH(nolock) WHERE (codigo_cliente = '" & Cliente & "' AND nombre_titular_destino IS Null AND monto_transferencia = " & FormatoNumeroSQL(Monto) & " AND Fecha_transferencia = '" & Date & "' AND codigo_moneda = '" & Moneda & "' AND estado_transferencia <> 0) UNION SELECT 99999 AS Transfers ORDER BY Transfers"
		End If
		
		Set rsGiro12 = EjecutarSQLCliente(Session("afxCnxAFEXchange"), sSQLt)
		If rsGiro12("Transfers") > 0 Then
			ValidarTransfer = True
		Else
			ValidarTransfer = False
		End If		
		rsGiro12.close
		Set rsGiro12 = Nothing
		
	End Function
	
'Objetivo:  ejecutar una consulta SQL
'Devuelve:  un recordset en el cliente desconectado
	Public Function EjecutarSQLCliente(ByVal Conexion, ByVal SQL)
		Dim rsESQL
		Dim Cnn
		Const adUseClient = 2
		Const adOpenStatic = 3
		Const adLockBatchOptimistic = 4
		Dim sError
		
		On Error Resume Next
	

		Set EjecutarSQLCliente = Nothing		
   
		Set Cnn = server.CreateObject("ADODB.Connection")
		Cnn.CommandTimeout = 600
		Cnn.Open Conexion
	   
		If Err.number <> 0 Then
			Cnn.Close
			Set Cnn = Nothing

			'mostrarerrorms sql & "//" & conexion

			MostrarErrorMS "Ejecutar SQL 1"
		End If

		Set rsESQL = server.CreateObject("ADODB.Recordset")
		rsESQL.CursorLocation = 3
		rsESQL.Open SQL, Cnn, 3, 4
	   
		If Err.number <> 0 Then		
			Set rsESQL = Nothing			
			'MostrarErrorMS "Ejecutar SQL 1"
			sError = err.description			
			err.clear			
			err.Raise 50000, "EjecutarSQLCliente", sError		
		End If
		
		if sError = empty then		
			'If rsESQL Is Nothing Then Exit Function
			Set rsESQL.ActiveConnection = Nothing
			'MostrarErrorMS "Despues"
		end if	
		
		Set EjecutarSQLCliente = rsESQL
		'rsESQL.Close
		Set rsESQL = Nothing
		
		Cnn.Close
		Set Cnn = Nothing
	End Function	
	
Function ValorRut(Byval Rut)
	ValorRut=""
	If Rut = "" Then Exit Function
   Rut = Replace(Trim(Rut), ".", "")
   Rut = Replace(Rut, "-", "")
   Rut = Replace(Rut, ",", "")
   Rut = Right("000000000" & Rut, 9)
   ValorRut = Rut
End Function

Sub EnviarEMail(ByVal Desde, Byval Para, Byval CC, ByVal Asunto, ByVal Mensaje, ByVal HTML)
	Dim objEMail
		
	On Error Resume Next
	Set objEMail = Server.CreateObject("CDONTS.NewMail")
	
	objEMail.From = Desde
	objEMail.To = Para
	objEMail.cc = CC
	objEMail.Subject = Asunto
	
	If HTML = 1 Then
		objEMail.BodyFormat = 0
		objEMail.MailFormat = 0
	Else
		objEMail.BodyFormat = 1
		objEMail.MailFormat = 1
	End If
	
	objEMail.Body = Mensaje
	objEMail.Send
	Set objEMail = Nothing	
End Sub


Function EvaluarStr(ByVal Valor)
	Dim Devuelve
	
	If Valor="" Then 
		EvaluarStr = "Null"	
	Else
		EvaluarStr = "'" & Valor & "'"
	End If

End Function

	Function BuscarTRF(ByVal Conexion, _
					   ByVal Campo, _
					   ByVal Valor, _
					   ByVal Registros, _
					   ByVal IncluirNoProcesadas, _
					   ByVal IncluirNulas)

		Dim aCampos(6)
		Dim sSQL
	   
		Const Where = 6
		
		On Error Resume Next
	   
		aCampos(1) = "correlativo_transferencia"
		aCampos(2) = "numero_transferencia"
		aCampos(3) = "codigo_cliente"
		aCampos(4) = "codigo_solitud"
		aCampos(5) = "numero_voucher"
		aCampos(6) = "WHERE "
	   
		Set BuscarTRF = Nothing
	   
		'Crea la sentencia SQL
		If Registros > 0 Then
			sSQL = "SELECT   TOP " & Registros & " * " & _
					 "FROM     VTransferencia "
		Else
			sSQL = "SELECT   * " & _
				   "FROM     VTransferencia "
		End If
		If Campo = Where Then
			sSQL = sSQL & " WHERE " & Valor
		Else
			sSQL = sSQL & "WHERE    " & aCampos(Campo) & " = '" & Valor & "' "
		End If
	      
		If Not IncluirNoProcesadas Then
			sSQL = sSQL & " And numero_transferencia is not null "
		End If
	   
		If Not IncluirNulas Then
			sSQL = sSQL & " And estado_transferencia <> " & afxTrfNulo & " "
		End If

		sSQL = sSQL & " ORDER BY correlativo_transferencia DESC"

		'Asigna al metodo el resultado de la consulta
		Set BuscarTRF = EjecutarSQLCliente(Conexion, sSQL)
	   
		If Err.Number <> 0 Then
			MostrarErrorMS "Obtener Transferencias"
		End If
			   
	End Function
	
	

	Function BuscarInformacionPagador(ByVal Conexion, _
									  ByVal CodigoAgente, _
									  ByVal CodigoPais, _
									  ByVal CodigoCiudad)

		Dim sSQL
	   
		Set BuscarInformacionPagador = Nothing
	   
		sSQL = "select * from manual where codigo_agente = " & EvaluarStr(Trim(CodigoAgente)) & " " & _
			   "and codigo_pais = " & EvaluarStr(Trim(CodigoPais)) & " " & _
			   "and codigo_ciudad = " & EvaluarStr(Trim(CodigoCiudad))
		
		'Asigna al metodo el resultado de la consulta
		Set BuscarInformacionPagador = EjecutarSQLCliente(Conexion, sSQL)
	   
		If Err.Number <> 0 Then
			MostrarErrorMS "Obtener Información Pagador"
		End If
	End Function

	Sub CargarTipos(ByVal Seleccionado, Tipo)
		Dim rs, sMayMin, sSelect, sSQL
				
		sSQL = "select nombre, codigo from tipo where tipo = " & evaluarstr(Tipo) & " order by nombre "
		Set rs = ejecutarsqlcliente(Session("afxCnxAFEXpress"), sSQL)
		If Err.number <> 0 Then
			Set rs = Nothing			
			MostrarErrorMS "Cargar Tipos 1."
		End If		
		
		If Not rs.EOF Then
			Response.write "<option value=''></option> "
			Do Until rs.eof	
				If UCASE(Trim(rs("codigo"))) = Ucase(Trim(Seleccionado)) Then 					
					sSelect = "SELECTED"
				Else
					sSelect = ""
				End If
				sMayMin = trim(rs("nombre"))
				sMayMin = MayMin(sMayMin) 
				Response.write "<option " & sSelect & " value=" & _
				trim(rs("codigo")) & ">" & _
				sMayMin & _
				" </option> "
				If Err.number <> 0 Then
					Set rs = Nothing
					MostrarErrorMS "Cargar Tipos 2."
				End If
				rs.MoveNext
			Loop
		End If
		Set rs = Nothing	
	End Sub

	' JFMG 31-07-2009 cargar un combo con los prefijos de las tarjetas giro club
	'Sub CargarPrefijoTarjeta()
	'	Response.Write "<option value=0007010>0007010</option> " & _
	'			"<option value=0008010>0008010</option> " & _
	'			"<option value=0009010>0009010</option> "
	'end sub

 ' PSS 06-05-2010 cargar combo con prefijos de las tarjetas Giro Club desde tabla Tipo en BD
	Sub CargarPrefijoTarjeta (Byval Seleccionado)
		dim sSQL, rs, sSelect 
		
		sSQL ="Execute	MuestraPrefijoTarjeta"
		
		set rs = EjecutarSQLCliente (Session("afxCnxAFEXpress"), sSQL)
		
		If err.number <> 0 Then
			MostrarErrorMS "Obtener Prefijo Tarjeta"
		End If	
		
		If not rs.eof then
			Response.write "<option value=''></option> "
			do until rs.eof
				If UCASE(Trim(rs("codigo"))) = Ucase(Trim(Seleccionado)) Then 					
					sSelect = "SELECTED"
				Else
					sSelect = ""
				End If
				
				Response.write "<option " & sSelect & " value=" & _
				trim(rs("codigo")) & ">" & _
				trim(rs("codigo")) & _
				" </option> "
				If Err.number <> 0 Then
					Set rs = Nothing
					MostrarErrorMS "Cargar Prefijo Moneda 2."
				End If
				rs.MoveNext
			loop
		end if
	End Sub
	Function ObtenerTipoCambioVenta(ByVal Codigo)
		Dim rsO
		dim sSql 
	
		ObtenerTipoCambioVenta = false
		On Error Resume Next
	
		sSql = "ObtenerTipoCambioVenta '" & codigo & "'" 
			
		Set rsO = EjecutarSQLCliente (Session("afxCnxAFEXpress"), ssql)

		'mostrarerrorms rso("codigo_giro")
		If Err.number <> 0 Then
			Set rsO = Nothing
			MostrarErrorMS "Obtener Tipo cambio venta 1"
		End If
	
		Set ObtenerTipoCambioVenta = rsO
		Set rsO = Nothing
	End Function
	
	Function ObtenerCorreoElectronicoContacto(byval Tipo)
		Dim rs
		dim sSql 
	
		ObtenerCorreoElectronicoContacto = ""
		
		On Error Resume Next
	
		sSql = " SELECT * FROM Administracion.CorreoElectronicoContacto WHERE idtipocorreoelectronicocontacto = " & Tipo 			
		
		'Response.Write ssql
		'response.end
		
		Set rs = EjecutarSQLCliente (Session("afxCnxAFEXpress"), ssql)
		
		If Err.number <> 0 Then
			Set rsO = Nothing
			MostrarErrorMS "Obtener Lista Correo Electrónico"
		End If
	
		if not rs.eof then
			ObtenerCorreoElectronicoContacto = rs("paracorreoelectronicocontacto") & "//" & rs("copiacorreoelectronicocontacto")
		end if
		
		Set rs = Nothing
	End Function

    ' INTERNO-1831 - JFMG 28-07-2014
    Sub CargarPaisPasaporte(ByRef Seleccionado)
		Dim sSelect
		Dim afxCOM
		Dim rs
		Dim sConexion
		Dim sMayMin
		Dim sCodigo
		
		On Error Resume Next
		Set rs = Server.CreateObject("ADODB.Recordset")
		If Err.number <> 0 Then response.Redirect "Error.asp?Titulo=Error en HágaseCliente&Number=" & Err.number   & "&Source=" & Err.Source & "&Description=" & Err.description		
		sConexion = Session("afxCnxCorporativa")
		sSelect = Empty

		Set afxCOM = Server.CreateObject("AfexCorporativo.Pais")				
		Set rs = afxCOM.Buscar(sConexion)
		sCodigo = Seleccionado				
	
		If Err.number <> 0 Then 
			MostrarErrorMS "Cargar País pasaporte"
		End If
				
		If afxCOM.ErrNumber <> 0 Then
			If afxCOM.ErrNumber = 10527 Or afxCOM.ErrNumber = 10532 Then	
				Set afxCOM = Nothing	
				Exit Sub
			Else
				MostrarErrorAFEX afxCOM, "Cargar País pasaporte"  & codigo
			End If
		End If
		
		Response.Write "<option value=></option>"
		If  Not rs.EOF Then
			Do Until rs.eof
			    If trim(rs("codigo")) <> "CL" Then ' solo lo agrega si no es CHILE
				    If UCASE(Trim(rs("codigo"))) = Ucase(Trim(sCodigo)) Then 					
					    sSelect = "SELECTED"				
				    Else
					    sSelect = ""
				    End If
				    sMayMin = trim(rs("nombre"))
				    sMayMin = MayMin(sMayMin) 
				    Response.write "<option " & sSelect & " value=" & _
				    trim(rs("codigo")) & ">" & _
				    sMayMin & _
				    " </option> "
				    If err.number <> 0 Then
					    MostrarErrorMS "Cargar País pasaporte"
				    End If
				End If														
				rs.MoveNext
			Loop
		End If
			
		Set afxCOM = Nothing	
	End Sub
    ' FIN INTERNO-1831 - JFMG 28-07-2014 

	
%>

