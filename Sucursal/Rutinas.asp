<!-- Rutinas.asp -->
<%

	Function CreditoUsado(ByVal CodigoClienteCP, ByVal DiasRetencion)
		Dim sSQL
		Dim rsCredito
		Dim sFecha
		Dim nCredito
		
		On Error Resume Next
		
		CreditoUsado = 0
		nCredito = 0
		
		sFecha = FormatoFechaSQL(CalcularFecha(Date, DiasRetencion, 0))
		
				
		' Jonathan 19-07-2006 canelo.cambios

		sSQL = "select (isnull(sum(total), 0) / j.tipo_cambio_observado) as Credito " & _
			   "from " & _
					"detalle_solicitud ds " & _
			   "inner join [cambiosdb.afex.cl].cambios.dbo.jornada j on j.fecha_jornada = ds.fecha " & _
			   "where " & _
					"codigo_cliente_corporativa = " & CodigoClienteCP & " " & _
			   "and   fecha >= '" & sFecha & "' " & _
			   "and   codigo_producto = 2 " & _
			   "and   estado = 1 " & _
			   "and   ds.codigo_moneda <> 'usd' " & _
			   "group by " & _
					"j.tipo_cambio_observado"
		
		Set rsCredito = EjecutarSQLCliente(Session("afxCnxCorporativa"), sSQL)
		
		If Err.number <> 0 Then
			Set rsCredito = Nothing
			Response.Redirect "http:../Compartido/Error.asp?Titulo=Calcular Credito&Number=" & Err.number   & "&Source=" & Err.Source & "&Description=" & Err.Description
			Exit Function
		End If

		
		Do Until rsCredito.EOF
			nCredito = cCur(0 & nCredito) + cCur(0 & rsCredito("credito"))
				
			rsCredito.Movenext
		Loop
		
		Set rsCredito = Nothing
		
		sSQL = "SELECT  isnull(SUM(monto), 0) as CreditoUsado " & _
			   "FROM    detalle_solicitud " & _
			   "WHERE   codigo_cliente_corporativa = " & CodigoClienteCP & " " & _
			   "and     fecha >= '" & sFecha & "' " & _
			   "and     codigo_producto = 2 " & _
			   "and     estado = 1 " & _
			   "and     codigo_moneda = 'usd'"

		Set rsCredito = EjecutarSQLCliente(Session("afxCnxCorporativa"), sSQL)
		
		If Err.Number <> 0 Then 
			Set rsCredito = Nothing
			MostrarErrorMS "Obtener Credito 1"
		End If
	   
		CreditoUsado = cCur(0 & rsCredito("CreditoUsado")) + cCur(0 & nCredito)
		
		Set rsCredito = Nothing
	End Function
	
	Function CalcularFecha(ByVal Fecha, _
	                       ByVal Dias, _
						   ByVal Operador)
	   Dim i
	   
	   CalcularFecha = Fecha
	   For i = 1 To Dias
	      
	      If Operador = 0 Then
	         Fecha = Fecha - 1
	      Else
	         Fecha = Fecha + 1
	      End If
	      
	      Select Case Weekday(Fecha)
	         Case vbSunday
	               If Operador = 0 Then
	                  Fecha = Fecha - 2
	                  'Fecha = Fecha - 3
	               Else
	                  Fecha = Fecha + 1
	                  'Fecha = Fecha + 2
	               End If
	               'i = i + 1
	               
	         Case vbSaturday
	               If Operador = 0 Then
	                  Fecha = Fecha - 1
	                  'Fecha = Fecha - 2
	               Else
	                  Fecha = Fecha + 2
	                  'Fecha = Fecha + 3
	               End If
	               'i = i + 1
	               
	      End Select
	   Next
	   
	   CalcularFecha = Fecha
	End Function

	Sub AgregarHistoria(ByVal CodigoCliente, ByVal Descripcion, ByVal Tipo, _
						Byval TipoAutorizacion) ' JFMG 16-04-2009 se agrega tipo autorizacion (Byval TipoAutorizacion)
						
		Dim afxConexion 
		Dim sSQL
		Dim BD
		
		On Error Resume Next
		
		' JFMG 16-04-2009 se agrega tipo autorizacion
		if cint("0" & TipoAutorizacion) = 0 then TipoAutorizacion = "null"
		' *********** FIN ***************

		afxConexion = Session("afxCnxCorporativa") '"DSN=AfexCorporativa;UID=corporativa;PWD=afxsqlcor;"	
		
		Set BD = Server.CreateObject("ADODB.Connection")
			
		BD.Open afxConexion                          'Abre la conexion
		If Err.Number <> 0 Then 
			Set BD = Nothing
			Response.Redirect "http:../compartido/error.asp?description=" &  "GuardarNuevaHistoria"
		End If	
		
		'MS 20/06/2013 Se agrega ip y nombre del pc del usuario 
		Dim  ipUsuario, nombrePC
		ipusuario = request.servervariables("REMOTE_ADDR")
	    nombrePC = request.servervariables("REMOTE_HOST")
	    
	    
		sSQL = "InsertarHistoria '" & FormatoFechaSQL(Date) & "', " & EvaluarStr(Left(Time, 8)) & ", " & _
											 CodigoCliente & ", " & EvaluarStr(Descripcion) & ", " & _
											 EvaluarStr(Session("NombreUsuarioOperador")) & ", " & Tipo & _
											 ", " & TipoAutorizacion & ", '" & ipusuario & "', '" & nombrePC & "','" & _
											  Session("NombreOperador") & "',1"
											' JFMG 16-04-2009 se agrega tipo autorizacion
											' MS 20/06/2013 Se agrega ip y nombre del pc del usuario 
											' MS 18/03/2014 Se agrega nombre del usuario 
		                                    
		'BD.BeginTrans
		
		BD.Execute sSQL

	    If Err.Number <> 0 Then 
			'BD.RollbackTrans
			Set BD = Nothing
			Response.Redirect "http:../compartido/error.asp?description=" &  Err.number & "<br>" & err.Description
			Exit Sub
	    End If
	    
	    'BD.CommitTrans
	    
	    Set BD = Nothing
	End Sub	


	Function ObtenerFichaCliente()
		Dim rs, sSQL
		
		Set ObtenerFichaCliente = Nothing
		
	   On Error Resume Next
	   sSQL = "SELECT	cf.*, gf.nombre AS nombre_grupo " & _
				 "FROM Configuracion_Ficha CF " & _
				 "LEFT OUTER JOIN Grupo_Ficha GF ON gf.codigo=cf.grupo " & _
				 "WHERE estado = 1 " & _
				 "ORDER BY grupo, orden "
	   	   
	   Set rs = EjecutarSQLCliente(Session("afxCnxCorporativa"), sSQL)

	   If Err.Number <> 0 Then 
			Set rs = Nothing
			MostrarErrorMS "Obtener Ficha 1"
		End If

		Set ObtenerFichaCliente = rs
		Set rs = Nothing
		
	End Function


	Function ObtenerProducto()
		Dim rs, sSQL
		
		Set ObtenerProducto = Nothing
		
	   On Error Resume Next
	   sSQL = "SELECT	* " & _
				 "FROM	Producto " & _
				 "ORDER BY codigo "
	   	   
	   Set rs = EjecutarSQLCliente(Session("afxCnxCorporativa"), sSQL)

	   If Err.Number <> 0 Then 
			Set rs = Nothing
			MostrarErrorMS "Obtener Producto 1"
		End If

		Set ObtenerProducto = rs
		Set rs = Nothing
		
	End Function

	Function ObtenerProductoFicha()
		Dim rs, sSQL
		
		Set ObtenerProductoFicha = Nothing
		
	   On Error Resume Next
	   sSQL = "SELECT pr.codigo AS Producto, cf.correlativo AS Campo, CASE WHEN fc.campo IS NOT NULL AND fc.producto IS NOT NULL THEN 1 ELSE 0 END AS estado, " & _
				 "			fc.monto_desde " & _
				 "FROM	producto pr " & _
				 "LEFT OUTER JOIN configuracion_ficha cf on cf.correlativo=cf.correlativo " & _
				 "LEFT OUTER JOIN ficha_producto fc on fc.campo=cf.correlativo and pr.codigo=fc.producto AND fc.tipo = 'MCS' " & _
				 "ORDER BY cf.grupo, cf.orden, pr.codigo " 
	   	   
	   Set rs = EjecutarSQLCliente(Session("afxCnxCorporativa"), sSQL)

	   If Err.Number <> 0 Then 
			Set rs = Nothing
			MostrarErrorMS "Obtener Producto Ficha 1"
		End If

		Set ObtenerProductoFicha = rs
		Set rs = Nothing
		
	End Function


	Sub CargarBanco(ByRef Seleccionado)
		Dim sSelect
		Dim rs
		Dim sConexion
		Dim sMayMin
		Dim sCodigo
		
		On Error Resume Next
		Set rs = Server.CreateObject("ADODB.Recordset")
		If Err.number <> 0 Then response.Redirect "Error.asp?Titulo=Error en Cargar Banco&Number=" & Err.number   & "&Source=" & Err.Source & "&Description=" & Err.description		
		sSelect = Empty

		Set rs = ObtenerBanco()
		sCodigo = Seleccionado
		If Err.number <> 0 Then 
			'response.Redirect "../Compartido/Error.asp?Titulo=Errores en HágaseCliente&Number=" & Err.number   & "&Source=" & Err.Source & "&Description=" & Err.description
			MostrarErrorMS "Cargar Banco 1"
		End If
				
		Response.Write "<option value=0>NINGUNO</option>"
		If  Not rs.EOF Then
			Do Until rs.eof	
				If UCASE(Trim(rs("codigo_banco"))) = Ucase(Trim(sCodigo)) Then 					
					sSelect = "SELECTED"				
				Else
					sSelect = ""
				End If
				sMayMin = trim(rs("nombre_banco"))
				sMayMin = MayMin(sMayMin) 
				Response.write "<option " & sSelect & " value=" & _
				trim(rs("codigo_banco")) & ">" & _
				sMayMin & _
				" </option> "
				If err.number <> 0 Then 
					'response.Redirect "../Compartido/Error.asp?Titulo=Error en HágaseCliente&Number=" & Err.Number  & "&Source=" & Err.source  & "&Description=" & Err.description	
					MostrarErrorMS "Cargar banco 3"
				End If																
				rs.MoveNext
			Loop
		End If
			
	End Sub

	Sub CargarEstadoTRF(ByRef Seleccionado)
		Dim sSelect
		Dim rs
		Dim sConexion
		Dim sMayMin
		Dim sCodigo
		
		On Error Resume Next
		Set rs = Server.CreateObject("ADODB.Recordset")
		If Err.number <> 0 Then response.Redirect "../Compartido/Error.asp?Titulo=Error en Cargar Estado Transfer&Number=" & Err.number   & "&Source=" & Err.Source & "&Description=" & Err.description		
		sSelect = Empty

		Set rs = ObtenerEstadoTRF()
		sCodigo = Seleccionado
		If Err.number <> 0 Then 
			'response.Redirect "../Compartido/Error.asp?Titulo=Errores en HágaseCliente&Number=" & Err.number   & "&Source=" & Err.Source & "&Description=" & Err.description
			MostrarErrorMS "Cargar Estado Transfer 1"
		End If
				
		'Response.Write "<option value=0>NINGUNO</option>"
		If  Not rs.EOF Then
			Do Until rs.eof	
				If Trim(rs("codigo_estado")) = Trim(sCodigo) Then 					
					sSelect = "SELECTED"				
				Else
					sSelect = ""
				End If
				sMayMin = trim(rs("nombre_estado"))
				sMayMin = MayMin(sMayMin) 
				Response.write "<option " & sSelect & " value=" & _
				trim(rs("codigo_estado")) & ">" & _
				sMayMin & _
				" </option> "
				If err.number <> 0 Then 
					'response.Redirect "../Compartido/Error.asp?Titulo=Error en HágaseCliente&Number=" & Err.Number  & "&Source=" & Err.source  & "&Description=" & Err.description	
					MostrarErrorMS "Cargar Estado Transfer 3"
				End If																
				rs.MoveNext
			Loop
		End If
	End Sub

	Function ObtenerEstadoTRF()
	   Dim rsEst
	   Dim sSQL

	   Set ObtenerEstadoTRF = Nothing

	   On Error Resume Next

	   sSQL = "SELECT	* " & _
			  "FROM estado " & _
			  "WHERE	codigo_estado in (1, 9) " & _
			  "and		nombre_campo = 'transferencia' " &  _
			  "ORDER BY nombre_estado"
	   	   
	   Set rsEst = EjecutarSQLCliente(Session("afxCnxAFEXchange"), sSQL)

	   If Err.Number <> 0 Then 
			Set rsEst = Nothing
			MostrarErrorMS "Obtener Estado Transfer 1"
		End If
	   
	   Set rsEst.ActiveConnection = Nothing
	   Set ObtenerEstadoTRF = rsEst

	   Set rsEst = Nothing
	End Function

	Function ObtenerBanco()
	   Dim rsATC
	   Dim sSQL

	   Set ObtenerBanco = Nothing

	   On Error Resume Next

	   sSQL = "SELECT	* " & _
				 "FROM Banco " & _
				 "WHERE	pais_banco='CHILE' " & _
				 "ORDER BY nombre_banco "
				 
	   	   
	   Set rsATC = EjecutarSQLCliente(Session("afxCnxAFEXchange"), sSQL)

	   If Err.Number <> 0 Then 
			Set rsATC = Nothing
			MostrarErrorMS "Obtener Banco 1"
		End If
	   
	   Set rsATC.ActiveConnection = Nothing
	   Set ObtenerBanco = rsATC

	   Set rsATC = Nothing
	End Function

	Sub CargarSucursal(ByRef Seleccionado)
		Dim sSelect
		Dim rs
		Dim sConexion
		Dim sMayMin
		Dim sCodigo
		
		On Error Resume Next
		Set rs = Server.CreateObject("ADODB.Recordset")
		If Err.number <> 0 Then response.Redirect "Error.asp?Titulo=Error en Cargar Banco&Number=" & Err.number   & "&Source=" & Err.Source & "&Description=" & Err.description		
		sSelect = Empty

		Set rs = ObtenerSucursal()
		sCodigo = Seleccionado
		If Err.number <> 0 Then 
			'response.Redirect "../Compartido/Error.asp?Titulo=Errores en HágaseCliente&Number=" & Err.number   & "&Source=" & Err.Source & "&Description=" & Err.description
			MostrarErrorMS "Cargar Sucursal 1"
		End If
				
		'Response.Write "<option value=>NINGUNA</option>"
		Response.Write "<option value=XX>TODAS</option>"
		
		If  Not rs.EOF Then
			Do Until rs.eof	
				If UCASE(Trim(rs("codigo_agente"))) = Ucase(Trim(sCodigo)) Then 					
					sSelect = "SELECTED"				
				Else
					sSelect = ""
				End If
				sMayMin = trim(rs("nombre"))
				sMayMin = MayMin(sMayMin) 
				Response.write "<option " & sSelect & " value=" & _
				trim(rs("codigo_agente")) & ">" & _
				sMayMin & _
				" </option> "
				If err.number <> 0 Then 
					MostrarErrorMS "Cargar Sucursal 3"
				End If																
				rs.MoveNext
			Loop
		End If
			
	End Sub

	Function ObtenerSucursal()
	   Dim rsATC
	   Dim sSQL

	   Set ObtenerSucursal = Nothing

	   On Error Resume Next

	   sSQL = "SELECT	* " & _
				 "FROM Cliente " & _
				 "WHERE	tipo = 4 " & _
				 "ORDER BY nombre "
				 
	   	   
	   Set rsATC = EjecutarSQLCliente(Session("afxCnxCorporativa"), sSQL)

	   If Err.Number <> 0 Then 
			Set rsATC = Nothing
			MostrarErrorMS "Obtener Sucursal 1"
		End If
	   
	   Set rsATC.ActiveConnection = Nothing
	   Set ObtenerSucursal = rsATC

	   Set rsATC = Nothing
	End Function

	Function CargarEjecutivos(ByVal sCodigoSucursal, ByVal nCodigoEjecutivo)
		Dim rsEjecutivo
		Dim sSQL
		Dim sSelect
		
		CargarEjecutivos = False
		
		sSQL = "Select " & _
					"nombre_completo as NombreEjecutivo, " & _
					"codigo_empleado as CodigoEjecutivo " & _
			   "From " & _
			   "	VEmpleado " & _
			   "Where " & _
			   "	codigo_agente = '" & sCodigoSucursal & "'"
			   
		Set rsEjecutivo = EjecutarSQLCliente(Session("afxCnxCorporativa"), sSQL)
		
		If Err.number <> 0 Then
			Set rsEjecutivo = Nothing
			Response.Redirect "http:../Compartido/Error.asp?Titulo=Agregar Cliente&Number=" & Err.number   & "&Source=" & Err.Source & "&Description=" & Err.Description
			Exit Function
		End If
	
		Do Until rsEjecutivo.EOF
			If Trim(rsEjecutivo("CodigoEjecutivo")) = Trim(nCodigoEjecutivo) Then 					
				sSelect = "SELECTED"				
			Else
				sSelect = ""
			End If

			Response.write "<option " & sSelect & " value=" & _
				Trim(rsEjecutivo("CodigoEjecutivo")) & ">" & _
				Trim(rsEjecutivo("NombreEjecutivo")) & "</option> "

			rsEjecutivo.MoveNext
		Loop
		
		Set rsEjecutivo = Nothing
		CargarEjecutivos = True
	End Function

	Sub CargarRubro(ByRef Seleccionado)
		Dim sSelect
		Dim rs
		Dim sConexion
		Dim sMayMin
		Dim sCodigo
		
		On Error Resume Next
		Set rs = Server.CreateObject("ADODB.Recordset")
		If Err.number <> 0 Then response.Redirect "Error.asp?Titulo=Error en Cargar Rubro&Number=" & Err.number   & "&Source=" & Err.Source & "&Description=" & Err.description		
		sSelect = Empty

		Set rs = ObtenerRubro()
		sCodigo = Seleccionado
		If Err.number <> 0 Then 
			'response.Redirect "../Compartido/Error.asp?Titulo=Errores en HágaseCliente&Number=" & Err.number   & "&Source=" & Err.Source & "&Description=" & Err.description
			MostrarErrorMS "Cargar Rubro 1"
		End If
				
		Response.Write "<option value=>NINGUNO</option>"
		If  Not rs.EOF Then
			Do Until rs.eof	
				If UCASE(Trim(rs("correlativo"))) = Ucase(Trim(sCodigo)) Then 					
					sSelect = "SELECTED"				
				Else
					sSelect = ""
				End If
				sMayMin = trim(rs("nombre"))
				sMayMin = MayMin(sMayMin) 
				Response.write "<option " & sSelect & " value=" & _
				trim(rs("correlativo")) & ">" & _
				sMayMin & _
				" </option> "
				If err.number <> 0 Then 
					MostrarErrorMS "Cargar Rubro 3"
				End If																
				rs.MoveNext
			Loop
		End If
	End Sub

	Function ObtenerRubro()
	   Dim rsRubro
	   Dim sSQL

	   Set ObtenerRubro = Nothing

	   On Error Resume Next

	   sSQL = "SELECT	* " & _
				 "FROM Rubro " & _
				 "ORDER BY nombre"
	   	   
	   Set rsRubro = EjecutarSQLCliente(Session("afxCnxCorporativa"), sSQL)

	   If Err.Number <> 0 Then 
			Set rsRubro = Nothing
			MostrarErrorMS "Obtener Rubro 1"
		End If
	   
	   Set rsRubro.ActiveConnection = Nothing
	   Set ObtenerRubro = rsRubro

	   Set rsRubro = Nothing
	End Function

	Function CargarEstado(ByVal Tipo, ByVal Seleccionado)
		Dim rs
		Dim sSQL
		Dim sSelect
		
		CargarEstado = False
		
		sSQL = "Select " & _
					"nombre_tipo as Nombre, " & _
					"codigo_tipo as Codigo " & _
			   "From " & _
			   "	estado " & _
			   "Where " & _
			   "	nombre_campo = '" & Tipo & "'"
			   
		Set rs = EjecutarSQLCliente(Session("afxCnxCorporativa"), sSQL)
		
		If Err.number <> 0 Then
			Set rs = Nothing
			Response.Redirect "http:../Compartido/Error.asp?Titulo=Agregar Cliente&Number=" & Err.number   & "&Source=" & Err.Source & "&Description=" & Err.Description
			Exit Function
		End If
	
		Do Until rs.EOF
			If Trim(rs("Codigo")) = Trim(Seleccionado) Then 					
				sSelect = "SELECTED"
			Else
				sSelect = ""
			End If

			Response.write "<option " & sSelect & " value=" & _
				Trim(rs("Codigo")) & ">" & _
				ucase(Trim(rs("Nombre"))) & "</option> "

			rs.MoveNext
		Loop
		
		Set rs = Nothing
		CargarEstado = True
	End Function

	function CargarPerfil(Byval Perfil, ByVal Seleccionado)
		Dim sSelect
		Dim rs
		Dim sSQL
		Dim sMayMin
		Dim sCodigo
		
		On Error Resume Next
		CargarPerfil = false

		sCodigo = Seleccionado	
		
		' consulta el detalle del perfil solicitado
		sSQL = " select * from detalleperfil where codigoperfil = " & perfil
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open sSQL, Session("afxCnxCorporativa")

		If Err.number <> 0 Then 
			MostrarErrorMS "Cargar detalle perfil 2." & Perfil & err.description
		End If				

		Response.write "<option value=0></option> "
		If Not rs.EOF Then
			Do Until rs.eof	
				If UCASE(Trim(rs("correlativo"))) = Ucase(Trim(sCodigo)) Then 					
					sSelect = "SELECTED"
				Else
					sSelect = ""
				End If
				sMayMin = trim(rs("descripcion"))
				sMayMin = MayMin(sMayMin) 
				Response.write "<option " & sSelect & " value=" & _
				trim(rs("correlativo")) & ">" & _
				sMayMin & _
				" </option> "
				If err.number <> 0 Then 
					MostrarErrorMS "Cargar perfiles 3"
				End If																
				rs.MoveNext
			Loop
		End If
		CargarPerfil = True
	end function

		Function CargarTipo(ByVal Tipo, ByVal Seleccionado)
		Dim rs
		Dim sSQL
		Dim sSelect
		
		CargarTipo = False
		
		sSQL = "Select " & _
					"nombre_tipo as Nombre, " & _
					"codigo_tipo as Codigo " & _
			   "From " & _
			   "	tipo " & _
			   "Where " & _
			   "	nombre_campo = '" & Tipo & "'"
			 'Response.Write SSQL
			   
		Set rs = EjecutarSQLCliente(Session("afxCnxCorporativa"), sSQL)
		
		If Err.number <> 0 Then
			Set rs = Nothing
			Response.Redirect "http:../Compartido/Error.asp?Titulo=Agregar Cliente&Number=" & Err.number   & "&Source=" & Err.Source & "&Description=" & Err.Description
			Exit Function
		End If
	
		Do Until rs.EOF
			If Trim(rs("Codigo")) = Trim(Seleccionado) Then 					
				sSelect = "SELECTED"
			Else
				sSelect = ""
			End If

			Response.write "<option " & sSelect & " value=" & _
				Trim(rs("Codigo")) & ">" & _
				ucase(Trim(rs("Nombre"))) & "</option> "

			rs.MoveNext
		Loop
		
		Set rs = Nothing
		CargarTipo = True
	End Function
	
	
	
	Sub CargarComboOpcionAutorizacion()
		Dim rs
		Dim sSQL		
		
		sSQL = " exec Cumplimiento.MostrarOpcionAutorizacion '" & Session("NombreUsuarioOperador") & "' "
			   
		Set rs = EjecutarSQLCliente(Session("afxCnxCorporativa"), sSQL)
		
		If Err.number <> 0 Then
			Set rs = Nothing			
			Response.write "<option selected value=0>Problemas al cargar...</option> "
		End If		
	
		if not rs.eof then		
			Response.write "<option selected value=0></option> "
			Do Until rs.EOF
				Response.write "<option value=" & _
					Trim(rs("Codigo")) & ">" & _
					ucase(Trim(rs("descripcionautorizacion"))) & "</option> "

				rs.MoveNext
			Loop
			rs.close
		end if
		
		Set rs = Nothing		
	End Sub
	
	
	Sub CargarComboValorOpcionAutorizacion()
		Dim rs
		Dim sSQL		
		
		sSQL = " exec Cumplimiento.MostrarOpcionAutorizacion '" & Session("NombreUsuarioOperador") & "' "
			   
		Set rs = EjecutarSQLCliente(Session("afxCnxCorporativa"), sSQL)
		
		If Err.number <> 0 Then
			Set rs = Nothing			
			Response.write "<option selected value=0>Problemas al cargar...</option> "
		End If		
	
		if not rs.eof then		
			Response.write "<option selected value=0></option> "
			Do Until rs.EOF
				Response.write "<option value=" & _
					Trim(rs("Codigo")) & ">" & _
					ucase(Trim(rs("valorhistoria"))) & "</option> "

				rs.MoveNext
			Loop
			rs.close
		end if
		
		Set rs = Nothing		
	End Sub
	
	Function CargarActividadEconomica(Byval Seleccionado)
		Dim sSelect
		Dim rs
		Dim sSQL
		Dim sMayMin
		Dim sCodigo
		
		On Error Resume Next
		CargarActividadEconomica = false

		sCodigo = Seleccionado	
		
		sSQL = " exec mostrarlistaactividadeconomica "
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open sSQL, Session("afxCnxCorporativa")

		If Err.number <> 0 Then 
			MostrarErrorMS "Cargar actividad economica 2. " & err.description
		End If				

		Response.write "<option value=0></option> "
		If Not rs.EOF Then
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
					MostrarErrorMS "Cargar actividad economica 3. " & err.Description 
				End If																
				rs.MoveNext
			Loop
		End If
		CargarActividadEconomica = True
	end function

	
'Envía correo desde el servidor BOLDO utilizando la aplicación Mensajería Afex. Se llama el SP Envio_Mail y se pasa como parámetros, 
'los correos de los destinatarios extraidos con el SP Mail_Listar_Destinatarios, el asunto y el cuerpo en formato HTML. 
	Sub EnviarEMailBD(ByVal IdAplicacion, ByVal IdMail, ByVal IdAmbiente, ByVal Asunto, ByVal Cuerpo)
		Dim rsListaCorreo, sSQLDestinatarios, sSQLEnvioMail, rs
		'Se buscan los destinatarios guardados al asociar el correo creado en la aplicación Mensajería Afex
		sSQLDestinatarios = "exec Mail_Listar_Destinatarios " & IdAplicacion & ", " & IdMail & ", " &  IdAmbiente
		set rsListaCorreo = EjecutarSQLCliente(session("afxCnxServidorCorreo"), sSQLDestinatarios)
		
		If Err.number <> 0 Then				
			MostrarErrorMS "Búsqueda destinatarios - " & err.Description 
		End If
		
		If not rsListaCorreo is nothing then
			If not rsListaCorreo.EOF then
			 'MostrarErrorMS "Búsqueda destinatarios - Sin Destinatarios"
			
			'Envía el correo a cada uno de los destinatarios encontrados
			  Do While not rsListaCorreo.EOF
			    Cuerpo = replace(Cuerpo,"[nombreDestinatario]",rsListaCorreo("nombre"))
			    sSQLEnvioMail = "exec Envio_Mail " & evaluarstr(rsListaCorreo("email")) & ",'', '" & _
			    Asunto & "', '" & Cuerpo & "' "
			    set rs = EjecutarSQLCliente(session("afxCnxServidorCorreo"), sSQLEnvioMail)
			    rsListaCorreo.MoveNext
			  Loop 
			End If			
		End If
		
		SET rsListaCorreo = Nothing			   
		SET rs = nothing

	end sub

    'INTERNO-9263 MS 19-01-2017
	Function CargarOcupacion(Byval Seleccionado)
		Dim sSelect
		Dim rs
		Dim sSQL
		Dim sMayMin
		Dim sCodigo
		
		On Error Resume Next
		CargarOcupacion = false

		sCodigo = Seleccionado	
		
		sSQL = "exec dbo.MostrarOcupaciones "
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open sSQL, Session("afxCnxCorporativa")

		If Err.number <> 0 Then 
			MostrarErrorMS "Cargar ocupaciones. " & err.description
		End If				

		Response.write "<option value=0></option> "
		If Not rs.EOF Then
			Do Until rs.eof	
				If rs("IdOcupacion") = sCodigo Then 					
					sSelect = "SELECTED"
				Else
					sSelect = ""
				End If
				sMayMin = trim(rs("DescripcionOcupacion"))
				sMayMin = MayMin(sMayMin) 
				Response.write "<option " & sSelect & " value=" & _
				rs("IdOcupacion") & ">" & _
				sMayMin & _
				" </option> "
				If err.number <> 0 Then 
					MostrarErrorMS "Cargar ocupaciones 2. " & err.Description 
				End If																
				rs.MoveNext
			Loop
		End If
		CargarOcupacion = True
	end function
	'FIN INTERNO-9263 MS 19-01-2017
	
%>
