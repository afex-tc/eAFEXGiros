<!-- Transferencias.asp -->
<%
	'Funciones y procedimientos para transferencias
	
	Function ValidarVigencia()
		Dim rsExpira, dFecha, sHora
		
		Set rsExpira = ObtenerVigencia()
		If Not rsExpira.EOF Then
			dFecha = CDate(rsExpira("nombre_tipo"))
			rsExpira.MoveNext 
			sHora = rsExpira("nombre_tipo")

			If (CDate(Date()) = dFecha AND Time() > sHora) OR _
			CDate(Date()) > dFecha Then
				ValidarVigencia = "El plazo de las paridades ha expirado<BR> Sólo puede operar con Dólar Americano"
			End If
		Else
			ValidarVigencia = "No se encontraron paridades vigentes"
		End If
		
		Set rsExpira = Nothing
	End Function

	Function ObtenerVigencia()
		Dim afx, rsVigencia
				
		Set afx = Server.CreateObject("AfexWebXP.Web")
		Set rsVigencia = afx.ObtenerVigenciaParidad(Session("afxCnxAFEXchange"))
		If Err.number <> 0 Then
			Set rsVigencia = Nothing
			Set afx = Nothing
			MostrarErrorMS "Enviar Transferencia 4"
		End If
		If afx.ErrNumber <> 0 Then			
			Set rsVigencia = Nothing
			MostrarErrorAFEX afx, "Enviar Transferencia 5"
		End If
		
		Set afx = Nothing
		Set ObtenerVigencia = rsVigencia
		Set rsMoneda = Nothing
	
	End Function
	
	Function ObtenerMonedasTransfer()
		Dim afxTransfer1
		Dim rsMoneda1
				
		Set afxTransfer1 = Server.CreateObject("AfexWebXP.Web")
		Set rsMoneda1 = afxTransfer1.ObtenerMonedasTransfer(Session("afxCnxAFEXchange"))		
		If Err.number <> 0 Then
			Set rsMoneda1 = Nothing
			Set afxTransfer1 = Nothing
			MostrarErrorMS "Obtener Monedas de Transferencia 1"
		End If
		If afxTransfer1.ErrNumber <> 0 Then			
			Set rsMoneda1 = Nothing
			MostrarErrorAFEX afxTransfer1, "Obtener Monedas  de Transferencia 2"
		End If
		
		Set afxTransfer1 = Nothing
		Set ObtenerMonedasTransfer = rsMoneda1
		Set rsMoneda1 = Nothing
		
	End Function	

	Sub CargarMonedasTransfer(ByVal Seleccionado)
		Dim rs, sMayMin, sSelect

				
		Set rs = ObtenerMonedasTransfer()
		
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

	Sub CargarParidadesTransfer(ByVal Tipo, ByVal Seleccionado)
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
		
		Set rsP = ObtenerMonedasTransfer()
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

	Sub CargarValuta(ByVal Seleccionada)
		Dim sSelect
		
		If Seleccionada="1" Then sSelect = "selected"
		Response.write "<option " & sSelect & " value=48>Normal 48 hrs</option>"
		sSelect = ""
		' Jonathan Miranda G. 06-10-2006		
		'If Seleccionada="2" Then sSelect = "selected"		
		'Response.write "<option " & sSelect & " value=24>Express 24 hrs</option>"
		'-------------------------- Fin ------------------------		
		sSelect = ""
		If Seleccionada="3" Then sSelect = "selected"		
		Response.write "<option " & sSelect & " value=0>Express</option>"		' Premium 0 hrs
	End Sub	

	Sub CargarFormaPago(ByVal Seleccionada)
		Dim sSelect
		
		If Seleccionada="1" Then sSelect = "selected"
		Response.write "<option " & sSelect & " value=1>Efectivo en USD</option>"
		sSelect = ""
		If Seleccionada="2" Then sSelect = "selected"
		Response.write "<option " & sSelect & " value=2>Efectivo en Pesos</option>"
		sSelect = ""
		If Seleccionada="3" Then sSelect = "selected"		
		Response.write "<option " & sSelect & " value=3>Depósito en USD</option>"		
		sSelect = ""
		If Seleccionada="4" Then sSelect = "selected"		
		Response.write "<option " & sSelect & " value=4>Depósito en Pesos</option>"		
		sSelect = ""
		If Seleccionada="5" Then sSelect = "selected"		
		Response.write "<option " & sSelect & " value=5>Valores en custodia USD</option>"		
	End Sub	

	Sub CargarParidadesTransferCargo(ByVal Tipo, ByVal Seleccionado, ByVal Cargo)
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
		
		Set rsP = ObtenerMonedasTransfer()
		If  Not rsP.EOF Then		
			Do Until rsP.eof	
				If UCASE(Trim(rsP("codigo_moneda"))) = Ucase(Trim(Seleccionado)) Then 					
					sSelect = "SELECTED"				
				Else
					sSelect = ""
				End If
				Response.write "<option " & sSelect & " value=" & _
				trim(rsP("codigo_moneda")) & "> " & (cCur(0 & rsP(sTipo)) + ROUND((cCur(0 & rsP(sTipo)) * cCur(0 & Cargo)) / 100, 7)) & " </option> "
				If Err.number <> 0 Then
					Set rsP = Nothing
					MostrarErrorMS "Cargar Paridades 1"
				End If
				rsP.MoveNext
			Loop
		End If
		Set rsP = Nothing	
	End Sub
	
	sub CargarMonedasMantenedor()
		dim rs
		dim sSQL		
		
		sSQL = " execute MostrarMonedasMantenedor "
		set rs = EjecutarSQLCliente(Session("afxCnxAfexchange"), sSQL)
		if err.number <> 0 then
			set rs = nothing
			mostrarerrorms "Cargar Monedas Mantenedor"
		end if
				
		Response.Write "<option value=></option>"
		If  Not rs.EOF Then
			Do Until rs.eof					
				sMayMin = trim(rs("nombre"))
				sMayMin = MayMin(sMayMin) 
				Response.write "<option " & sSelect & " value=" & _
				trim(rs("codigo")) & ">" & _
				sMayMin & _
				" </option> "
				If err.number <> 0 Then 
					'response.Redirect "../Compartido/Error.asp?Titulo=Error en HágaseCliente&Number=" & Err.Number  & "&Source=" & Err.source  & "&Description=" & Err.description	
					MostrarErrorMS "Cargar Monedas Mantenedor"
				End If																
				rs.MoveNext
			Loop
			rs.close
		End If
		
		set rs = nothing
	end sub
	
	Sub CargarParidadesMantenedor()
		Dim rs, sMayMin, sSelect, sTipo

		sSQL = " execute MostrarParidadesMantenedor "
		set rs = EjecutarSQLCliente(Session("afxCnxAfexchange"), sSQL)
		if err.number <> 0 then
			set rs = nothing
			mostrarerrorms "Cargar Paridades Monedas Mantenedor"
		end if	
		
		If Not rs.EOF Then
			Response.Write "<option value=></option>"
			Do Until rs.eof				
				Response.write "<option  value=" & _
				trim(rs("codigo")) & "> " & rs("paridad") & " </option> "
				If Err.number <> 0 Then
					Set rs = Nothing
					MostrarErrorMS "Cargar Paridades Monedas Mantenedor 1"
				End If
				rs.MoveNext
			Loop
			rs.close
		End If
		Set rs = Nothing	
	End Sub
	
	sub CargarMonedasCheques()
		dim rs
		dim sSQL		
		
		sSQL = " execute MostrarMonedasMantenedor 1 "
		set rs = EjecutarSQLCliente(Session("afxCnxAfexchange"), sSQL)
		if err.number <> 0 then
			set rs = nothing
			mostrarerrorms "Cargar Monedas Mantenedor"
		end if
				
		Response.Write "<option value=></option>"
		If  Not rs.EOF Then
			Do Until rs.eof					
				sMayMin = trim(rs("nombre"))
				sMayMin = MayMin(sMayMin) 
				Response.write "<option " & sSelect & " value=" & _
				trim(rs("codigo")) & ">" & _
				sMayMin & _
				" </option> "
				If err.number <> 0 Then 
					'response.Redirect "../Compartido/Error.asp?Titulo=Error en HágaseCliente&Number=" & Err.Number  & "&Source=" & Err.source  & "&Description=" & Err.description	
					MostrarErrorMS "Cargar Monedas Cheque 3"
				End If																
				rs.MoveNext
			Loop
			rs.close
		End If
		
		set rs = nothing
	end sub
	
	Sub CargarParidadesTransfer2(ByVal Tipo, ByVal Seleccionado)
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
		
		Set rsP = ObtenerMonedasTransfer()
		If  Not rsP.EOF Then	
			Response.Write "<option value=></option>"
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
	
	Sub CargarPaisTransfer(ByVal Seleccionado)
		Dim rs, sSQL, sMayMin, sSelect

		sSQL = "select nombre_pais, codigo_pais from pais order by nombre_pais"
		Set rs = ejecutarsqlcliente(session("afxCnxAFEXchange"), sSQL)
		
		If Err.number <> 0 Then
			Set rs = Nothing
			MostrarErrorMS "Cargar País Transferencias"
		End If
		
		If  Not rs.EOF Then
			Response.write "<option value=''></option> "
			Do Until rs.eof	
				If UCASE(Trim(rs("codigo_pais"))) = Ucase(Trim(Seleccionado)) Then
					sSelect = "SELECTED"
				Else
					sSelect = ""
				End If
				sMayMin = trim(rs("nombre_pais"))
				sMayMin = MayMin(sMayMin) 
				Response.write "<option " & sSelect & " value=" & _
				trim(rs("codigo_pais")) & ">" & _
				sMayMin & _
				" (" & UCase(trim(rs("codigo_pais"))) & ")" & _
				" </option> "
				If Err.number <> 0 Then
					Set rs = Nothing
					MostrarErrorMS "Cargar País Transferencias 1"
				End If
				rs.MoveNext
			Loop
		End If
		Set rs = Nothing
	End Sub	
	
%>
