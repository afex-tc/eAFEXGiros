<%@ Language=VBScript %>
<!--#INCLUDE virtual="/Compartido/TimeOut.asp" -->
<!--#INCLUDE virtual="/Compartido/Constantes.asp" -->
<%

	'Variables
	Dim rsGiro, sCodigo, afxGiro, nTipoLlamada, nTipoLista
	Dim sAFEXchange, sAFEXpress, nEstado
	Dim nDecimales, sColorMoneda
	Dim sNumeroPIN, Correlativo, sCodigoCaptador, sCodigoPagador
	Dim nCategoria, sPaisBeneficiario, sCiudadBeneficiario
	Dim cMontoGiro, dFechaGiro
	
	sCodigo = Request("Codigo")	
	nTipoLlamada = cInt(0 & Request("TipoLlamada"))
	nTipoLista = cInt(0 & Request("TipoLista"))
	sAFEXchange = Request("AFEXchange")
	sAFEXpress = Request("Cliente")
	nCategoria = Session("Categoria")
	
	sNumeroPIN = ""
	Correlativo = 0
		
	If Session("ModoPrueba") Then	
		Set rsGiro = ObtenerGiroXP(sCodigo)
	Else
		Set rsGiro = ObtenerGiroXP(sCodigo)
	End If
	
	dFechaGiro = EvaluarVar(rsGiro("fecha_captacion"), "")
	If dFechaGiro = Empty Then
		dFechaGiro = Date
	End If
	
	sCodigoCaptador = EvaluarVar(rsGiro("agente_captador"), "")
	sCodigoPagador = EvaluarVar(rsGiro("agente_pagador"), "")
	sPaisBeneficiario = EvaluarVar(rsGiro("pais_beneficiario"), "")
	sCiudadBeneficiario = EvaluarVar(rsGiro("ciudad_beneficiario"), "")
	cMontoGiro = cCur(0 & rsGiro("monto_giro"))
	TipoGiro = cInt(0 & rsGiro("tipo_giro"))
	
	'response.write sCodigoPagador
	'response.end
	
	If sAFEXpress = ""  And EvaluarVar(rsGiro("codigo_beneficiario"), "") <> "" Then
		CargarCliente
	End If

	'Response.write cInt(0 & Request("Accion")) & ", " & afxAccionPagar & ", " & afxAccionPagarTercero
	'Set afxGiro = Server.CreateObject("AfexGiro.Giro")
	'response.Redirect "../compartido/error.asp?description=" & sCodigo & ", " & rsgiro("nombre_retira")

	'Set afxGiro = CargarGiro(sCodigo)
	nEstado = cInt(rsGiro("estado_giro"))
	If rsGiro("codigo_moneda") = Session("MonedaNacional") Then
		nDecimales = 0
		sColorMoneda = "DodgerBlue"
	Else
		nDecimales = 2
		sColorMoneda = "#4dc087" '"MediumSeaGreen"
	End If
	On Error Resume Next


	Sub CargarCliente()
		Dim rsCliente, nCampo

		nCampo = afxCampoCodigoExpress
		sArgumento = rsGiro("codigo_beneficiario")
		sArgumento2 = ""
		sArgumento3 = ""
		If nCampo = 0 Then Exit Sub
					
		Set rsCliente = BuscarCliente(nCampo, sArgumento, sArgumento2, sArgumento3)
		'Response.Redirect "../Compartido/Error.asp?Titulo=" & rs.eof & ", " & ncampo & ", " &  sargumento  & ", " & sargumento2
		If Err.number <> 0 Then
			Set rsCliente = Nothing
			MostrarErrorMS "Cargar Cliente 1"
		End If
		If rsCliente.EOF Then
			rsCliente.Close
			Set rsCliente = Nothing
			Exit Sub
		End If
		If Not rsCliente.EOF Then
			Session("ATCAFEXpress") = rsCliente("Express")
			Session("ATCAFEXchange") = rsCliente("Exchange")
			sAFEXpress = rsCliente("Express")
			sAFEXchange = rsCliente("Exchange")
			If Trim(rsCliente("rut") & rsCliente("pasaporte")) <> "" Then
				Session("IdCliente") = 1
			Else
				Session("IdCliente") = 0
			End If
			If Err.number <> 0 Then
				rsCliente.Close
				Set rsCliente = Nothing
				MostrarErrorMS "Cargar Cliente 2"
			End If
		End If
		Set rsCliente = Nothing
	End Sub
	
	Function ActualizarPIN(ByVal nCorrelativo, ByVal CodigoGiro)
		Dim cnn, Sql
		dim fecha
		fecha= Date()
		On Error Resume Next
	
		ActualizarPIN = False
		
		'If (Session("CodigoAgente") = "AM" And sCodigoCaptador = "AM") Or (Session("CodigoAgente") = "AX" And sCodigoCaptador = "AX") Or  _
		'   (Session("CodigoAgente") = "AG" And sCodigoCaptador = "AG") Or (Session("CodigoAgente") = "ZC" And sCodigoCaptador = "ZC") Then

		If cDate(dFechaGiro) >= cDate("08/05/2006") Then
			'If (cMontoGiro >= 100 ) Then
				If UCase(Trim(sPaisBeneficiario)) = "PE" Then
					If (cInt(0 & nCategoria) = 1 Or cInt(0 & nCategoria) = 2) And (Session("CodigoAgente") = sCodigoCaptador)  Then
						Set cnn = CreateObject("ADODB.Connection") 
						'rs.Open Sql, Session("afxCnxAFEXpress"), 3, 1
						cnn.Open Session("afxCnxAFEXpress")
		
						If Err.number <> 0 Then
							Set cnn = Nothing
							MsgBox Err.Description
							Exit Function
						End If
		
						Sql = "Update tarjeta Set codigo_giro = '" & CodigoGiro & "', fecha = '" & fecha & "',codigo_agente= '" & Session("CodigoAgente") & "', codigo_usuario='"& Session("NombreUsuarioOperador") &"' " & _
							  "Where correlativo = " & nCorrelativo & " " & _
							  "and   tipo_PIN = 1"
		
						cnn.Execute Sql
	
						If Err.number <> 0 Then
							Set cnn = Nothing
							MsgBox Err.Description
							Exit Function
						End If
						Set cnn = Nothing
					End If
				End If
			'ElseIf cMontoGiro >= 150 And Ucase(Trim(sCodigoPagador)) = Trim(Session("CodigoMGEnvio")) Then
			'	Set cnn = CreateObject("ADODB.Connection") 
			'	'rs.Open Sql, Session("afxCnxAFEXpress"), 3, 1
			'	cnn.Open Session("afxCnxAFEXpress")
		
			'	If Err.number <> 0 Then
			'		Set cnn = Nothing
			'		MsgBox Err.Description
			'		Exit Function
			'	End If
		
			'	Sql = "Update tarjeta Set codigo_giro = '" & CodigoGiro & "' " & _
			'		  "Where correlativo = " & nCorrelativo & " " & _
			'		  "and   tipo_PIN = 3"
		
			'	cnn.Execute Sql
	
			'	If Err.number <> 0 Then
			'		Set cnn = Nothing
			'		MsgBox Err.Description
			'		Exit Function
			'	End If
		
			'	Set cnn = Nothing
			'End If
		End If
		ActualizarPIN = True
	End Function

	Function BuscarPIN(ByVal CodigoGiro)
		Dim rs, Sql
		
		On Error Resume Next
		
		BuscarPIN = False
		
		If cDate(dFechaGiro) >= cDate("15/12/2007")  Then

			'If cMontoGiro >= 100  Then
				If UCase(Trim(sPaisBeneficiario)) = "PE" Then
					If (cInt(0 & nCategoria) = 1 Or cInt(0 & nCategoria) = 2) And (Session("CodigoAgente") = sCodigoCaptador) Then
						Sql = "select correlativo, " & _
							  "numero_pin from tarjeta " & _
							  "where codigo_giro = '" & CodigoGiro & "'" & _
							  "and   tipo_PIN = 1"
			'MostrarErrorMS sql
			 
						Set rs = CreateObject("ADODB.Recordset") 
						'Response.Write sql
						'Response.End 
						rs.Open Sql, Session("afxCnxAFEXpress"), 3, 1
						
						If Err.number <> 0 Then
							Set rs = Nothing
							MostrarErrorMS "BuscarPIN"
							'MsgBox "Se produjo un error al buscar Nº del PIN"
							Exit Function
						End If
		
						If rs.EOF Then
							sNumeroPIN = ""
							
							Set rs = Nothing
							Exit Function
						End If
		
						sNumeroPIN = rs("numero_pin")
		
						Set rs = Nothing
					End If
				End If
			'ElseIf cMontoGiro >= 150 And Ucase(Trim(sCodigoPagador)) = Trim(Session("CodigoMGEnvio")) Then
			'	Sql = "select correlativo, " & _
			'		  "numero_pin from tarjeta " & _
			'		  "where codigo_giro = '" & CodigoGiro & "' " & _
			'		  "and   tipo_PIN = 3"
		
			'	Set rs = CreateObject("ADODB.Recordset") 
			'			
			'	rs.Open Sql, Session("afxCnxAFEXpress"), 3, 1
			'			
			'	If Err.number <> 0 Then
			'		Set rs = Nothing
			'		MostrarErrorMS "BuscarPIN"
			'		'MsgBox "Se produjo un error al buscar Nº del PIN"
			'		Exit Function
			'	End If
		
			'	If rs.EOF Then
			'		sNumeroPIN = ""
			'				
			'		Set rs = Nothing
			'		Exit Function
			'	End If
		
			'	sNumeroPIN = rs("numero_pin")
		
			'	Set rs = Nothing
			'End If
		End If
	
		BuscarPIN = True
		
	End Function

	Function ObtenerDatosPIN()
		Dim rs, Sql
		
		On Error Resume Next
		
		'tipo_PIN = 1 Promoción; 2 Venta Tarjeta Telefónica
		
		ObtenerDatosPIN = False		
		
		If Ucase(Trim(sCodigoPagador)) <> Trim(Session("CodigoMGEnvio")) Then
			Sql = "select top 1 min(correlativo) as correlativo, " & _
				  "numero_pin from tarjeta " & _
				  "where codigo_giro is null and tipo_PIN = 1 " & _
				  "group by numero_pin " & _
				  "order by correlativo"
		'ElseIf Ucase(Trim(sCodigoPagador)) = Trim(Session("CodigoMGEnvio")) Then
		'	Sql = "select top 1 min(correlativo) as correlativo, " & _
		'		  "numero_pin from tarjeta " & _
		'		  "where codigo_giro is null and tipo_PIN = 3 " & _
		'		  "group by numero_pin " & _
		'		  "order by correlativo"
		End If
			
		Set rs = CreateObject("ADODB.Recordset") 
		
		'If Err.number <> 0 Then
		'	MsgBox Err.Description
		'	Exit Function
		'End If
		'MostrarErrorMS sql
		rs.Open Sql, Session("afxCnxAFEXpress"), 3, 1
		
		If Err.number <> 0 Then
			Set rs = Nothing
			MostrarErrorMS "ObtenerDatosPIN", Err.Description
			Exit Function
		End If
		
		If rs.EOF Then
			sNumeroPIN = ""
			Correlativo = 0
			
			Set rs = Nothing
			Exit Function
		End If
		
		sNumeroPIN = rs("numero_pin")
		Correlativo = rs("correlativo")
		
		ObtenerDatosPIN = True
		
		Set rs = Nothing
	End Function
	
	Function BuscarHistoria(ByVal Conexion, _
							ByVal Giro)
		Dim sSQL
	      
		'Manejo de errores
		On Error Resume Next
	   
		'Crea la consulta
		sSQL = "SELECT    * " & _
	           "FROM      VHistoria with(nolock) " & _
	           "WHERE     codigo_giro = '" & Giro & "' "
	   
		'Asigna al metodo el resultado de la consulta
		Set BuscarHistoria = EjecutarSQLCliente(Conexion, sSQL)
	   
		If Err.number <> 0 Then
			MostrarErrorMS "Obtener Historia"
		End If
	End Function

	Response.Expires = 0

%>
<!--#INCLUDE virtual="/Compartido/Rutinas.asp" -->
<!--#INCLUDE virtual="/Compartido/Giros.asp" -->
<!--#INCLUDE virtual="/Compartido/Errores.asp" -->
<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link rel="stylesheet" type="text/css" href="../Estilos/Cliente.css">
<link rel="stylesheet" type="text/css" href="../Estilos/Reportes.css">
</head>
<script LANGUAGE="VBScript">
<!--
	On Error Resume Next
	
	Sub window_onLoad()

	<%
		Select Case nTipoLLamada
		Case afxAgente
	%>
			CargarMenuAgente
	<%
		Case afxCliente
	%>
			CargarMenuCliente
	<%
		End Select
		Response.Expires = 0
	%>
		'window.offscreenBuffering = False
		'window.setInterval "BlinkMoneda", 700, "vbscript"
		
	<%
		
		 Select Case Request("Accion")
		 Case afxAccionPagar	' Pagar
		 
	%>
			objMenu_OnScriptletEvent("linkClick", "Pagar")
			
	<%			
		 Case afxAccionPagarTercero	' Pagar a Tercero
	%>
			objMenu_OnScriptletEvent("linkClick", "PagarTercero")
	
	<%	 
		 Case Else
		 End Select
	%>
	End Sub

	Sub BlinkMoneda
		If tdMoneda.style.display = "" Then
			tdNN.style.display = ""
			tdMoneda.style.display = "none"
		Else
			tdNN.style.display = "none"
			tdMoneda.style.display = ""
		End If
	End Sub
	
	Sub CargarMenuAgente()
		Dim sId, i, nCont	
		
		objmenu.bgColor = document.bgColor 
		objmenu.stylesheet = "../Estilos/Cliente.css"
		sId = objmenu.addparent("Opciones")
		nCont = 0
	
				
		<%	
					
		Select Case rsGiro("estado_giro")
			Case 1, 2, 4 		
				If rsGiro("agente_captador") = Session("CodigoAgente") _
				   And (Session("Categoria") = 1 Or Session("Categoria") = 2 Or Session("Categoria") = 3 ) _
				      Then	
		%>
					objMenu.addchild sId, "Boleta de Servicios", "Servicios", "Principal"
					objMenu.addchild sId, "Ver Boleta", "BS", "Principal"
					objMenu.addchild sId, "Oficinas de Pago", "OfPago", "Principal"
				<% if   rsGiro("estado_giro")<> 2 then%>
					objMenu.addchild sId, "Anular Giro", "Anular", "Principal"
					objMenu.addchild sId, "Corregir Datos", "Corregir", "Principal"
				<% end if %>
					nCont = nCont + 5
		<%	
				elseif rsGiro("agente_captador") = Session("CodigoAgente") _
				   And Session("Categoria") = 4 and Trim(sPaisBeneficiario) = "CL" Then
				%>
					objMenu.addchild sId, "Anular Giro", "Anular", "Principal"
					objMenu.addchild sId, "Corregir Datos", "Corregir", "Principal"
					nCont = nCont + 2 
					
			<%				
			End If 

				If Session("Categoria") = 1 Then 
					If (rsGiro("sw_editado") = 1 Or rsGiro("tipo_giro") = 0) And Trim(sPaisBeneficiario) = Trim(Session("PaisMatriz")) Then
					'If (rsGiro("agente_pagador") = Session("CodigoAgente") Or _
					' rsGiro("agente_pagador") = Session("CodigoMatriz")) And rsGiro("sw_editado") = 1 And _
					' rsGiro("agente_captador") <> Session("CodigoAgente") Then
				'(rsGiro("ciudad_beneficiario") = Session("CiudadMatriz") Or rsGiro("forma_pago") = 1 ) And _					 
		%>
					'msgbox "<%=rsGiro("ciudad_beneficiario")%>" & ", " & <%=rsGiro("agente_pagador")%> & ", " & <%=Session("CodigoAgente")%> & ", " & <%=Session("CodigoMatriz")%>
						objMenu.addchild sId, "Pagar", "Pagar", "Principal"
						objMenu.addchild sId, "Pagar a Tercero", "PagarTercero", "Principal"
						'objMenu.addchild sId, "Avisar", "Avisar", "Principal"
						'objMenu.addchild sId, "Reclamar", "Reclamar", "Principal"
						objMenu.addchild sId, "Actualizar Datos Cliente", "ADC", "Principal"
						nCont = nCont + 4
		<%	
					End If 
				ElseIf Session("Categoria") = 2 Then 
					If (rsGiro("sw_editado") = 1 Or rsGiro("tipo_giro") = 0) And Trim(sPaisBeneficiario) = Trim(Session("PaisMatriz")) Then
					'If ( _
					'		rsGiro("agente_pagador") = Session("CodigoAgente") _
					'		Or ( _
					'				InStr(1, "AV;AA;AQ;AB", Session("CodigoAgente")) <> 0 _
					'				And _
					'				InStr(1, "AV;AA;AQ;AB", rsGiro("agente_pagador")) <> 0 _									
					'			) _
					'		Or rsGiro=("tipo_giro") = 0 _
					'		Or rsGiro("agente_pagador") = Session("CodigoMatriz") _
					'	) _
					'	And rsGiro("sw_editado") = 1  _
					'	And rsGiro("agente_captador") <> Session("CodigoAgente") _
					'Then
					
'InStr(1, "KNA;VAP;QPE", rGiro("ciudad_beneficiario")) <> 0 					
		%>
						objMenu.addchild sId, "Pagar", "Pagar", "Principal"
						objMenu.addchild sId, "Pagar a Tercero", "PagarTercero", "Principal"
		<%
						If ( _
							rsGiro("agente_pagador") = Session("CodigoAgente") _
								Or ( _
										InStr(1, "AV;AA;AQ;AB", Session("CodigoAgente")) <> 0 _
										And _
										InStr(1, "AV;AA;AQ;AB", rsGiro("agente_pagador")) <> 0 _									
									) _
								Or rsGiro("tipo_giro") = 0 _
								Or rsGiro("agente_pagador") = Session("CodigoMatriz") _
							) _
							And rsGiro("sw_editado") = 1  _
							And rsGiro("agente_captador") <> Session("CodigoAgente") _
						Then
		%>		
						objMenu.addchild sId, "Avisar", "Avisar", "Principal"
		
		<%				End If	%>
						'objMenu.addchild sId, "Reclamar", "Reclamar", "Principal"
						objMenu.addchild sId, "Actualizar Datos Cliente", "ADC", "Principal"
						nCont = nCont + 5
		<%			
					End If 
				ElseIf Session("Categoria") = 3 Then
					If rsGiro("sw_editado") = 1 Or rsGiro("tipo_giro") = 0 Then
					'If (rsGiro("agente_pagador") = Session("CodigoAgente") And rsGiro("sw_editado")=1) Or rsGiro("tipo_giro")=0 Then
						'If rsGiro("estado_giro") <> 1 Or Session("CodigoAgente") = Session("CodigoMoneyBroker") Or rsGiro("tipo_giro")=0 And _
						'rsGiro("agente_captador") <> Session("CodigoAgente") Then
		%>
							objMenu.addchild sId, "Pagar", "Pagar", "Principal"
							objMenu.addchild sId, "Pagar a Tercero", "PagarTercero", "Principal"
							'objMenu.addchild sId, "Avisar", "Avisar", "Principal"
							'objMenu.addchild sId, "Reclamar", "Reclamar", "Principal"
							objMenu.addchild sId, "Actualizar Datos Cliente", "ADC", "Principal"
							nCont = nCont + 4
		<%	
						'End If
					End If
				ElseIf Session("Categoria") = 4 Then
					If rsGiro("agente_pagador") = Session("CodigoAgente") And rsGiro("sw_editado") = 1 Then
						If rsGiro("estado_giro") <> 1 Or Session("CodigoAgente") = Session("CodigoMoneyBroker") Then
							'If Session("ModoPrueba") Then
		%>
								objMenu.addchild sId, "Pagar", "PagarInt", "Principal"
								objMenu.addchild sId, "Avisar", "Avisar", "Principal"
								'objMenu.addchild sId, "Reclamar", "Reclamar", "Principal"
								nCont = nCont + 3
		<%	
							'End If
						End If
					End If
					
				End If 
								
				'If rsGiro("agente_captador") = Session("CodigoAgente") _
				'And (Session("Categoria") = 1 Or Session("Categoria") = 2 Or Session("Categoria") = 3) _
				'Then	
		%>
				'	objMenu.addchild sId, "Boleta de Servicios", "Servicios", "Principal"
				'	objMenu.addchild sId, "Ver Boleta", "BS", "Principal"
				'	nCont = nCont + 2
		<%	
				'End If 
			Case 5
				If rsGiro("agente_pagador") = Session("CodigoAgente") _
				And (Session("Categoria") = 1 Or  Session("Categoria") = 2 Or Session("Categoria") = 3) Then
		%>
					objMenu.addchild sId, "Comprobante de Pago", "Comprobante", "Principal"
					objMenu.addchild sId, "Ver Comprobante", "CP", "Principal"
					nCont = nCont + 2
		<%
				End If

			Case 7			
				If rsGiro("agente_captador") = Session("CodigoAgente") Then
		%>
					'objMenu.addchild sId, "Solucionar", "Solucionar", "Principal"
					'nCont = nCont + 1
		<%
				End If
				
'00000000000000000000000000000000000000000000000000000000000000000000000000000000000000000
'00000000000000000000000000000000000000000000000000000000000000000000000000000000000000000
'00000000000000000000000000000000000000000000000000000000000000000000000000000000000000000
				If uCase(Left(rsGiro("tipo_reclamo"), 3)) = "TEL" Then
					If Session("Categoria") = 1 Then 
						If rsGiro("sw_editado") = 1 Then
						'If (rsGiro("agente_pagador") = Session("CodigoAgente") Or _
						' rsGiro("agente_pagador") = Session("CodigoMatriz")) And rsGiro("sw_editado") = 1 Then
			%>
							'objMenu.addchild sId, "Pagar", "Pagar", "Principal"
							'objMenu.addchild sId, "Pagar a Tercero", "PagarTercero", "Principal"
							objMenu.addchild sId, "Solucionar", "Solucionar", "Principal"
							objMenu.addchild sId, "Actualizar Datos Cliente", "ADC", "Principal"
							nCont = nCont + 3
			<%	
						End If 
					ElseIf Session("Categoria") = 2 Then 
						'If ( _
						'		rsGiro("agente_pagador") = Session("CodigoAgente") _
						'		Or ( _
						'				InStr(1, "AV;AA;AQ;AB", Session("CodigoAgente")) <> 0 _
						'				And _
						'				InStr(1, "AV;AA;AQ;AB", rsGiro("agente_pagador")) <> 0 _									
						'			) _
						'	) _
						'	And rsGiro("sw_editado") = 1  _
						If rsGiro("sw_editado") = 1	Then
			%>
							'objMenu.addchild sId, "Pagar", "Pagar", "Principal"
							'objMenu.addchild sId, "Pagar a Tercero", "PagarTercero", "Principal"
							objMenu.addchild sId, "Solucionar", "Solucionar", "Principal"
							objMenu.addchild sId, "Actualizar Datos Cliente", "ADC", "Principal"
							nCont = nCont + 3
			<%	
						End If 
					ElseIf Session("Categoria") = 3 Then
						If rsGiro("sw_editado") = 1 Then
						'If rsGiro("agente_pagador") = Session("CodigoAgente") And rsGiro("sw_editado") = 1 Then
							If rsGiro("estado_giro") <> 1 Or Session("CodigoAgente") = Session("CodigoMoneyBroker") Then
			%>
							objMenu.addchild sId, "Solucionar", "Solucionar", "Principal"
							'	objMenu.addchild sId, "Pagar", "Pagar", "Principal"
							'	objMenu.addchild sId, "Pagar a Tercero", "PagarTercero", "Principal"
							'	objMenu.addchild sId, "Actualizar Datos Cliente", "ADC", "Principal"
							'	nCont = nCont + 3
			<%	
							End If
						End If
					ElseIf Session("Categoria") = 4 Then
						If rsGiro("agente_captador") = Session("CodigoAgente") And rsGiro("sw_editado") = 1 Then
						'If rsGiro("agente_pagador") = Session("CodigoAgente") And rsGiro("sw_editado") = 1 Then
						'	If rsGiro("estado_giro") <> 1 Or Session("CodigoAgente") = Session("CodigoMoneyBroker") Then
			%>
							objMenu.addchild sId, "Solucionar", "Solucionar", "Principal"
							'		objMenu.addchild sId, "Pagar", "PagarInt", "Principal"
							'		nCont = nCont + 1
			<%	
						'	End If
						'End If
						End If
					End If 
				End If								
'00000000000000000000000000000000000000000000000000000000000000000000000000000000000000000
'00000000000000000000000000000000000000000000000000000000000000000000000000000000000000000
'00000000000000000000000000000000000000000000000000000000000000000000000000000000000000000				
				
			End Select 
		
		%>
		For i = nCont To 6
			objMenu.addchild sId, "", "", ""
		Next
	End Sub

	Sub CargarMenuCliente()
		Dim sId, i, nCont

		objmenu.bgColor = document.bgColor 
		objmenu.stylesheet = "../Estilos/Cliente.css"
		sId = objmenu.addparent("Opciones")
				nCont = 0
		For i = nCont To 6
			objMenu.addchild sId, "", "", ""
		Next
	End Sub

	Sub txtMonto_OnKeyPress()
		IngresarTexto 1
	End Sub 

	Sub txtMonto_OnBlur()
		txtMonto.value = FormatNumber(txtMonto.value, 2)
		MostrarCotizacion	
	End Sub 
	
	Sub imgCalcular_onClick()
		MostrarCotizacion
	End Sub

	Sub MostrarCotizacion()
	
		lblTarifaIni.innerText = "12,00"
		lblTotalIni.innerText = FormatNumber( cCur(lblTarifaIni.innerText) + cCur(0 & txtMonto.value), 2)
		
	End Sub
 
 	Sub MostrarPaso(ByVal sPaso)
		window.tabPaso1.style.display = "none"
		window.tabPaso2.style.display = "none"
		window.tabPaso3.style.display = "none"
		If sPaso="tabPaso1" Then
			window.tabPaso1.style.display = ""
		ElseIf sPaso="tabPaso2" Then
			window.tabPaso2.style.display = ""
		Else
			window.tabPaso3.style.display = ""
		End If
	End Sub
	
	Sub MostrarPaso3(ByVal sPaso)
		MostrarPaso sPaso
		lblNombre.innerText = txtNombre.value & " " & txtApellido.value 
		lblMonto.innerText = txtMonto.value
		lblTarifa.innerText = lblTarifaIni.innerText
		lblTotal.innerText = lblTotalIni.innerText 
		lblMoneda.innerText = cbxMoneda.innerText 
		lblDireccion.innerText = txtDireccion.value 
		lblCiudad.innerText = cbxCiudad.innerText 
		lblPais.innerText = cbxPais.innerText 
		
	End Sub

	Sub CambiarCursor(Byval sControl)

		document.all.item(sControl).style.cursor = "Hand"
	
	End Sub
-->
</script>

<body>
<script LANGUAGE="VBScript">
<!--

	Const sEncabezadoFondo = ""
	Const sEncabezadoTitulo = "Detalle de Giro"
	Const sClass = "TituloPrincipal"
	
-->
</script>
<!--#INCLUDE virtual="/Compartido/Encabezado.htm" -->
<form id="frmGiro" method="post">
	<input type="hidden" name="txtRetira">
	<input type="hidden" name="txtRutRetira" value="<%=rsGiro("rut_retira")%>">
	<input type="hidden" name="txtPassRetira" value="<%=rsGiro("pasaporte_retira")%>">
	<input type="hidden" name="txtPaisPassRetira" value="<%=rsGiro("paispasap_retira")%>">
	<input type="hidden" name="txtNombresRetira" value="<%=MayMin(rsGiro("nombre_retira"))%>">
	<input type="hidden" name="txtApellidosRetira" value="<%=MayMin(rsGiro("apellido_retira"))%>">
	<input type="hidden" name="txtCodigoGiro" value="<%=sCodigo%>">
	<input type="hidden" name="txtMonedaPago" value="<%=rsGiro("codigo_moneda")%>">	
	<input type="hidden" name="txtCodigoBeneficiario" value="<%=rsGiro("codigo_beneficiario")%>">
	<input type="hidden" name="txtCodigoRemitente" value="<%=rsGiro("codigo_remitente")%>">
	<input type="hidden" name="txtTipoAviso">
	<input type="hidden" name="txtDescripcionAviso">
	<input type="hidden" name="txtParentesco">
	<input type="hidden" name="txtNombreParentesco">
	<input type="hidden" name="txtTipoReclamo">
	<input type="hidden" name="txtDescripcionReclamo">
	<input type="hidden" name="txtNombreSolucion">
	<input type="hidden" name="txtApellidoSolucion">
	<input type="hidden" name="txtDireccionSolucion">
	<input type="hidden" name="txtPaisFonoSolucion">
	<input type="hidden" name="txtAreaFonoSolucion">
	<input type="hidden" name="txtFonoSolucion">
	<input type="hidden" name="txtDescripcionSolucion">
	<input type="hidden" name="txtCodigoPagador" value="<%=rsGiro("codigo_pagador")%>">
	<input type="hidden" name="txtCodigoCaptador" value="<%=rsGiro("codigo_captador")%>">
	<input type="hidden" name="txtTipoGiro" value="<%=rsGiro("tipo_giro")%>">

	<% ' monto equivalente
		If ccur(rsGiro("monto_equivalente")) = ccur(0) Then
	%>
			<input type="hidden" name="txtMontoEquivalente" value="">
	<%
			
		Else
	%>
			<input type="hidden" name="txtMontoEquivalente" value="    BRR$ <%=FormatNumber(rsGiro("monto_equivalente"), 2)%>">			
	<%
		End If
	%>	

</form>

<OBJECT id=Printer1 style="LEFT: 0px; TOP: 0px"  codebase="AfexPrinter.CAB#version=1,0,0,0" 
	classid=CLSID:210A8E07-6FF9-4C4E-A664-5844036C0E33 VIEWASTEXT>
	<PARAM NAME="_ExtentX" VALUE="1">
	<PARAM NAME="_ExtentY" VALUE="1"></OBJECT>
<!--
<OBJECT id="Printer1" style="LEFT: 0px; TOP: 0px"  codebase="AfexPrinter.CAB#version=1,0,0,0" 
	classid=CLSID:4EB55F79-E861-11D7-AF26-00E04C9B1440 VIEWASTEXT>
	<PARAM NAME="_ExtentX" VALUE="1">
	<PARAM NAME="_ExtentY" VALUE="1"></OBJECT>
-->

<table cellspacing="0" cellpadding="0" border="0" style="LEFT: 6px; POSITION: absolute; TOP: 60px">
<tr><td>
	<!-- Paso 1 -->
	<table class="borde" ID="tabPaso1" CELLSPACING="0" CELLPADDING="2" BORDER="0" HEIGHT="100" WIDTH="560" STYLE="LEFT: 0px; POSITION: relative; TOP: 0px">
		<tr HEIGHT="15">
			<td colspan="5" CLASS="Titulo">&nbsp;&nbsp;Datos del Remitente</td>		
		</tr>
		<tr HEIGHT="15">
			<td></td>
			<input type="hidden" id="txtRut" value="<%=rsGiro("rut_beneficiario")%>">
			<input type="hidden" id="txtPass" value="<%=rsGiro("pasaporte_beneficiario")%>">
			<input type="hidden" id="txtPaisPass" value="<%=rsGiro("paispasap_beneficiario")%>">
			<input type="hidden" id="txtNombresR" value="<%=MayMin(rsGiro("nombre_remitente"))%>">
			<input type="hidden" id="txtApellidosR" value="<%=MayMin(rsGiro("apellido_remitente"))%>">
			<input type="hidden" id="txtDireccionR" value="<%=MayMin(rsGiro("direccion_remitente"))%>">
			<input type="hidden" id="txtRutR" value="<%=rsGiro("rut_remitente")%>">
			<input type="hidden" id="txtPassR" value="<%=rsGiro("pasaporte_remitente")%>">
			<input type="hidden" id="txtPaisPassR" value="<%=rsGiro("paispasap_remitente")%>">
			<input type="hidden" id="txtGastos" value="<%=rsGiro("gastos_transferencia")%>">
			<input type="hidden" id="txtTipoCambio" value="<%=rsGiro("tipo_cambio")%>">
			<input type="hidden" id="txtIva" value="<%=rsGiro("monto_iva")%>">
			<input type="hidden" id="txtPaisFonoR" value="<%=rsGiro("codpais_remitente")%>">
			<input type="hidden" id="txtAreaFonoR" value="<%=rsGiro("codarea_remitente")%>">
			<input type="hidden" id="txtFonoR" value="<%=rsGiro("fono_remitente")%>">
			<input type="hidden" id="txtCiudadR" value="<%=MayMin(rsGiro("nombre_ciudad_remitente"))%>">
			<input type="hidden" id="txtPaisR" value="<%=MayMin(rsGiro("nombre_pais_remitente"))%>">
			<td colspan="3">Nombre<br><input id="txtNombreR" SIZE="25" style="HEIGHT: 22px; WIDTH: 400px" value="<%=MayMin(Trim(rsGiro("nombre_remitente")) & " " & Trim(rsGiro("apellido_remitente")))%>" disabled></td>
			
		</tr>
		<tr HEIGHT="10">
			<td colspan="4" CLASS="Titulo">&nbsp;&nbsp;Datos del Beneficiario</td>
		</tr>
		<tr HEIGHT="15">
			<td colspan="4">
				<table><tr>
				<td>Nombre<br><input id="txtNombres" style="HEIGHT: 22px; WIDTH: 200px" value="<%=MayMin(Trim(rsGiro("nombre_beneficiario")))%>" disabled></td>
				<td>Apellido<br><input id="txtApellidos" style="HEIGHT: 22px; WIDTH: 180px" value="<%=MayMin(Trim(rsGiro("apellido_beneficiario")))%>" disabled></td>
				<% If Isnull(Trim(rsGiro("rut_beneficiario"))) Then %>
					<td>Identificacion<br><input id="txtId" style="HEIGHT: 22px; WIDTH: 100px" value="<%=Trim(rsGiro("pasaporte_retira"))%>" disabled></td>
				<% Else %>
					<td>Identificacion<br><input id="txtId" style="HEIGHT: 22px; WIDTH: 100px" value="<%=FormatoRut(Trim(rsGiro("rut_beneficiario")))%>" disabled></td>
				<% End If %>
				</tr></table>
			</td>
		</tr>
	<% If Session("Categoria") = 4 Then %>
			<tr HEIGHT="15" style="display: none">
	<%	Else %>
			<tr HEIGHT="15">
	<% End If %>
			<td colspan="4">
				<table><tr>
				<td COLSPAN="2">Dirección<br>
				<input id="txtDireccion" style="HEIGHT: 22px; WIDTH: 350px" value="<%=MayMin(rsGiro("direccion_beneficiario"))%>" disabled></td>
						<td>País<br>
							<input STYLE="HEIGHT: 22px; WIDTH: 80px" NAME="txtPais" value="<%=MayMin(rsGiro("nombre_pais_beneficiario"))%>" disabled>			
						</td>
						<td>Ciudad<br>
							<input STYLE="HEIGHT: 22px; WIDTH: 110px" NAME="txtCiudad" value="<%=MayMin(rsGiro("nombre_ciudad_beneficiario"))%>" disabled>			
						</td>
				</tr></table>
			</td>
		</tr>
		<tr HEIGHT="15">
			<td colspan="4">
				<table><tr>
				<td>Teléfono<br>
				<input disabled id="txtPaisFono" style="width: 20px" value="<%=rsGiro("codpais_beneficiario")%>">
				<input disabled id="txtAreaFono" style="width: 20px" value="<%=rsGiro("codarea_beneficiario")%>">
				<input id="txtFono" style="WIDTH: 80px" value="<%=rsGiro("fono_beneficiario")%>" disabled>
				</td>
				<td>Mensaje al Beneficiario<br><input id="txtMensaje" style="HEIGHT: 21px; WIDTH: 410px" SIZE="40" value="<%=Trim(rsGiro("mensaje"))%>">
				</tr></table>
			</td>
		</tr>
		<tr HEIGHT="10">
			<td colspan="4" CLASS="Titulo">&nbsp;&nbsp;Datos del Agente</td>			
		</tr>
	<% If Session("Categoria") = 4 Then %>
		<tr HEIGHT="15" style="display: none">
	<% Else %>
		<tr HEIGHT="15">
	<% End If %>
			<td colspan="4">
				<table><tr>
				<td>Agente Captador<br>
					<input STYLE="HEIGHT: 22px; WIDTH: 200px" NAME="txtCaptador" value="<%=MayMin(rsGiro("nombre_captador"))%>" disabled>			
				</td>			
				<td>Agente Pagador<br>
					<input STYLE="HEIGHT: 22px; WIDTH: 200px" NAME="txtPagador" value="<%=MayMin(rsGiro("nombre_pagador"))%>" disabled>			
				</td>			
				<td>Confirmación<br>
					<input STYLE="HEIGHT: 22px; WIDTH: 100px" NAME="txtConfirmacion" value="<%=rsGiro("codigo_confirmacion")%>" disabled>
				</td>
				</tr></table>
			</td>
		</tr>
		<tr HEIGHT="15">
			<td colspan="4">
				<table><tr>
				<td id="tdNN" STYLE="HEIGHT: 10px; WIDTH: 140px; display: none"></td>				
				<td id="tdMoneda">Moneda de Pago<br>				
					<input STYLE="HEIGHT: 22px; WIDTH: 140px; font-weight: bold; text-color: white; background-color: <%=sColorMoneda%>" NAME="txtCodigoMoneda" value="<%=MayMin(rsGiro("moneda"))%>" disabled>
				</td>
				<td>Monto<br>
					<input STYLE="HEIGHT: 22px; TEXT-ALIGN: right; WIDTH: 100px; font-weight: bold; text-color: white; background-color: <%=sColorMoneda%>" NAME="txtMonto"  value="<%=FormatNumber(rsGiro("monto_giro"), nDecimales)%>" disabled>
				</td>
				<td style="display: none">Tarifa<br>
					<input STYLE="HEIGHT: 22px; TEXT-ALIGN: right; WIDTH: 80px" NAME="txtTarifa" value="<%=FormatNumber(rsGiro("tarifa_cobrada"), nDecimales)%>" disabled>
				</td>
				<td style="display: none">Total<br>
					<input STYLE="HEIGHT: 22px; TEXT-ALIGN: right; WIDTH: 100px" NAME="txtTotal" value="<%=FormatNumber(cCur(0 & rsGiro("monto_giro")) + cCur(0 & rsGiro("tarifa_cobrada")), nDecimales)%>" disabled>
				</td>
				<td>Giro<br>
					<input STYLE="HEIGHT: 22px; WIDTH: 80px" NAME="txtGiro"  value="<%=rsGiro("codigo_giro")%>" disabled>
				</td>
				<td>Estado<br>
					<input STYLE="HEIGHT: 22px; WIDTH: 180px" NAME="txtEstado" value="<%=rsGiro("estado")%>" disabled>
				</td>
				</tr></table>
			</td>
		</tr>
		<tr HEIGHT="15">
			<td colspan="4">
				<table><tr>
				<td>Invoice<br>
					<input STYLE="HEIGHT: 22px; WIDTH: 100px" NAME="txtInvoice" value="<%=rsGiro("invoice")%>" disabled>
				</td>
				<td>Orden<br>
					<input STYLE="HEIGHT: 22px; WIDTH: 70px" NAME="txtOrden" value="<%=rsGiro("correlativo_salida")%>" disabled>
				</td>
				<td>Nº Boleta<br>
					<input STYLE="HEIGHT: 22px; TEXT-ALIGN: right; WIDTH: 70px" NAME="txtBoleta" value="<%=rsGiro("numero_documento")%>" disabled>
				</td>
				<td>Mensaje al Agente Pagador<br><input id="txtNota" style="HEIGHT: 21px; WIDTH: 295px" value="<%=rsGiro("nota")%>">
				</tr></table>
			</td>
		</tr>
		<tr HEIGHT="15">
			<td></td>
		</tr>
		<tr HEIGHT="18">
			<td colspan="4" CLASS="Titulo">&nbsp;&nbsp;Historia</td>
		</tr>
		<tr>
			<td colspan="4">
			<table cellspacing="1" cellpadding="1" ID="tbReporte" border="0" ALIGN="center" STYLE="COLOR: #505050; FONT-FAMILY: Verdana; FONT-SIZE: 10px; POSITION: relative; TOP: 0px">
				<tr CLASS="Encabezado" style="height: 20px">
					<td WIDTH="90px">
						<b>Fecha</b>
					</td>
					<td WIDTH="80px">
						<b>Hora</b>
					</td>
					<td WIDTH="360px">
						<b>Detalle</b>
					</td>
				</tr>
					
				<% 
					Dim sDetalle, rsHistoria,deta1
					
					Set rsHistoria = CreateObject("ADODB.Recordset")
					
					Set rsHistoria = BuscarHistoria(Session("afxCnxAFEXpress"), sCodigo)
										
					Do Until rsHistoria.EOF
						If sDetalle <> "Detalle1" Then
							sDetalle = "Detalle1"
						Else
							sDetalle = "Detalle2"
						End If
						deta1= rsHistoria("descripcion")
				%>
						<tr CLASS="<%=sDetalle%>">
						<td><%=rsHistoria("fecha") %></td>
						<td><%=rsHistoria("hora_formato") %></td>
						<td><%=rsHistoria("descripcion") %></td>				
						</tr>
					<% 
						rsHistoria.MoveNext
					Loop 
					%>
				</table>	
			</td>
		</tr>
	</table>
<% 
	Dim nTop
	
	nTop = 303
	If Session("Categoria") = 4 Then nTop = 210
%>
<div style="position: absolute; top: <%=nTop%>; left: 7; background-color: <%=sColorMoneda%>; color: white; height: 20; width: 138; font-size: 10pt; font-weight: bold">
<%=MayMin(rsGiro("moneda"))%>
</div>
<div style="position: absolute; top: <%=nTop%>; left: 151; background-color: <%=sColorMoneda%>; color: white; height: 20; width: 98; font-size: 10pt; font-weight: bold" align="right">
<%=FormatNumber(rsGiro("monto_giro"), nDecimales)%>
</div>

	<%
		'Set afxGiro = Nothing
		Set rsHistoria = Nothing
		Set rs = Nothing
		Set rsGiro = Nothing
	%>
</td>
<td valign="top">
</td></tr>
</table>
<table BORDER="0" cellspacing="0" cellpadding="0" STYLE="LEFT: 421px; POSITION: absolute; TOP: 40px">	
<tr><td>
    <object align="left" id="objMenu" style="HEIGHT: 111px; LEFT: 0px; POSITION: relative; TOP: 0px; WIDTH: 160px" type="text/x-scriptlet" width="170" VIEWASTEXT border="0" valign="top"><param NAME="Scrollbar" VALUE="0"><param NAME="URL" VALUE="http:../Scriptlets/Menu.htm"></object>
</td></tr>
</table>
</body>
<script LANGUAGE="VBScript">
<!--

	Function ValidarMBR()
		ValidarMBR = True
		Exit Function
		
		If frmGiro.txtCodigoCaptador.Value = "MG" And _
			Left(txtInvoice.Value, 2) = "MA" _ 
		Then
		
			window.showModalDialog "MsjMBR.htm", , "center=yes"
			MsgBox "GIRO EXACT CHANGE (ESPAÑA)" & vbCrLf & _
					 "CURSAR PAGO SÓLO CON AUTORIZACION DE CÓDIGO"  & vbCrLf & _
					 "LLAMANDO A ERIKA PEÑALOZA MONEYBROKER" & vbCrLf & _
					 "TEL. 2- 672 17 32", , "AFEX En Linea"
			Exit Function
		End If
		ValidarMBR = True
	End Function
	
	Sub objMenu_OnScriptletEvent(strEventName, varEventData)
		
	   Select Case strEventName
	   
			Case "linkClick"
				If Right(varEventData, 5) = "Pagar" Then
					If Not ValidarMBR() Then Exit Sub
					
					If <%=Session("IdCliente")%> = 0 Then
						'MsgBox "Imposible pagar el giro, el cliente no tiene identificación." , ,"Pagar"
						window.navigate "ActualizacionCliente.asp?AFEXchange=<%=sAFEXchange%>&AFEXpress=<%=sAFEXpress%>&Tipo=DetalleGiro&Giro=<%=sCodigo%>&Accion=<%=afxAccionPagar%>"
						window.close 
						Exit Sub
					End If
					If PagarBeneficiario() Then 
						If Not CajaPreguntaSiNo("AFEX En Linea", "Está seguro que desea pagar el giro?") Then
							Exit Sub
						End If
						Pagar
					End If

				ElseIf Right(varEventData, 8) = "PagarInt" Then
					PagarTercero
					If ValidarTercero() Then
						Pagar
					End If
'					If PagarBeneficiario() Then 
'						If Not CajaPreguntaSiNo("AFEX En Linea", "Está seguro que desea pagar el giro?") Then
'							Exit Sub
'						End If
'						Pagar
'					End If
					
				ElseIf Right(varEventData, 12) = "PagarTercero" Then
					If Not ValidarMBR() Then Exit Sub

					If <%=Session("IdCliente")%> = 0 Then
						window.navigate "ActualizacionCliente.asp?AFEXchange=<%=sAFEXchange%>&AFEXpress=<%=sAFEXpress%>&Tipo=DetalleGiro&Giro=<%=sCodigo%>&Accion=<%=afxAccionPagarTercero%>"
						window.close 
						'MsgBox "Imposible pagar el giro, el cliente no tiene identificación.",,"Pagar a Tercero"
						Exit Sub
					End If
					PagarTercero
					If ValidarTercero() Then
						Pagar
					End If

				ElseIf Right(varEventData, 15) = "PagarTerceroInt" Then
					PagarTercero
					If ValidarTercero() Then
						Pagar
					End If
					
				ElseIf Right(varEventData, 6) = "Avisar" Then
					Avisar
				ElseIf Right(varEventData, 6) = "Anular" Then
					window.navigate "AnularGiro.asp?AFEXchange=<%=sAFEXchange%>&AFEXpress=<%=sAFEXpress%>&detalle=<%=deta1%>&Monto=<%=cMontoGiro%>&Giro=<%=scodigo%>&Invoice=<%=rsGiro("invoice")%>"
					window.close
				ElseIf Right(varEventData, 8) = "Corregir" Then
					window.navigate "CorregirBeneficiario.asp?AFEXchange=<%=sAFEXchange%>&AFEXpress=<%=sAFEXpress%>&detalle=<%=deta1%>&Monto=<%=cMontoGiro%>&Giro=<%=scodigo%>&Invoice=<%=rsGiro("invoice")%>"
					window.close 								
				ElseIf Right(varEventData, 8) = "Reclamar" Then
					If Not CajaPreguntaSiNo("AFEX En Linea", "Está seguro que desea reclamar el giro?") Then
						Exit Sub
					End If
					Reclamar

				ElseIf Right(varEventData, 10) = "Solucionar" Then
					Solucionar

				ElseIf Right(varEventData, 11) = "Comprobante" Then
					If Not CajaPregunta("AFEX En Linea", "Coloque el comprobante en la impresora y haga click en Aceptar") Then						
						Exit Sub
					End If
					<% If Session("ModoPrueba") Then %>
							ImprimirComprobantePago
					<% Else %>
							ImprimirComprobantePago
					<% End If %> 
					window.navigate "AtencionClientes.asp"
					window.close 
					
				ElseIf Right(varEventData, 9) = "Servicios" Then
					If Not CajaPregunta("AFEX En Linea", "Coloque la boleta en la impresora y haga click en Aceptar") Then
						Exit Sub
					End If
					<% If Session("ModoPrueba") Then %>
							ImprimirBoletaServicios
					<% Else %>
							ImprimirBoletaServicios
					<% End If %> 
					window.navigate "AtencionClientes.asp"
					window.close 

				ElseIf Right(varEventData, 3) = "ATC" Then
					window.navigate "AtencionClientes.asp?Accion=<%=afxAccionBuscar%>&Campo=<%=afxCampoCodigoExpress%>&Argumento=<%=sAFEXpress%>"
					window.close 

				ElseIf Right(varEventData, 3) = "ADC" Then
					'window.navigate "ActualizacionCliente.asp?AFEXchange=<%=Request("AFEXchange")%>&AFEXpress=<%=Request("AFEXpress")%>&Tipo=DetalleGiro&Giro=<%=sCodigo%>"
					window.navigate "ActualizacionCliente.asp?AFEXchange=<%=sAFEXchange%>&AFEXpress=<%=sAFEXpress%>&Tipo=DetalleGiro&Giro=<%=sCodigo%>"
					'window.navigate "ActualizacionCliente.asp?AFEXchange=<%=sAFEXchange%>&AFEXpress=<%=sAFEXpress%>&Tipo=DetalleGiro&Giro=<%=sCodigo%>&Accion=<%=afxAccionPagar%>"					
					window.close 
									
				ElseIf Right(varEventData, 6) = "OfPago" Then
					'window.navigate "ActualizacionCliente.asp?AFEXchange=<%=Request("AFEXchange")%>&AFEXpress=<%=Request("AFEXpress")%>&Tipo=DetalleGiro&Giro=<%=sCodigo%>"
					window.showModalDialog "OficinasPago.asp?CodPag=<%=sCodigoPagador%>&PaisB=<%=sPaisBeneficiario%>&CiudadB=<%=sCiudadBeneficiario%>", , "dialogwidth:40;dialogheight:20"
					'window.navigate "ActualizacionCliente.asp?AFEXchange=<%=sAFEXchange%>&AFEXpress=<%=sAFEXpress%>&Tipo=DetalleGiro&Giro=<%=sCodigo%>&Accion=<%=afxAccionPagar%>"					
					window.close 
				ElseIf Right(varEventData, 2) = "BS" Then
					MostrarBS
				ElseIf Right(varEventData, 2) = "CP" Then
					MostrarCP
				End If

		End Select		
	End Sub

	Sub PagarTercero()
		Dim sString, aT
				
		frmGiro.txtRutRetira.value = ""
		frmGiro.txtPassRetira.value = ""
		frmGiro.txtPaisPassRetira.value = ""
		frmGiro.txtNombresRetira.value = ""
		frmGiro.txtApellidosRetira.value = ""
		frmGiro.txtRetira.value = 1
		sString = Empty
		sString = window.showModalDialog( _
													 "DatosTercero.asp?Rut=" & frmGiro.txtRutRetira.value & _
													 "&Pasaporte=" & frmGiro.txtPassRetira.value & _
													 "&PaisPasaporte=" & frmGiro.txtPaisPassRetira.value & _
													 "&Nombres=" & frmGiro.txtNombresRetira.value & _
													 "&Apellidos=" & frmGiro.txtApellidosRetira.value _
													)

		If sString <> Empty Then
			' crea un arreglo con los datos de la transferencia seleccionada
			aT = Split(sString, ";", 5)
			
			' asigna los datos de la transferencia
			frmGiro.txtRutRetira.value = aT(0)
			frmGiro.txtPassRetira.value = aT(1)
			frmGiro.txtPaisPassRetira.value = aT(2)
			frmGiro.txtNombresRetira.value = aT(3)
			frmGiro.txtApellidosRetira.value = aT(4)
		End If
		
	End Sub

	Function ValidarTercero()
	
		ValidarTercero = False
		
		If Trim(frmGiro.txtRutRetira.value & frmGiro.txtPassRetira.value) = "" Then
			MsgBox "Debe ingresar la identificación de la persona que retira el dinero",, "Pago a Tercero"
			Exit Function
		End If				
		If Trim(frmGiro.txtPassRetira.value) <> "" And Trim(frmGiro.txtPaisPassRetira.value) = "" Then
			MsgBox "Debe ingresar el pais de la persona que retira el dinero",, "Pago a Tercero"
			Exit Function
		End If				
		If Trim(frmGiro.txtNombresRetira.value) = "" Then
			MsgBox "Debe ingresar el nombre de la persona que retira el dinero",, "Pago a Tercero"
			Exit Function
		End If				
		If Trim(frmGiro.txtApellidosRetira.value) = "" Then
			MsgBox "Debe ingresar los apellidos de la persona que retira el dinero",, "Pago a Tercero"
			Exit Function
		End If				
			
		ValidarTercero = True
		
	End Function

	Function PagarBeneficiario()	
		frmGiro.txtRetira.value = 0
		frmGiro.txtRutRetira.value = txtRut.value
		frmGiro.txtPassRetira.value = txtpass.value
		frmGiro.txtPaisPassRetira.value = txtPaisPass.value
		frmGiro.txtNombresRetira.value = txtNombres.value
		frmGiro.txtApellidosRetira.value = txtApellidos.value
		PagarBeneficiario = True
	End Function
		
	Sub Pagar()		
		<% If nEstado = 7 Then %>		
				frmGiro.action = "GrabarPagoGiroReclamo.asp"
		<% Else %>
				frmGiro.action = "GrabarPagoGiro.asp?eg=<%=nEstado%>"
		<% End If %>	
		frmGiro.submit 
		frmGiro.action = ""
	End Sub

	Function DatosAviso()
		Dim sAviso, aAviso
		
		DatosAviso = False
		sAviso = Trim(window.showModalDialog("DatosAviso.asp"))
		
		If sAviso = "" Then Exit Function
		
		aAviso = Split(sAviso, ";", 4)
		frmGiro.txtTipoAviso.value = aAviso(0)
		frmGiro.txtDescripcionAviso.value = aAviso(1)
		frmGiro.txtParentesco.value = aAviso(2)
		frmGiro.txtNombreParentesco.value = aAviso(3)
		DatosAviso = True
		
	End Function
	
	Function ValidarAviso()
	
		ValidarAviso = False
		
		If Trim(frmGiro.txtTipoAviso.value) = "" Then
			MsgBox "Debe seleccionar el tipo de aviso",, "Avisar"
			Exit Function
		End If				
					
		ValidarAviso = True
		
	End Function

	Sub Avisar()
		If Not DatosAviso Then Exit Sub
		If Not ValidarAviso() Then Exit Sub
		frmGiro.action = "GrabarAvisoGiro.asp?TipoLista=<%=nTipoLista%>" ' &Pagador=" & frmGiro.txtCodigoPagador.value &Cliente=<%=sAFEXpress%>"
		frmGiro.submit 
		frmGiro.action = ""	
	End Sub	

	Function DatosReclamo()
		Dim sReclamo, aReclamo
		
		DatosReclamo = False
		sReclamo = Trim(window.showModalDialog("DatosReclamo.asp"))
		
		If sReclamo = "" Then Exit Function
		
		aReclamo = Split(sReclamo, ";", 3)
		'msgbox areclamo(0) & ", " & areclamo(1) & ", " & areclamo(2)
		'exit function
		frmGiro.txtTipoReclamo.value = aReclamo(0)
		frmGiro.txtDescripcionReclamo.value = aReclamo(1)		
		DatosReclamo = True
		
	End Function
	
	Function ValidarReclamo()
	
		ValidarReclamo = False
		
		If Trim(frmGiro.txtTipoReclamo.value) = "" Then
			MsgBox "Debe seleccionar el tipo de reclamo",, "Reclamar"
			Exit Function
		End If				
					
		ValidarReclamo = True
		
	End Function

	Sub Reclamar()
		If Not DatosReclamo Then Exit Sub
		If Not ValidarReclamo() Then Exit Sub
		frmGiro.action = "GrabarReclamoGiro.asp"
		frmGiro.submit 
		frmGiro.action = ""	
	End Sub	

	Function DatosSolucion()
		Dim sSolucion, aSolucion

		frmGiro.txtNombreSolucion.value = txtNombres.value
		frmGiro.txtApellidoSolucion.value = txtApellidos.value
		frmGiro.txtDireccionSolucion.value = txtDireccion.value
		frmGiro.txtPaisFonoSolucion.value = txtPaisFono.value
		frmGiro.txtAreaFonoSolucion.value = txtAreaFono.value
		frmGiro.txtFonoSolucion.value = txtFono.value
		frmGiro.txtDescripcionSolucion.value = frmGiro.txtDescripcionSolucion.value 
		
		DatosSolucion = False
		sSolucion = window.showModalDialog("DatosSolucion.asp?Nombres=" & frmGiro.txtNombreSolucion.value & _
						"&Apellidos=" & frmGiro.txtApellidoSolucion.value & _
						"&Direccion=" & frmGiro.txtDireccionSolucion.value & _
						"&AreaFono=" & frmGiro.txtAreaFonoSolucion.value & _
						"&PaisFOno=" & frmGiro.txtPaisFonoSolucion.value & _
						"&Fono=" & frmGiro.txtFonoSolucion.value & _
						"&Descripcion=" & frmGiro.txtDescripcionSolucion.value  _
						) 
		
		If sSolucion = "" Then Exit Function
		
		aSolucion = Split(sSolucion, ";", 5)
		frmGiro.txtNombreSolucion.value = aSolucion(0)
		frmGiro.txtApellidoSolucion.value = aSolucion(1)
		frmGiro.txtDireccionSolucion.value = aSolucion(2)
		frmGiro.txtFonoSolucion.value = aSolucion(3)
		frmGiro.txtDescripcionSolucion.value = aSolucion(4)
		DatosSolucion = True
		
	End Function
	
	Function ValidarSolucion()
	
		ValidarSolucion = False
		
		If Trim(frmGiro.txtDescripcionSolucion.value) = "" Then
			MsgBox "Debe ingresar la descripción de la solución",, "Solucionar"
			Exit Function
		End If				
					
		ValidarSolucion = True
		
	End Function

	Sub Solucionar()
		If Not DatosSolucion Then Exit Sub
		If Not ValidarSolucion() Then Exit Sub
		'Exit Sub
		frmGiro.action = "GrabarSolucionGiro.asp"
		frmGiro.submit 
		frmGiro.action = ""	
	End Sub	
	
	Function ImprimirComprobantePagoOld()
		Dim sLinea, afxPrinter, sTipoId, sId, sIdT, sTipoIdT

		ImprimirComprobantePagoOld = False
		If txtRut.value <> "" Then
			sTipoId = "R"
			sId = FormatoRut(txtRut.value)
		Else
			sTipoId = "P"
			sId = Trim(txtPass.value) & ";" & Trim(txtPaisPass.value)
		End If
		If Trim(frmGiro.txtPassRetira.value) <> "" Then
			sTipoIdT = "P"
			sIdT = Trim(frmGiro.txtPassRetira.value) & ";" & Trim(frmGiro.txtPaisPassRetira.value)
		Else
			sTipoIdT = "R"
			sIdT = FormatoRut(frmGiro.txtRutRetira.value)
		End If
		
		sLinea = "<%=Date()%>" & "\\" & "<%=Time()%>" & "\\" & " " & "\\" & _
					Trim(Trim(txtNombres.value) & " " & Trim(txtApellidos.value)) & "\\" & sTipoId & "\\" & _
					sId & "\\" & txtDireccion.value & "\\" & _
					txtCiudad.value & "\\(" & _
					txtPaisFono.value & txtAreaFono.value & ")" & txtFono.value & "\\" & _
					frmGiro.txtCodigoGiro.value & "\\" & txtInvoice.value & "\\" & _
					txtCaptador.value & "\\" & _
					txtPagador.value & "\\" & txtNombreR.value & "\\" & _
					txtCiudadR.value & "\\" & txtPaisR.value & "\\" & _
					txtMensaje.value & "\\" & _
					frmGiro.txtNombresRetira.value & " " & frmGiro.txtApellidosRetira.value & "\\" & _
					sTipoIdT & "\\" & sIdT & "\\" & _
					"US$   " & txtMonto.value
		
		'msgbox slinea
		On Error GoTo 0
		Set afxPrinter = CreateObject("afxOcx.AFEX_Imprimir")
		afxPrinter.Lin_Impre = sLinea 
		If Not afxPrinter.Imprimir_ComPago Then
			Set afxPrinter = Nothing
			MsgBox "Se produjo un error al intentar imprimir el Comprobante de Pago"
		End If
		
		ImprimirComprobantePagoOld = True
		Set afxPrinter = Nothing
		On Error Resume Next
		
	End Function


	Function ImprimirComprobantePago()
		Dim sLinea, afxPrinter, sTipoId, sId, sIdT, sTipoIdT

		ImprimirComprobantePago = False
		If txtRut.value <> "" Then
			sTipoId = "R"
			sId = FormatoRut(txtRut.value)
		Else
			sTipoId = "P"
			sId = Trim(txtPass.value) & ";" & Trim(txtPaisPass.value)
		End If
		If Trim(frmGiro.txtPassRetira.value) <> "" Then
			sTipoIdT = "P"
			sIdT = Trim(frmGiro.txtPassRetira.value) & ";" & Trim(frmGiro.txtPaisPassRetira.value)
		Else
			sTipoIdT = "R"
			sIdT = FormatoRut(frmGiro.txtRutRetira.value)
		End If
		Dim sPrefijo
		If frmGiro.txtMonedaPago.value = "<%=Session("MonedaNacional")%>" Then
			sPrefijo = "  $"
		Else
			sPrefijo = "US$"
		End If
		
		sLinea = "<%=Date()%>" & "\\" & "<%=Time()%>" & "\\" & " " & "\\" & _
					Trim(Trim(txtNombres.value) & " " & Trim(txtApellidos.value)) & "\\" & sTipoId & "\\" & _
					sId & "\\" & txtDireccion.value & "\\" & _
					txtCiudad.value & "\\(" & _
					txtPaisFono.value & txtAreaFono.value & ")" & txtFono.value & "\\" & _
					frmGiro.txtCodigoGiro.value & "\\" & txtInvoice.value & "\\" & _
					txtCaptador.value & "\\" & _
					txtPagador.value & "\\" & txtNombreR.value & "\\" & _
					txtCiudadR.value & "\\" & txtPaisR.value & "\\" & _
					txtMensaje.value & "\\" & _
					frmGiro.txtNombresRetira.value & " " & frmGiro.txtApellidosRetira.value & "\\" & _
					sTipoIdT & "\\" & sIdT & "\\" & _
					sPrefijo & "   " & txtMonto.value
		
		'msgbox slinea
		On Error GoTo 0
		'Set afxPrinter = CreateObject("afxOcx.AFEX_Imprimir")
		'afxPrinter.Lin_Impre = sLinea 
		Set afxPrinter = CreateObject("AfexPrinter.Printer")
		If Not afxPrinter.ComprobantePago(sLinea) Then
		'If Not Printer1.ComprobantePago(sLinea) Then
			Set afxPrinter = Nothing
			MsgBox "Se produjo un error al intentar imprimir el Comprobante de Pago"
		End If
		ImprimirComprobantePago = True
		Set afxPrinter = Nothing
		On Error Resume Next
		
	End Function

	Function ImprimirBoletaServicios()
		Dim sLinea, afxPrinter2, sTipoId, sId, sIdR, sTipoIdR
		Dim sPromocion, sPIN, nCorrelativo, rs, Categoria
		
		ImprimirBoletaServicios = False
		
		sPIN = ""
		nCorrelativo = 0
		Categoria = "<%=nCategoria%>"
		'msgbox "Captador: " & "<%=sCodigoCaptador%>"
		
		If cDate("<%=dFechaGiro%>") >= cDate("15/12/2007") Then
	
			If Ucase(Trim("<%=sCodigoPagador%>")) = "IK" Then
				
				If UCase(Trim(txtPais.value)) = "PERU" Or UCase(trim(txtPais.value)) = "PERÚ" Then				
				
				  'If ("<%=Session("CodigoAgente")%>" = "AM" And "<%=sCodigoCaptador%>" = "AM") Or ("<%=Session("CodigoAgente")%>" = "AX" And "<%=sCodigoCaptador%>" = "AX") Or _
				  '   ("<%=Session("CodigoAgente")%>" = "AG" And "<%=sCodigoCaptador%>" = "AG") Or ("<%=Session("CodigoAgente")%>" = "ZC" And "<%=sCodigoCaptador%>" = "ZC") Or _
				  '	 ("<%=Session("CodigoAgente")%>" = "AA" And "<%=sCodigoCaptador%>" = "AA") Or ("<%=Session("CodigoAgente")%>" = "AB" And "<%=sCodigoCaptador%>" = "AB") Or _
					' ("<%=Session("CodigoAgente")%>" = "AC" And "<%=sCodigoCaptador%>" = "AC") Or ("<%=Session("CodigoAgente")%>" = "AD" And "<%=sCodigoCaptador%>" = "AD") Or _
					' ("<%=Session("CodigoAgente")%>" = "AF" And "<%=sCodigoCaptador%>" = "AF") Or ("<%=Session("CodigoAgente")%>" = "AH" And "<%=sCodigoCaptador%>" = "AH") Or _
					' ("<%=Session("CodigoAgente")%>" = "AI" And "<%=sCodigoCaptador%>" = "AI") Or ("<%=Session("CodigoAgente")%>" = "AJ" And "<%=sCodigoCaptador%>" = "AJ") Or _
					' ("<%=Session("CodigoAgente")%>" = "AK" And "<%=sCodigoCaptador%>" = "AK") Or ("<%=Session("CodigoAgente")%>" = "AI" And "<%=sCodigoCaptador%>" = "AI") Or _
					' ("<%=Session("CodigoAgente")%>" = "AL" And "<%=sCodigoCaptador%>" = "AL") Or ("<%=Session("CodigoAgente")%>" = "AN" And "<%=sCodigoCaptador%>" = "AN") Or _
					' ("<%=Session("CodigoAgente")%>" = "AO" And "<%=sCodigoCaptador%>" = "AO") Or ("<%=Session("CodigoAgente")%>" = "AP" And "<%=sCodigoCaptador%>" = "AP") Or _
					' ("<%=Session("CodigoAgente")%>" = "AQ" And "<%=sCodigoCaptador%>" = "AQ") Or ("<%=Session("CodigoAgente")%>" = "AR" And "<%=sCodigoCaptador%>" = "AR") Or _
					' ("<%=Session("CodigoAgente")%>" = "AS" And "<%=sCodigoCaptador%>" = "AS") Or ("<%=Session("CodigoAgente")%>" = "AT" And "<%=sCodigoCaptador%>" = "AT") Or _
					' ("<%=Session("CodigoAgente")%>" = "AU" And "<%=sCodigoCaptador%>" = "AU") Or ("<%=Session("CodigoAgente")%>" = "AV" And "<%=sCodigoCaptador%>" = "AV") Or _
					' ("<%=Session("CodigoAgente")%>" = "AW" And "<%=sCodigoCaptador%>" = "AW") Or ("<%=Session("CodigoAgente")%>" = "AY" And "<%=sCodigoCaptador%>" = "AY") Or _
					' ("<%=Session("CodigoAgente")%>" = "AZ" And "<%=sCodigoCaptador%>" = "AZ") Or ("<%=Session("CodigoAgente")%>" = "ZA" And "<%=sCodigoCaptador%>" = "ZA") Or _
					' ("<%=Session("CodigoAgente")%>" = "ZB" And "<%=sCodigoCaptador%>" = "ZB") Or ("<%=Session("CodigoAgente")%>" = "ZD" And "<%=sCodigoCaptador%>" = "ZD") Or _
					' ("<%=Session("CodigoAgente")%>" = "ZD" And "<%=sCodigoCaptador%>" = "ZD") Or ("<%=Session("CodigoAgente")%>" = "ZE" And "<%=sCodigoCaptador%>" = "ZE") Or _
					' ("<%=Session("CodigoAgente")%>" = "ZF" And "<%=sCodigoCaptador%>" = "ZF") Or ("<%=Session("CodigoAgente")%>" = "ZG" And "<%=sCodigoCaptador%>" = "ZG") Or _ 
					' ("<%=Session("CodigoAgente")%>" = "ZH" And "<%=sCodigoCaptador%>" = "ZH") Or ("<%=Session("CodigoAgente")%>" = "ZJ" And "<%=sCodigoCaptador%>" = "ZJ") Or _
					' ("<%=Session("CodigoAgente")%>" = "ZK" And "<%=sCodigoCaptador%>" = "ZK") Or ("<%=Session("CodigoAgente")%>" = "ZL" And "<%=sCodigoCaptador%>" = "ZL") Then
					if ("<%=Session("CodigoAgente")%>" = "<%=sCodigoCaptador%>")and (Categoria = 1 Or Categoria = 2) then	
					'If (Categoria = 1 Or Categoria = 2)  Then
											
							'On Error Resume Next
					<%		If Not BuscarPIN(sCodigo) Then									
								If Not ObtenerDatosPIN Then								
									'Exit Function									
								End If
							End If
					%>
							
							sPIN = "<%=sNumeroPIN%>"
							
						<%
							If cCur(0 & Correlativo) <> 0 Then
								
								If Not ActualizarPIN(Correlativo, sCodigo) Then %>									
									'Exit Function
						<%		End If
							End If
						%>	
						
						If Trim(sPIN) <> "" Then						
							sPromocion = "Promoción: Llame gratis desde Santiago desde red fija al 760 0112 " & _
										 "o desde Santiago y regiones desde red fija al 112 800 112 344 y disque el Nº " & space(17) & _
										 "PIN " & sPIN & "para realizar su llamada al Perú."
						End If		
					Else
						sPromocion = Trim(txtNota.value)
					End If
				'Else
				'	sPromocion = Trim(txtNota.value)
				'End If
			'ElseIf "<%=cMontoGiro%>" >= 150 And Ucase(Trim("<%=sCodigoPagador%>")) = Trim("<%=Session("CodigoMGEnvio")%>") Then
			'	<% If Not BuscarPIN(sCodigo) Then
			'	       If Not ObtenerDatosPIN Then
			'		       'Exit Function
			'		   End If
			'	   End If
			'	%>
			'	   sPIN = "<%=sNumeroPIN%>"
			'	<%
			'	   If cCur(0 & Correlativo) <> 0 Then		
			'	      If Not ActualizarPIN(Correlativo, sCodigo) Then %>
			'		      'Exit Function
			'		<%End If
			'	   End If
			'	%>	
			'	If Trim(sPIN) <> "" Then
			'		sPromocion = "Promoción: Llame gratis al 800 211221 y utilice el Nº " & space(16) & _
			'					 "PIN " & sPIN & " para realizar su llamada telefónica"
			'	End If		
			Else
				sPromocion = Trim(txtNota.value)
			End If
		End If
	end if	
						
		If txtRut.value <> "" Then
			sTipoId = "R"
			sId = FormatoRut(txtRut.value)
		Else
			sTipoId = "P"
			sId = txtPass.value & txtPaisPass.value			
		End If
		If txtRutR.value <> "" Then
			sTipoIdR = "R"
			sIdR = FormatoRut(txtRutR.value)
		Else
			sTipoIdR = "P"
			sIdR = txtPassR.value & txtPaisPassR.value			
		End If
		Dim sPrefijo
		If frmGiro.txtMonedaPago.value = "<%=Session("MonedaNacional")%>" Then
			sPrefijo = "  $"
			txtTipoCambio.value = 1
		Else
			sPrefijo = "US$"
		End If
		
		sLinea = "<%=Date()%>" & "\\" & "<%=Time()%>" & "\\" & _
		" " & "\\" & trim(txtNombresR.value) & "\\" & _
		trim(txtApellidosR.value) & "\\" & trim(sIdR) & "\\" & trim(txtDireccionR.value) & _
		"\\" & trim(txtCiudadR.value) & "\\" & _
		"(" & trim(txtPaisFonoR.value) & trim(txtAreaFonoR.value) & ") " & trim(txtFonoR.value) & _
		"\\" & trim(txtCiudad.value) & "\\" & trim(txtPais.value) & "\\" & _
		trim(txtNombres.value) & "\\" & trim(txtApellidos.value) & "\\" & _
		trim(txtDireccion.value) & "\\" & _
		"(" & trim(txtPaisFono.value) & trim(txtAreaFono.value) & ") " & trim(txtFono.value) & _
		"\\" & trim(txtMensaje.value) & "\\" & trim(frmGiro.txtMonedaPago.value) & "\\" & trim(frmGiro.txtMonedaPago.value) & "\\" & _
		trim(txtMonto.value) & "\\" & trim(txtGastos.value) & "\\" & trim(txtIva.value) & "\\" & _
		trim(txtTipoCambio.value) & "\\" & trim(frmGiro.txtCodigoGiro.value) & "\\" & trim(sPromocion)  & "\\" & ROUND(cCur(0 & txtIva.value) * cCur(0 & txtTipoCambio.value), 0) & _
		"\\" & frmGiro.txtMontoEquivalente.value
		On Error GoTo 0
		Set afxPrinter2 = CreateObject("AfexPrinter.Printer")
		
		If Not afxPrinter2.BoletaServicios(sLinea, 10) Then
			Set afxPrinter2 = Nothing
			MsgBox "Se produjo un error al intentar imprimir la Boleta de Servicios"
			Exit Function
		End If
		
		ImprimirBoletaServicios = True
		Set afxPrinter2 = Nothing
		' aqui On Error Resume Next
		
	End Function
	
	Function ImprimirBoletaServiciosOld()
		Dim sLinea, afxPrinter2, sTipoId, sId, sIdR, sTipoIdR

		ImprimirBoletaServiciosOld = False
		If txtRut.value <> "" Then
			sTipoId = "R"
			sId = FormatoRut(txtRut.value)
		Else
			sTipoId = "P"
			sId = txtPass.value & txtPaisPass.value			
		End If
		If txtRutR.value <> "" Then
			sTipoIdR = "R"
			sIdR = FormatoRut(txtRutR.value)
		Else
			sTipoIdR = "P"
			sIdR = txtPassR.value & txtPaisPassR.value			
		End If
		
		sLinea = "<%=Date()%>" & "\\" & "<%=Time()%>" & "\\" & _
		" " & "\\" & trim(txtNombresR.value) & "\\" & _
		trim(txtApellidosR.value) & "\\" & trim(sIdR) & "\\" & trim(txtDireccionR.value) & _
		"\\" & trim(txtCiudadR.value) & "\\" & _
		"(" & trim(txtPaisFonoR.value) & trim(txtAreaFonoR.value) & ") " & trim(txtFonoR.value) & _
		"\\" & trim(txtCiudad.value) & "\\" & trim(txtPais.value) & "\\" & _
		trim(txtNombres.value) & "\\" & trim(txtApellidos.value) & "\\" & _
		trim(txtDireccion.value) & "\\" & _
		"(" & trim(txtPaisFono.value) & trim(txtAreaFono.value) & ") " & trim(txtFono.value) & _
		"\\" & trim(txtMensaje.value) & "\\" & trim(frmGiro.txtMonedaPago.value) & "\\" & trim(frmGiro.txtMonedaPago.value) & "\\" & _
		trim(txtMonto.value) & "\\" & trim(txtGastos.value) & "\\" & trim(txtIva.value) & "\\" & _
		trim(txtTipoCambio.value) & "\\" & trim(frmGiro.txtCodigoGiro.value) & "\\" & trim(txtNota.value)
		On Error GoTo 0
		Set afxPrinter2 = CreateObject("afxOcx.AFEX_Imprimir")
		afxPrinter2.Lin_Impre = sLinea 
		If Not afxPrinter2.Imprimir_Boleta_Servicios() Then
			Set afxPrinter2 = Nothing			
			MsgBox "Se produjo un error al intentar imprimir la Boleta de Servicios"
			Exit Function
		End If
		
		ImprimirBoletaServiciosOld = True
		Set afxPrinter2 = Nothing
		' aqio On Error Resume Next
		
	End Function

	Sub window_onunload()		
		dim a
		
	End Sub


	Sub MostrarBS()
		Dim sSolucion, aSolucion, sDetalle, nDec
		
		txtDireccionR.value = Replace(txtDireccionR.value, "#", "")
		txtDireccion.value = Replace(txtDireccion.value, "#", "")
		txtMensaje.value = Replace(txtMensaje.value, "#", "")
		If frmGiro.txtMonedaPago.value = "<%=Session("MonedaNacional")%>" Then nDec = 0 Else nDec = 2
		sDetalle = 	"?Nombres=" & txtNombresR.value & _
						"&Apellidos=" & txtApellidosR.value & _
						"&Direccion=" & txtDireccionR.value & _
						"&AreaFono=" & txtAreaFonoR.value & _
						"&PaisFono=" & txtPaisFonoR.value & _
						"&Fono=" & txtFonoR.value & _
						"&Ciudad=" & Trim(txtCiudadR.value) & _
						"&Rut=" & txtRutR.value & _
						"&Codigo=" & txtGiro.value & _
						"&sPIN=" & "<%=sNumeroPIN%>" & _
						"&CiudadB=" & txtCiudad.value & _
						"&PaisB=" & txtPais.value & _
						"&NombresB=" & txtNombres.value & _
						"&ApellidosB=" & txtApellidos.value & _
						"&DireccionB=" & txtDireccion.value & _
						"&FonoB=(" & txtPaisFono.value & txtAreaFono.value &  ") " & txtFono.value & _
						"&Mensaje=" & txtMensaje.value & _
						"&Monto=" & frmGiro.txtMonedaPago.value & "  " & FormatNumber(cCur(0 & txtMonto.value), nDec) & _
						"&Gastos=" & frmGiro.txtMonedaPago.value & "  " & FormatNumber(cCur(0 & trim(txtGastos.value)), nDec) & _
						"&Comision=" & frmGiro.txtMonedaPago.value & "  " & FormatNumber(cCur(0 & trim(txtIva.value)), nDec) & _
						"&Total=" & frmGiro.txtMonedaPago.value & "  " & Formatnumber(cCur(0 & trim(txtMonto.value)) +  cCur(0 & trim(txtGastos.value)), nDec) & _
						"&MontoEquivalente=" & frmGiro.txtMontoEquivalente.value & _ 
						"&TotalNacional=" & "<%=Session("MonedaNacional")%> "
		If frmGiro.txtMonedaPago.value = "<%=Session("MonedaNacional")%>" Then
			sDetalle = sDetalle & Formatnumber(ROUND(cCur(0 & txtIva.value), 0), nDec)
		Else
			sDetalle = sDetalle & Formatnumber(ROUND(cCur(0 & txtIva.value) * cCur(0 & txtTipoCambio.value), 0), nDec)
			sDetalle = sDetalle & "&TipoCambio=" & frmGiro.txtMonedaPago.value & "  " & FormatNumber(cCur(0 & trim(txtTipoCambio.value)), nDec)
		End If
		sSolucion = window.showModalDialog("BoletaServicios.asp" & sDetalle, , "center=yes" ) 				
	End Sub


	Sub MostrarCP()
		Dim sIdR, sTipoIdR, sIdB, sTipoIdB, nDec
	
		If frmGiro.txtRutRetira.value <> "" Then
			sTipoIdR = "R"
			sIdR = FormatoRut(frmGiro.txtRutRetira.value)
		Else
			sTipoIdR = "P"
			sIdR = frmGiro.txtPassR.value & frmGiro.txtPaisPassRetira.value			
		End If

		If txtRut.value <> "" Then
			sTipoIdB = "R"
			sIdB = FormatoRut(txtRut.value)
		Else
			sTipoIdB = "P"
			sIdB = txtPassR.value & txtPaisPass.value			
		End If

		If frmGiro.txtMonedaPago.value = "<%=Session("MonedaNacional")%>" Then nDec = 0 Else nDec = 2
		txtDireccionR.value = Replace(txtDireccionR.value, "#", "")
		txtDireccion.value = Replace(txtDireccion.value, "#", "")
		txtMensaje.value = Replace(txtMensaje.value, "#", "")
		sSolucion = window.showModalDialog("ComprobantePago.asp?Nombres=" & txtNombresR.value & _
						"&Apellidos=" & txtApellidosR.value & _
						"&Direccion=" & txtDireccionR.value & _
						"&AreaFono=" & txtAreaFonoR.value & _
						"&PaisFono=" & txtPaisFonoR.value & _
						"&Fono=" & txtFonoR.value & _
						"&Ciudad=" & Trim(txtCiudadR.value) & _
						"&Pais=" & txtPaisR.value & _
						"&Codigo=" & txtGiro.value & _					
						"&CiudadB=" & txtCiudad.value & _
						"&PaisB=" & txtPais.value & _
						"&NombresB=" & txtNombres.value & _
						"&ApellidosB=" & txtApellidos.value & _
						"&RutB=" & sIdB & _						
						"&DireccionB=" & txtDireccion.value & _
						"&FonoB=(" & txtPaisFono.value & txtAreaFono.value &  ") " & txtFono.value & _
						"&Mensaje=" & txtMensaje.value & _
						"&Monto=" & FormatNumber(trim(txtMonto.value), nDec) & _
						"&NombreRetira=" & frmGiro.txtNombresRetira.value & " " & frmGiro.txtApellidosRetira.value & _
						"&RutRetira=" & sIdR & "&Prefijo=" & frmGiro.txtMonedaPago.Value & _						
						"&Invoice=" & txtInvoice.value & _ 
						"&Captador=" & txtCaptador.value & _
						"&Pagador=" & txtPagador.value _
						) 
				
	End Sub
	
-->
</script>
<!--#INCLUDE virtual="/Compartido/Rutinas.htm" -->
</html>
