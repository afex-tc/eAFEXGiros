<%@ Language=VBScript %>
<%
	Response.Expires = 0
	Response.buffer = true	
	If Session("CodigoCliente") = "" Then
		response.Redirect "../Compartido/TimeOut.htm"
		response.end
	End If
%>
<!--#INCLUDE virtual="/Compartido/Constantes.asp" -->
<!--#INCLUDE virtual="/Compartido/Errores.asp" -->
<!--#INCLUDE virtual="/Compartido/Rutinas.asp" -->
<!--#INCLUDE virtual="/Compartido/Giros.asp" -->
<%
	'Variables de módulo
	'Variables para encabezado	
	Dim sEncabezadoFondo
	Dim sEncabezadoTitulo
	
	Dim sTitulo
	Dim sCliente
	Dim sCaptador, sPagador
	Dim nTipo	
	Dim afxGiro
	Dim rGiro, sDesde, sHasta, sNombre, sApellido, nTipoLlamada
	Dim rGiro2, rGiro3, rGiro4, nAgentes, nTelefono, nTelefono2, nPagina
	Dim sRut, sPasaporte, sMoneda, sTipoGiro, nEstadoGiro
	
	'Rescata parámetros
	sTitulo = Trim(Request("Titulo"))
	nTipo = cInt(0 & Request("Tipo"))
	sCliente = Trim(Request("Cliente"))
	sCaptador = Request("Captador")
	sPagador = Request("Pagador")
	sDesde = Request("Desde")
	sHasta = Request("Hasta")
	sNombre = Trim(Request("NombreCliente"))
	sApellido = Trim(Request("ApellidoCliente"))
	nTipoLlamada = cInt(0 & request("TipoLlamada"))
	sDesde = Request("Desde")
	sHasta = Request("Hasta") 
	nTelefono = cCur(0 & Request("Telefono"))
	nTelefono2 = cCur(0 & Request("Telefono2"))
	If nTelefono2 > 0 Then
		nTelefono = nTelefono2
	End If
	nAgentes = 1
	nPagina = cInt(0 & Request("Pagina"))
	sRut = Request("Rut")
	sPasaporte = Request("Pasaporte")
	sMoneda = Request("mn")
	sTipoGiro = Request("tg")
	nEstadoGiro = cCur(0 & Replace(Request("st"), -1, 0))
	
	If nEstadoGiro = 0 Then
		nEstadoGiro = -1
	End If
	'**Dependiendo del tipo de pantalla, rescata la lista	
	On Error Resume Next
	
	Set afxGiro = Server.CreateObject("AfexGiro.Giro")
	Dim sListaCiudad, sListaAgente, bLista
	sListaCiudad = ""
	sListaAgente = ""
	bLista = False
	'MostrarErrorMS Session("afxgirEnvPen")&"-"& ntipo
	
	Select Case nTipo
		Case afxGirosPendientes
			If sTitulo = "" Then sTitulo = "Giros Pendientes de Pago"
			If sCliente <> "" Then
				Set rGiro = ListaCompleta(Session("afxCnxAFEXpress"), afxGirosPendientes, True, "" , "", sCliente, "", "", -1, "", "", 0, "", "", "", 0, "", "", "")
			ElseIf Session("CiudadMatriz") = Session("CiudadCliente") Then
				Set rGiro = ListaCompleta(Session("afxCnxAFEXpress"), afxGirosPendientes, False, "", Session("CodigoMatriz"), "", sNombre, sApellido, -1, "", "", 0, "", "", "", nTelefono, sRut, sPasaporte, "")
			ElseIf InStr(1, "KNA;VAP;QPE", Session("CiudadCliente")) <> 0 Then
				bLista = True
				sListaCiudad = "KNA;VAP;QPE"
				sListaAgente = "AV;AA;AQ;AB"
				'Set rGiro = afxGiro.Lista(Session("afxCnxAFEXpress"), afxGirosPendientes,afxSi,, ,,,,,,,,,Session("CodigoMatriz"))
				Select Case nPagina
				Case 1
				Case 2
				Case 3
				Case 4
				End Select 
				Set rGiro = ListaCompleta(Session("afxCnxAFEXpress"), afxGirosPendientes, False, "", "AV', 'AB', 'AA', 'AQ", "", sNombre, sApellido, -1, "", "", 0, "", "", "", nTelefono, sRut, sPasaporte, "")
				'Set rGiro2 = ListaCompleta(Session("afxCnxAFEXpress"), afxGirosPendientes, False, "", "AB", "", sNombre, sApellido, -1, "", "", 0, "", "", "", nTelefono, sRut, sPasaporte)
				'Set rGiro3 = ListaCompleta(Session("afxCnxAFEXpress"), afxGirosPendientes, False, "", "AA", "", sNombre, sApellido, -1, "", "", 0, "", "", "", nTelefono, sRut, sPasaporte)
				'Set rGiro4 = ListaCompleta(Session("afxCnxAFEXpress"), afxGirosPendientes, False, "", "AQ", "", sNombre, sApellido, -1, "", "", 0, "", "", "", nTelefono, sRut, sPasaporte)
				nAgentes = 1
			ElseIf Session("Categoria") = 4 Then
				Set rGiro = ListaCompleta(Session("afxCnxAFEXpress"), afxGirosPendientes, afxSi, "", Session("CodigoAgente"), "", sNombre, sApellido, -1, "", "", 0, "", "", "", 0, "", "", "")
			Else
				Set rGiro = ListaCompleta(Session("afxCnxAFEXpress"), afxGirosPendientes, afxSi, "", Session("CodigoAgente"), "", sNombre, sApellido, -1, "", "", 0, "", "", "", 0, "", "", "")
			End If

		Case afxGirosAviso
			If sTitulo = "" Then sTitulo = "Giros Pendientes de Aviso"
			If sCliente <> "" Then 
				Set rGiro = ListaCompleta(Session("afxCnxAFEXpress"), afxGirosAviso, False, "" , "", sCliente, "", "", -1, "", "", 0, "", "", "", 0, "", "", "")
			ElseIf Session("CiudadMatriz") = Session("CiudadCliente") Then
				Set rGiro = ListaCompleta(Session("afxCnxAFEXpress"), afxGirosAviso, False, "", Session("CodigoMatriz"), "", "", "", -1, "", "", 0, Session("CodigoMoneyBrokerInt"), "", "", "", "", "","")
			ElseIf InStr(1, "KNA;VAP;QPE", Session("CiudadCliente")) <> 0 Then
				bLista = True
				sListaCiudad = "KNA;VAP;QPE"
				sListaAgente = "AV;AA;AQ;AB"
				Set rGiro = ListaCompleta(Session("afxCnxAFEXpress"), afxGirosAviso, False, "", sPagador, "", "", "", -1, "", "", 0, "", "", "", "", "", "","")
				nAgentes = 1
			Else
				Set rGiro = ListaCompleta(Session("afxCnxAFEXpress"), afxGirosAviso, False, "", Session("CodigoAgente"), "", "", "", -1, "", "", 0, "", "", "", "", "", "","")
			End If

		Case afxGirosReiteraAviso
			If sTitulo = "" Then sTitulo = "Giros Reiteración de Aviso"
			If sCliente <> "" Then 
				Set rGiro = ListaCompleta(Session("afxCnxAFEXpress"), afxGirosReiteraAviso, False, "", "", "", "", "", -1, "", "", 0, "", "", "", "", "", "","")
			ElseIf Session("CiudadMatriz") = Session("CiudadCliente") Then
				Set rGiro = ListaCompleta(Session("afxCnxAFEXpress"), afxGirosReiteraAviso, False, "", "", "", "", "", -1, "", "", 0, "", "", "", "", "", "","")
			ElseIf InStr(1, "KNA;VAP;QPE", Session("CiudadCliente")) <> 0 Then
				bLista = True
				sListaCiudad = "KNA;VAP;QPE"
				sListaAgente = "AV;AA;AQ;AB"
				Set rGiro = ListaCompleta(Session("afxCnxAFEXpress"), afxGirosReiteraAviso, False, "", sPagador, "", "", "", -1, "", "", 0, "", "", "", "", "", "","")
				nAgentes = 1
			Else
				Set rGiro = ListaCompleta(Session("afxCnxAFEXpress"), afxGirosReiteraAviso, False, "", Session("CodigoAgente"), "", "", "", -1, "", "", 0, "", "", "", "", "", "","")
			End If
			
		Case afxGirosRecibidos
			If nEstadoGiro = afxEstadoGiroNulo Then 
				sTitulo = "Giros Recibidos y Anulados"
			ElseIf sTitulo = "" Then 
				sTitulo = "Giros Recibidos"
			End If
			If sMoneda="" Then sMoneda = Session("MonedaExtranjera")
			If sTipoGiro="" Then sTipoGiro = "1"
			If sCliente <> "" Then
				If sDesde = "" Then sDesde = cDate("01-01-2004")
				If sHasta = "" Then sHasta = Date()	
				'Response.Write "ggg"			
				Set rGiro = ListaCompleta(Session("afxCnxAFEXpress"), afxGirosRecibidos, True, "", "", sCliente, "", "", nEstadoGiro, cDate(sDesde), cDate(sHasta), 0, "", "", "", "", "", "","")
			Else
				If sDesde = "" Then sDesde = Date() 
				If sHasta = "" Then sHasta = Date()
				If Trim(sCaptador) = "" Then
					Set rGiro = ListaCompleta(Session("afxCnxAFEXpress"), afxGirosRecibidos, True, sCaptador, sPagador, _
											  sCliente, sNombre, sApellido, nEstadoGiro, cDate(sDesde), cDate(sHasta), 0, _
											  Session("CodigoMGPago"), "", "", "", "", "","")
				Else
					Set rGiro = ListaCompleta(Session("afxCnxAFEXpress"), afxGirosRecibidos, True, sCaptador, sPagador, sCliente, sNombre, sApellido, nEstadoGiro, cDate(sDesde), cDate(sHasta), 0, "", "", "", "", "", "","")
				End If
			End If

		Case afxGirosEnviados
			If nEstadoGiro = afxEstadoGiroNulo Then
				sTitulo = "Giros Enviados y Anulados"
			ElseIf sTitulo = "" Then 
				sTitulo = "Giros Enviados"
			End If
			If sMoneda="" Then sMoneda = Session("MonedaExtranjera")
			If sTipoGiro="" Then sTipoGiro = "1"
			If sCliente <> "" Then
				If sDesde = Empty Then
					sDesde = Date() - 180
				End If
				If sHasta = "" Then sHasta = Date()
				Set rGiro = ListaCompleta(Session("afxCnxAFEXpress"), afxGirosEnviados, True, "", "", sCliente, sNombre, sApellido, nEstadoGiro, cDate(sDesde), cDate(sHasta), 0, "", "", "", "", "", "",sTipoGiro)
			Else
				If sDesde = "" Then sDesde = Date() 
				If sHasta = "" Then sHasta = Date()
				
				If Trim(sPagador) = "" Then										
					Set rGiro = ListaCompleta(Session("afxCnxAFEXpress"), afxGirosEnviados, True, Session("CodigoAgente"), "", "", sNombre, sApellido, nEstadoGiro, cDate(sDesde), cDate(sHasta), 0, "", Session("CodigoMGEnvio"), "", "", "", "" ,sTipoGiro)
				Else
					Set rGiro = ListaCompleta(Session("afxCnxAFEXpress"), afxGirosEnviados, True, Session("CodigoAgente"), sPagador, "", sNombre, sApellido, nEstadoGiro, cDate(sDesde), cDate(sHasta), 0, "", "", "", "", "", "",sTipoGiro)
				
					
				End If
			End If

		Case afxGirosCartola
			If sTitulo = "" Then sTitulo = "Movimientos del Cliente"
			'If sMoneda="" Then sMoneda = Session("MonedaExtranjera")
			'If sTipoGiro="" Then sTipoGiro = "1"
			If sDesde = "" Then sDesde = Date() - 180
			If sHasta = "" Then sHasta = Date()
			Set rGiro = ListaCompleta(Session("afxCnxAFEXpress"), afxGirosCartola, True, "", "", sCliente, "", "", -1, cDate(sDesde), cDate(sHasta), 0, "", "", "", "", "", "","")
	
		Case afxListaGirosCodigo
			If sTitulo = "" Then sTitulo = "Lista de Giros"
			'Set rGiro = afxgiro.Buscar(session("afxCnxAFEXpress"), Request("Giro"))
			Set rGiro = ListaCompleta(Session("afxCnxAFEXpress"), afxListaGirosCodigo, True, "", "", "", "", "", -1, "", "", "", "", "", Request("Giro"), "", "", "","")
		
		Case Ccur(Session("afxTipoGirEnvPen"))
		  'mostrarErrorms Session("CodigoAgente")
			sTitulo = "Giros Enviados y Pendientes"	
			'mostrarerrorMS "holasss"		
			If sMoneda="" Then sMoneda = Session("MonedaExtranjera")
			If sTipoGiro="" Then sTipoGiro = "1"
			If sCliente <> "" Then
				If sDesde = Empty Then
					sDesde = Date() - 180
				End If
				If sHasta = "" Then sHasta = Date()
				Set rGiro = ListaCompleta(Session("afxCnxAFEXpress"), Session("afxTipoGirEnvPen"), False, "", "", sCliente, sNombre, sApellido, nEstadoGiro, cDate(sDesde), cDate(sHasta), 0,Session("CodigoAgente"), "", "", "", "", "","")
				
			Else
				If sDesde = "" Then sDesde = Date() 
				If sHasta = "" Then sHasta = Date()
				
				If Trim(sCaptador) = "" Then										
				Set rGiro = ListaCompleta(Session("afxCnxAFEXpress"), Session("afxTipoGirEnvPen"), False, sCaptador, sPagador, _
											  sCliente, sNombre, sApellido, nEstadoGiro, cDate(sDesde), cDate(sHasta), 0, _
											  Session("CodigoAgente"), "", "", "", "", "","")
				Else
					Set rGiro = ListaCompleta(Session("afxCnxAFEXpress"), Session("afxTipoGirEnvPen"), False, sCaptador, sPagador, sCliente, sNombre, sApellido, nEstadoGiro, cDate(sDesde), cDate(sHasta), 0, Session("CodigoAgente"), "", "", "", "", "")
				End If
			End If				        
	
		' Jonathan miranda G. 27-02-2007
		case ccur(98)
			sTitulo = "Giros Recibidos no Editados"
			
			If sMoneda="" Then sMoneda = Session("MonedaExtranjera")
			If sTipoGiro="" Then sTipoGiro = "1"
			
			Set rGiro = ListaCompleta(Session("afxCnxAFEXpress"), 98, false, sCaptador, _
									  sPagador, sCliente, sNombre, sApellido, nEstadoGiro, "", _
									  "", 0, "", "", "", "", "", "","")
			
		'------------------------ Fin --------------------------
	End Select
	
	If Err.number <> 0 Then
		Set rGiro = Nothing
		Set afxGiro = Nothing
		MostrarErrorMS "" 
	End If
	
	sEncabezadoFondo = "Consultas"
	sEncabezadoTitulo = sTitulo '& " - (sólo 2004)"
	
'***** Permite ir al detalle del giro si sólo se encuentra un giro y existe el cliente y
'***** el informe es de	Giros Pendientes de Pago
	
	If Not rGiro.EOF Then
		If nAgentes = 1 And rGiro.RecordCount = 1 And ((sCliente <> "" And nTipo = afxGirosPendientes) Or nTipo = afxListaGirosCodigo) Then
			Dim sGiro			
			sGiro = rGiro("codigo_giro")
			Set rGiro = Nothing
			Set afxGiro = Nothing					
			Response.Redirect "DetalleGiro.asp?Codigo=" & sGiro & "&TipoLlamada=" & nTipoLlamada & "&TipoLista=" & nTipo & "&Cliente=" & sCliente & "&AFEXpress=" & Request("AFEXpress") & "&AFEXchange=" & Request("AFEXchange")
		End If
	End If
	'**
	
	
	Public Function ListaCompleta(ByVal Conexion, ByVal TipoLista, ByVal NoEditado, ByVal Captador, _
								  ByVal Pagador, ByVal CodigoCliente, ByVal NombreCliente, _
								  ByVal ApellidoCliente, ByVal Estado, ByVal FechaDesde, _
								  ByVal FechaHasta, ByVal Registros, ByVal ExcluirCaptador, _
								  ByVal ExcluirPagador, ByVal CodigoGiro, ByVal Telefono, _
								  ByVal Rut, ByVal Pasaporte, ByVal TipoGiro )
	   Dim sSQL
	   Dim sPuntero		


	   'Manejo de errores
	   On Error Resume Next
		Set ListaCompleta = Nothing   
	   'Crea la consulta
	   sSQL = "SELECT  DISTINCT  "
	   
	   If cInt(0 & Registros) > 0 Then
	      sSQL = sSQL & "TOP " & Registros & " "
	   End If
	  

	   sSQL = sSQL & "codigo_giro, fecha_captacion, agente_captador, agente_pagador, codigo_beneficiario, comision_pesos, " & _
					 "codigo_remitente, ciudad_beneficiario, ciudad_remitente, fono_beneficiario, invoice, " & _
					 "correlativo_salida, monto_giro, comision_captador, comision_pagador, estado_giro, estado, " & _
					 "sw_editado, beneficiario, remitente, prefijo_moneda, g.codigo_moneda, moneda, numero_documento, " & _
					 "case codigo_moneda when 'USD' then (monto_giro*g.tipo_cambio) when 'CLP' then monto_giro end pesos, g.tipo_cambio, " & _
					 " codigo_region, tipo_giro, tarifaclpafex, tarifamoneygram, monto_iva, isnull(gastos_transferencia, 0) as gastos_transferencia, " & _
					 " tarifa_cobrada, pais_beneficiario, pais_remitente, fecha_pago " 
		
		If TipoLista = afxTipoLista.afxGirosEnviados and sPagador ="ME"  Then
			If session("Categoria") = 3 Then
				sSQL = sSQL & ", 0 as diferencia  "	
			else
			
				sSQL = sSQL & ",( select vi.haber_nacional from vvoucher vi " & _
							  " where vi.detalle_voucher = 'Recepción de Giro' and " & _
							  " vi.estado_voucher <> 9 and vi.numero_linea = 3 and vi.codigo_operacion=g.codigo_giro ) as diferencia "
			End if
		End IF
		
		sSQL = sSQL & "FROM      VGiro g with(nolock) " 
	          
	

		if Ccur(TipoLista) = ccur(98) then
			sSQL = sSQL & " WHERE     (sw_editado = 0 "
		else
			sSQL = sSQL & " WHERE     (sw_editado = 1 " 
		end if

	          
	   If NoEditado Then
	      sSQL = sSQL & " OR sw_editado = 0 "
	   End If
	   sSQL = sSQL & " ) "
		
	   'Verifica el tipo de lista que desea
	Select Case Ccur(TipoLista)		
		' Jonathan Miranda G. 27-02-2007
		case ccur(98)	

			If NombreCliente <> Empty Then
				sSQL = sSQL & " AND nombre_beneficiario like '%" & NombreCliente & "%' "
			End If
			If ApellidoCliente <> Empty Then
               sSQL = sSQL & " AND apellido_beneficiario like '%" & ApellidoCliente & "%' "
            End If
		'----------------- Fin ---------------------------


		Case Ccur(Session("afxTipoGirEnvPen"))				
			If CodigoCliente <> Empty Then
				sSQL = sSQL & " AND codigo_remitente = " & EvaluarStr(CodigoCliente)
			Else
				If NombreCliente <> Empty Then
					sSQL = sSQL & " AND nombre_remitente like '%" & NombreCliente & "%' "
				End If
				If ApellidoCliente <> Empty Then
					sSQL = sSQL & " AND apellido_remitente like '%" & ApellidoCliente & "%' "
	            End If
	            
	        End If
	        ' verifica estado
	        If Estado <> -1 Then
				sSQL = sSQL & " AND estado_giro = " & Estado
			End If
			
			If Captador = "*"  Then
			
			ElseIf Session("CodigoAgente") <> "" Then 
				'sSQL = sSQL & " AND agente_pagador IN (CASE WHEN tipo_giro = 0 THEN agente_pagador ELSE " & EvaluarStr(Pagador) & " END)"
				sSQL = sSQL & " AND agente_Captador = (" & EvaluarStr(Session("CodigoAgente")) & ")"
			End If			
			
			If Pagador <> "*" then
				sSql = sSql & " and agente_pagador= " & EvaluarStr(Pagador )
			end if
			   
	    Case afxTipoLista.afxGirosAviso ' giros que pueden ser pagados o avisados			

	         sSQL = sSQL & " AND estado_giro in(" & afxEstadoGiro.afxGiroCaptacion & ", " & _
	                                                afxEstadoGiro.afxGiroEnvio & ") " & _
	                        "AND (numero_aviso = 0 OR numero_aviso is null)"
				'sSQL = sSQL & " AND codigo_region = CASE WHEN tipo_giro = 0 THEN " & cInt(0 & Session("RegionAgente")) & " ELSE codigo_region END "
				If Pagador = "*"  Then
				ElseIf Pagador <> "" Then 
					sSQL = sSQL & " AND agente_pagador IN (CASE WHEN tipo_giro = 0 And " & Session("Categoria") & " <> 4 THEN agente_pagador ELSE " & EvaluarStr(Pagador) & " END)"
				End If
		      
	    Case afxTipoLista.afxGirosPendientes ' giros que pueden ser pagados o avisados
			 Dim bRegion 			

	         sSQL = sSQL & " AND estado_giro in(" & afxEstadoGiro.afxGiroAviso & ", " & _
	                                                afxEstadoGiro.afxGiroCaptacion & ", " & _
	                                                afxEstadoGiro.afxGiroEnvio & ", " & _
	                                                afxEstadoGiro.afxReclamo & ") "
	         ' verifica cliente
	         bRegion = false
	         If CodigoCliente <> Empty Then
	            sSQL = sSQL & " AND codigo_beneficiario = " & EvaluarStr(CodigoCliente)
		         bRegion = False
	         ElseIf Rut <> "" Then
	            Rut = Replace(Rut, ".", "")
	            Rut = Replace(Rut, "-", "")
	            Rut = Right("000000000" & Rut, 9)
	            sSQL = sSQL & " AND rut_beneficiario = " & EvaluarStr(Rut)
		         bRegion = False   
	         ElseIf Pasaporte <> "" Then
	            sSQL = sSQL & " AND pasaporte_beneficiario = " & EvaluarStr(Pasaporte)
	            bRegion = False
	         ElseIf Telefono > 0 Then
					sSQL = sSQL & " AND (fono_beneficiario = " & Telefono & " or fono2_beneficiario = " & Telefono & ") "
					bRegion = False	
	         Else
					If NombreCliente <> Empty Then
						sSQL = sSQL & " AND nombre_beneficiario like '%" & NombreCliente & "%' "
						bRegion = False
					End If
	            If ApellidoCliente <> Empty Then
	               sSQL = sSQL & " AND apellido_beneficiario like '%" & ApellidoCliente & "%' "
	               bRegion = False
	            End If
	         End If
		      
		     If bRegion Then
					'sSQL = sSQL & " AND codigo_region = CASE WHEN tipo_giro = 0 THEN " &  cInt(0 & Session("RegionAgente")) & " ELSE codigo_region END AND pais_beneficiario = '" & Session("PaisMatriz") & "' "
			 Else
			 End If
			 If Pagador = "*"  Then
			 
			 ElseIf Pagador <> "" Then 
				If Session("Categoria") <> 4 Then
					'sSQL = sSQL & " AND agente_pagador IN (CASE WHEN tipo_giro = 0 THEN agente_pagador ELSE " & EvaluarStr(Pagador) & " END)"
			 		sSQL = sSQL & " AND (agente_pagador IN (" & EvaluarStr(Pagador) & ") OR ( tipo_giro = 0 AND agente_pagador = '" & Session("CodigoMatriz") & "'))"
			 	Else
			 		sSQL = sSQL & " AND (agente_pagador IN (" & EvaluarStr(Pagador) & "))"
			 	End If
			 End If
		
		Case afxTipoLista.afxGirosReiteraAviso  'giros que pueden ser pagados o avisados
	         sSQL = sSQL & " AND estado_giro in(" & afxEstadoGiro.afxGiroAviso & ") " & _
	                        "AND numero_aviso <> 0 AND fecha_ultaviso < '" & Date & "'"
				'sSQL = sSQL & " AND codigo_region = CASE WHEN tipo_giro = 0 THEN " & cInt(0 & Session("RegionAgente")) & " ELSE codigo_region END "
				If Pagador = "*"  Then
				ElseIf Pagador <> "" Then 
					sSQL = sSQL & " AND agente_pagador IN (CASE WHEN tipo_giro = 0 THEN agente_pagador ELSE " & EvaluarStr(Pagador) & " END)"
				End If
		
		                
		Case afxTipoLista.afxGirosEnviados     ' giros enviados por el agente, cliente o ambos
	         ' verifica cliente
			If CodigoCliente <> Empty Then
				sSQL = sSQL & " AND codigo_remitente = " & EvaluarStr(CodigoCliente)
	        Else
				If NombreCliente <> Empty Then
					sSQL = sSQL & " AND nombre_remitente like '%" & NombreCliente & "%' "
				End If
				If ApellidoCliente <> Empty Then
					sSQL = sSQL & " AND apellido_remitente like '%" & ApellidoCliente & "%' "
	            End If
	        End If
	        ' verifica estado
	        If Estado <> -1 Then
				sSQL = sSQL & " AND estado_giro = " & Estado
			End If
			
			If Pagador = "*"  Then
			
			ElseIf Pagador <> "" Then 
				'sSQL = sSQL & " AND agente_pagador IN (CASE WHEN tipo_giro = 0 THEN agente_pagador ELSE " & EvaluarStr(Pagador) & " END)"
				sSQL = sSQL & " AND agente_pagador IN (" & EvaluarStr(Pagador) & ")"
				
			End If
		      
	    Case afxTipoLista.afxGirosRecibidos    ' giros recibidos por el agente, cliente o ambos
			' verifica cliente
	        If CodigoCliente <> Empty Then
				sSQL = sSQL & " AND codigo_beneficiario = " & EvaluarStr(CodigoCliente)
	        Else
				If NombreCliente <> Empty Then
					sSQL = sSQL & " AND nombre_beneficiario like '%" & NombreCliente & "%' "
				End If
				If ApellidoCliente <> Empty Then
	               sSQL = sSQL & " AND apellido_beneficiario like '%" & ApellidoCliente & "%' "
	            End If
	            sSQL = sSQL & " AND TIPO_GIRO = " & sTipoGiro ' agregado PSS 25-09-2009
	        End If
	        ' verifica estado
	        If Estado <> -1 Then
				sSQL = sSQL & " AND estado_giro = " & Estado
			Else
				sSQL = sSQL & " AND estado_giro in (5, 6) "
			End If
			If Pagador <> Empty Then
				sSQL = sSQL & " AND agente_pagador = " & EvaluarStr(Pagador)
			End If
		      
	    Case afxTipoLista.afxGirosCartola      ' giros enviados y recibidos por el cliente
	         sSQL = sSQL & " AND (codigo_beneficiario = " & EvaluarStr(CodigoCliente) & _
	                       "  OR codigo_remitente = " & EvaluarStr(CodigoCliente) & ") "
		                       
	    Case afxTipoLista.afxGirosCodigo
			' JFMG 01-04-2011 se cometa para buscardirecto en un procedimiento almacendado
	         'sSQL = sSQL & " AND (codigo_giro = " & EvaluarStr(CodigoGiro)	 
	         
	        '' JFMG 13-12-2010 verifica si el código escrito puede que sea de XOOM (9 digitos con un 9 al comienzo)
			''If Len(CodigoGiro) = 9 and Left(CodigoGiro, 1) = "9" Then
			'	sSQL = sSQL & " or invoice = " & EvaluarStr(CodigoGiro)			
			''End If
			'if isnumeric(CodigoGiro) then 
			'	sSQL = sSQL & " or g.codigo_giro " & _
			'			" in(SELECT g.codigo_giro " & _
			'			" FROM giro.codigoautomatico cg with(nolock) " & _
			'			" WHERE cg.codigoautomatico = " & CodigoGiro & ")"	' si se elimina esta linea también se debe eliminar la apertura del parentesis en la línea de más arriba
			'end if
			'sSQL = sSQL & ") "
			'' FIN JFMG 13-12-2010
			
			sSQL = "exec MostrarDatosGiroCodigo " & evaluarstr(CodigoGiro)
			' FIN JFMG 01-04-2011
			
		'Response.Write ssql
		'Response.End 
	End Select
	'	response.Write sSQL
	'	response.End 
	
		' JFMG 01-04-2011
		IF ccur(TipoLista) <> afxTipoLista.afxGirosCodigo Then
		' FIN 01-04-2011
	
			'Verifica agente
			If Captador = "*"  Then
			ElseIf Captador <> "" Then
				sSQL = sSQL & " AND agente_captador IN  (" & EvaluarStr(Captador) & ") "			
				If Ucase(Trim(Captador)) = Ucase(Trim(Session("CodigoTXPago"))) Then
					sSQL = sSQL & " AND left(codigo_giro, 2) = 'MT' "
				End If
			End If
			
			If Captador = "*" Then
			ElseIf (ExcluirCaptador <> Empty)  Then
			
				sSQL = sSQL & " AND agente_captador <> 'MB' "
				sSQL = sSQL & " AND left(codigo_giro, 2) <> 'MT' "
			End If
			If Pagador = "*"  Then
			ElseIf ExcluirPagador <> Empty Then	
				sSQL = sSQL & " AND agente_pagador <> " & EvaluarStr(ExcluirPagador)
			End If

			iF AfxTipoLista.afxGirosCartola = TipoLista or AfxTipoLista.AfxGirosEnviados = TipoLista  Then ' codigo nuevo PSS 21-09-2009
			else	
				If sMoneda="***" Or sMoneda="" Then
				Else
					sSQL = sSQL & " AND G.codigo_moneda=" & EvaluarStr(sMoneda)
					
				End If
			end if
			
			If TipoLista <> afxTipoLista.afxGirosRecibidos and TipoLista <> afxTipoLista.afxGirosPendientes and TipoLista <> Session("afxTipoGirEnvPen") Then
				If nEstadoGiro = -1 Then 
					sSQL = sSQL & " AND estado_giro <> 9 "
				End If
			elseif TipoLista = Session("afxTipoGirEnvPen")then
				sSQL = sSQL & " AND estado_giro not in (2,3,5,6,9) "
			End If
		'	Response.Write tipogiro
			' codigo nuevo PSS 06-01-2010 para mostrar solo los giros internacionales
			If afxTipoLista.afxGirosEnviados= TipoLista Then
				If Session("Categoria") = 4 Then
				Else
					If TipoGiro = 0 then
						sSql = sSql & " and categoria_captador in (1,2,3) and categoria_pagador in (0,1,2,3) " 
					elseif TipoGiro= 1 then
						sSql = ssql & " and categoria_captador in (1,2,3) and categoria_pagador = 4 " 
					end If
				End if 
			end if				
			'Verifica las fechas
			If FechaDesde <> Empty Then
				If nEstadoGiro = afxEstadoGiroNulo Then
					sSQL = sSQL & " AND convert(char,fecha_anulacion,112) between " & EvaluarStr(formatofechasql(FechaDesde)) & " AND " & EvaluarStr(formatofechasql(FechaHasta))
					sSQL = sSQL & " AND fecha_giro <> fecha_anulacion"
				ElseIf TipoLista = afxTipoLista.afxGirosRecibidos And CodigoCliente = "" And (Session("Categoria")=1 Or Session("Categoria")=2) Then
				
					sSQL = sSQL & " AND convert(char,fecha_pago,112) >= " & EvaluarStr(formatofechasql(FechaDesde))
					If FechaHasta <> Empty Then
						sSQL = sSQL & " AND convert(char,fecha_pago,112) <= " & EvaluarStr(formatofechasql(FechaHasta))
					End If
					sSQL = sSQL & " order by g.codigo_moneda desc, fecha_pago desc, codigo_giro desc "
					  
				ElseIf TipoLista = afxTipoLista.afxGirosRecibidos And Session("Categoria")=3 Then
					'If sTipoGiro="0" Then
						sSQL = sSQL & " AND convert(char,fecha_pago,112) >= " & EvaluarStr(formatofechasql(FechaDesde))
						If FechaHasta <> Empty Then
							sSQL = sSQL & " AND convert(char,fecha_pago,112) <= " & EvaluarStr(formatofechasql(FechaHasta))
						End If
						sSQL = sSQL & " order by g.codigo_moneda desc, fecha_pago desc, codigo_giro desc "					
					
				Else
					sSQL = sSQL & " AND convert(char,fecha_captacion,112) >= " & EvaluarStr(formatofechasql(FechaDesde))
					If FechaHasta <> Empty Then
						sSQL = sSQL & " AND convert(char,fecha_captacion,112) <= " & EvaluarStr(formatofechasql(FechaHasta))
						
					End If
					
					If Session("Categoria") = 3 Or Session("Categoria") = 4 Then
						sSQL = sSQL & " order by g.codigo_moneda desc, fecha_captacion desc, correlativo_salida desc, invoice desc, codigo_giro desc "
					Else
						sSQL = sSQL & " order by g.codigo_moneda desc, fecha_captacion desc, codigo_giro desc "
						
					End If
					
				End If
			Else
			  
				If Session("Categoria") = 3 Or Session("Categoria") = 4 Then
					sSQL = sSQL & " order by g.codigo_moneda desc, fecha_captacion desc, correlativo_salida, codigo_giro desc "
				Else
					sSQL = sSQL & " order by g.codigo_moneda desc, fecha_captacion desc, codigo_giro desc "
				End If
			End If

		' JFMG 01-04-2011
		End If
		' FIN JFMG 01-04-2011
	
		''Asigna al metodo el resultado de la consulta		
		'response.Write sSQL
		'response.End 
		Set ListaCompleta = EjecutarSQLCliente(Conexion, sSQL)		
		 
	   'Si se produjeron errores en la consulta
	   If Err.Number <> 0 Then
			MostrarErrorMS "Lista de Giros"
	  End If

	End Function

	Function FormatoFecha(Byval Fecha)
		fecha=cDate(fecha)
		FormatoFecha = Year(Fecha) & Right("0" & Month(Fecha), 2) & Right("0" & Day(Fecha), 2)
	End Function	


Public Function EvaluarStr(ByVal Valor)
	Dim Devuelve
		
  EvaluarStr = "'" & Valor & "'"

End Function
	
%>
<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<title>AFEX</title>
<link rel="stylesheet" type="text/css" href="../Estilos/Cliente.css">
<link rel="stylesheet" type="text/css" href="../Estilos/Reportes.css">
</head>
<script LANGUAGE="VBScript">
<!--
	Sub imgAceptar_onClick()
		
		If window.tbReporte.style.display = "" then
			window.tbReporte.style.display = "none"
		Else 
			window.tbReporte.style.display = ""
		End If
	End Sub		

	Function imgAceptar_onMouseOver()
		window.imgAceptar.style.cursor = "Hand"		
	End Function

	Sub window_onload()
		<% If nTipo <> afxGirosPendientes And nTipo <> afxGirosAviso _
			And nTipo <> afxGirosReiteraAviso And nTipo <> afxListaGirosCodigo And _
			nTipo <> 98 Then %>
				objConsulta.Desde = Replace("<%=sDesde%>", "/", "-")
				objConsulta.Hasta = Replace("<%=sHasta%>", "/", "-")
				objConsulta.Tipo = <%=Request("Tipo")%>
				objConsulta.CodigoCliente = "<%=sCliente%>"
				objConsulta.NombreCliente = "<%=sNombre%>"
				objConsulta.ApellidoCliente = "<%=sApellido%>"
				objConsulta.Captador = "<%=sCaptador%>"
				objConsulta.Pagador = "<%=sPagador%>"
				objConsulta.Moneda = "<%=sMoneda%>"
				objConsulta.TipoGiro = "<%=sTipoGiro%>"
		<% End If %>

	End Sub
	
	Sub Encabezado(ByVal Moneda)
		' JFMG 02-10-2008 se agrega la comparación
		if "<%=sPagador%>" = "ME" then
			' si el pagador es moneygram enviados agrega la columna de comision en Pesos
			document.write "<tr><td colspan=2 align=left style="" font-size: 12px; font-weight: bold"">" & Moneda & "</td></tr><tr CLASS=Encabezado><td WIDTH=100><b>Fecha</b></td><td WIDTH=300><b>Remitente</b></td><td WIDTH=300><b>Beneficiario</b></td><% If nTipo = afxGirosAviso Or nTipo = afxGirosReiteraAviso Or nTipo = afxGirosPendientes Then %><td WIDTH=100><b>Telefono</b></td><% End If %><td WIDTH=100><b>Codigo</b></td><% If (Session(Categoria) = 3 Or Session(Categoria) = 4) Then %><td WIDTH=80><b>Orden</b></td><% Else %><td WIDTH=80><b>Invoice</b></td><% End If %><td WIDTH=80><b>Nº Boleta</b></td><td WIDTH=200><b>Monto $ B/S</b></td><td WIDTH=10><b></b></td><td WIDTH=300><b>Monto</b></td><td WIDTH=200><b>Monto Pesos</b></td><td WIDTH=200><b>Dif.Cambio</b></td><td WIDTH=200><b>Comisión US$</b></td><td WIDTH=200><b>Comisión $</b></td><td WIDTH=100><b>Estado</b></td></tr>"
		else
			' JFMG 12-11-2008 se agregan los campos de boleta de servicio si son giros enviados
			<% If nTipo = afxGirosEnviados Then %>
				if "<%=sMoneda%>" = "USD" then
					document.write "<tr><td colspan=2 align=left style="" font-size: 12px; font-weight: bold"">" & Moneda & "</td></tr><tr CLASS=Encabezado><td WIDTH=100><b>Fecha</b></td><td WIDTH=300><b>Remitente</b></td><td WIDTH=300><b>Beneficiario</b></td><td WIDTH=100><b>Codigo</b></td><% If (Session(Categoria) = 3 Or Session(Categoria) = 4) Then %><td WIDTH=80><b>Orden</b></td><% Else %><td WIDTH=80><b>Invoice</b></td><% End If %><td WIDTH=80><b>Nº Boleta</b></td><td WIDTH=80><b>Monto $ B/S</b></td><td WIDTH=80><b>T/C</b></td><td WIDTH=80><b>Monto US$ B/S</b></td><td WIDTH=10><b></b></td><td WIDTH=300><b>Monto</b></td><td WIDTH=200><b>Monto Pesos</b></td><td WIDTH=200><b>Dif.Cambio</b></td><td WIDTH=200><b>Comision</b></td><td WIDTH=100><b>Estado</b></td></tr>"
				elseif "<%=sMoneda%>" = "CLP" then
					document.write "<tr><td colspan=2 align=left style="" font-size: 12px; font-weight: bold"">" & Moneda & "</td></tr><tr CLASS=Encabezado><td WIDTH=100><b>Fecha</b></td><td WIDTH=300><b>Remitente</b></td><td WIDTH=300><b>Beneficiario</b></td><td WIDTH=100><b>Codigo</b></td><% If (Session(Categoria) = 3 Or Session(Categoria) = 4) Then %><td WIDTH=80><b>Orden</b></td><% Else %><td WIDTH=80><b>Invoice</b></td><% End If %><td WIDTH=80><b>Nº Boleta</b></td><td WIDTH=80><b>Monto $ B/S</b></td><td WIDTH=10><b></b></td><td WIDTH=300><b>Monto</b></td><td WIDTH=200><b>Monto Pesos</b></td><td WIDTH=200><b>Dif.Cambio</b></td><td WIDTH=200><b>Comision</b></td><td WIDTH=100><b>Estado</b></td></tr>"
				end if
			<% else %>
				document.write "<tr><td colspan=2 align=left style="" font-size: 12px; font-weight: bold"">" & Moneda & "</td></tr><tr CLASS=Encabezado><td WIDTH=100><b>Fecha</b></td><td WIDTH=300><b>Remitente</b></td><td WIDTH=300><b>Beneficiario</b></td><% If nTipo = afxGirosAviso Or nTipo = afxGirosReiteraAviso Or nTipo = afxGirosPendientes Then %><td WIDTH=100><b>Telefono</b></td><% End If %><td WIDTH=100><b>Codigo</b></td><% If (Session(Categoria) = 3 Or Session(Categoria) = 4) And (nTipo <> afxGirosEnviados) Then %><td WIDTH=80><b>Orden</b></td><% Else %><td WIDTH=80><b>Invoice</b></td><% End If %><td WIDTH=80><b>Nº Boleta</b></td><td WIDTH=10><b></b></td><td WIDTH=300><b>Monto</b></td><td WIDTH=200><b>Monto Pesos</b></td><td WIDTH=200><b>Dif.Cambio</b></td><% If nTipo = afxGirosEnviados Or nTipo = afxGirosRecibidos Then %><td WIDTH=200><b>Comision</b></td><% End If %><td WIDTH=100><b>Estado</b></td></tr>"
			<% end if %>
			' *********************************** FIN ******************************
		end if
	End Sub
	
//-->
</script>
<body onMouseMove="window.status=''">
<!--#INCLUDE virtual="/Compartido/Encabezado.asp" -->
<!--#INCLUDE virtual="/Compartido/Rutinas.htm" -->
<table border="0" cellspacing="0" cellpadding="0" style="LEFT: 0px; POSITION: relative; TOP: 0px">
<tr>
	<td align="middle">
		<!--Si la página es para mostrar los giros pendientes del cliente, no se muestra el 
		filtro-->
		<%If nTipo <> afxGirosPendientes And nTipo <> afxGirosAviso _
			And nTipo <> afxGirosReiteraAviso And nTipo <> afxListaGirosCodigo And _
			nTipo <> 98 Then%>		
			<OBJECT id=objConsulta style="HEIGHT: 240px; LEFT: 0px; TOP: 0px; WIDTH: 544px" type=text/x-scriptlet width=544 VIEWASTEXT>
			<PARAM NAME="Scrollbar" VALUE="0">
			<PARAM NAME="URL" VALUE="http:ConfiguracionConsulta.asp"></OBJECT>
		<%End If%>
	</td>
</tr>
<tr height="10"><td>
	<table cellspacing="1" cellpadding="1" ID="tbReporte" border="0" ALIGN="center" STYLE="COLOR: #505050; FONT-FAMILY: Verdana; FONT-SIZE: 10px; POSITION: relative; TOP: 0px">
	<% If bLista And nTipo <> afxGirosPendientes Then %>
		<tr>
			<td width="100%"></td>
			<td colspan="7" align="right">
			<table>			
			<tr>
				<td><a href="ListaGiros.asp?Tipo=<%=afxGirosAviso%>&Pagador=<%="AV"%>">Viña</a></td>
				<td><a href="ListaGiros.asp?Tipo=<%=afxGirosAviso%>&Pagador=<%="AA"%>">Valparaiso</a></td>
				<td><a href="ListaGiros.asp?Tipo=<%=afxGirosAviso%>&Pagador=<%="AQ"%>">Quilpue</a></td>
				<td><a href="ListaGiros.asp?Tipo=<%=afxGirosAviso%>&Pagador=<%="AB"%>">Viña 2</a></td>
			<tr>
			</table>
		</td></tr>
	<% End If %>
		<%				

		Dim nTotal, sDetalle, nMonto, nComision, nTotalComision, nCantidad, nGiros, nMontoPesos, nTotalPesos, nDiferencia
		Dim nTotalDif
		Dim i, sColor, nDec, sMn
		Dim codigoGiroAnt 'MS 29-11-2013
		
		' JFMG 12-11-2008 campos de boleta de servicio
		dim nMontoBoletaPesos, nMontoBoletaDolares, nTotalBoletaPesos, nTotalBoletaDolares
		nTotalBoletaPesos = 0
		nTotalBoletaDolares = 0
		' ***************************** FIN ************************
		
		nTotal = 0
		nTotalPesos = 0
		nComision = 0
		nTotalComision = 0
		nCantidad = 0
		nGiros = 0
		codigoGiroAnt = "" 'MS 29-11-2013
		
		i = 0
	For i = 1 To nAgentes
	
		Do Until rGiro.EOF
			'MS 29-11-2013
			IF codigoGiroAnt = rGiro("codigo_giro") Then
				rGiro.MoveNext
				exit for
			End If	
			' FIN MS 29-11-2013				
			' JFMG 12-11-2008 campos de boleta de servicio
			nMontoBoletaPesos = 0
			nMontoBoletaDolares	= 0	
			' ***************************** FIN ************************

			If Not bLista Then
				sListaCiudad = rGiro("ciudad_beneficiario")
				sListaAgente = rGiro("agente_pagador")
				
			End If			
			
			If (InStr(1, sListaCiudad, rGiro("ciudad_beneficiario")) <> 0 And InStr(1, sListaAgente, rGiro("agente_pagador")) <> 0)  _
			Or  (InStr(1, sListaAgente, rGiro("agente_pagador")) <> 0 And Session("Categoria") = 4) _
			Or (sListaAgente = "ME") _
			Then	'Or (rGiro("tipo_giro") = 0 And rGiro("codigo_region") = cInt(0 & Session("RegionAgente"))) _
				If nTipo = afxGirosPendientes And sCliente <> "" Then
					nGiros = rGiro.RecordCount
				End If
				If rGiro("sw_editado") = 0 Then
					sColor="gray" 
				Else
					sColor=""
				End If
				If sMn <> rGiro("codigo_moneda") Then
			%>
					<script>Encabezado("<%=rGiro("moneda")%>")</script>
			<%
				End If
				sMn = rGiro("codigo_moneda")
				If rGiro("codigo_moneda") = Session("MonedaNacional") Then
					nDec = 0
				Else
					nDec = 2
				End If
				'response.Write rgiro("invoice")
				'response.End 
			
		%>		
				
				<a href="DetalleGiro.asp?Codigo=<%=rGiro("codigo_giro")%>&TipoLlamada=<%=nTipoLlamada%>&TipoLista=<%=nTipo%>&Cliente=<%=sCliente%>&AFEXpress=<%=Request("AFEXpress")%>&AFEXchange=<%=Request("AFEXchange")%>&GS=<%=nGiros%>" language="javascript"  onmouseout="window.status=''" onmouseover="window.status='Ver detalle'" onclick="">
				<tr style="HEIGHT: 25px; color: <%=sColor%>" language="javascript" onmouseover="javascript:this.bgColor='#a4dded'; window.status='<%=rGiro("codigo_giro")%>'" onmouseout="javascript:this.bgColor='#DAF6FF'; window.status=''" bgColor="#dbf7ff" style="cursor: hand" title="<%=rGiro("ciudad_beneficiario")%>">
					<td><%=rGiro("fecha_captacion")%></td>
					<td><%=rGiro("remitente")%></td>
					<td><%=rGiro("beneficiario")%></td>
					<% If nTipo = afxGirosAviso Or nTipo = afxGirosReiteraAviso  Or nTipo = afxGirosPendientes Then %>
						<td><%=rGiro("fono_beneficiario")%></td>
					<% End If %>
					<td><%=rGiro("codigo_giro")%></td>
					<% If (Session("Categoria") = 3 Or Session("Categoria") = 4) And (nTipo <> afxGirosEnviados) and  (nTipo <> Ccur(Session("afxTipoGirEnvPen")))Then %>
						<td><%=rGiro("correlativo_salida")%></td>
					<% Else %>
						<td><%=rGiro("invoice")%></td>					
					<% End If %>
					<td><%=rGiro("numero_documento")%></td>
					
					<!-- JFMG 12-11-2008 si son giros enviados se agregan datos de boleta de servicio -->
					<% If nTipo = afxGirosEnviados Then
						If sListaAgente = "ME" Then
							if not isnull(rGiro("tarifaclpafex")) then ' giros nuevos, solo pesos
							    if rGiro("codigo_moneda") = "USD" then
								    ' JFGM 11-07-2012 nMontoBoletaPesos = round(ccur(rGiro("tarifaclpafex")) - ((ccur(rGiro("tarifamoneygram"))*ccur(rGiro("tipo_cambio"))) * 0.7), 0)
								    nMontoBoletaPesos = formatnumber(ccur(rGiro("comision_pesos")), 0)
								else
								    nMontoBoletaPesos = round(ccur(rGiro("tarifaclpafex")) - (ccur(rGiro("tarifamoneygram")) * 0.7), 0)
							    end if
							else ' giros antiguos
								if rGiro("codigo_moneda") = "USD" then
									nMontoBoletaPesos = ccur(formatnumber(ccur(rGiro("monto_iva")) * ccur(rGiro("tipo_cambio")), 0))
									nMontoBoletaDolares = rGiro("monto_iva")
								elseif rGiro("codigo_moneda") = "CLP" then
									nMontoBoletaPesos = rGiro("monto_iva")
									nMontoBoletaDolares = 0
								end if
							end if
						else ' giros afex
							if rGiro("codigo_moneda") = "USD" then
								if rGiro("pais_beneficiario") <> "CL" then	' giro internacional
									nMontoBoletaPesos = ccur(formatnumber(ccur(rGiro("monto_iva")) * ccur(rGiro("tipo_cambio")), 0))
									nMontoBoletaDolares = rGiro("monto_iva")
								else		' giro nacional
									nMontoBoletaPesos = ccur(formatnumber(ccur(rGiro("tarifa_cobrada")) * ccur(rGiro("tipo_cambio")), 0))
									nMontoBoletaDolares = rGiro("tarifa_cobrada")
								end if
							elseif rGiro("codigo_moneda") = "CLP" then
								if rGiro("pais_beneficiario") <> "CL" then	' giro internacional
									nMontoBoletaPesos = rGiro("monto_iva")
									nMontoBoletaDolares = 0
								else		' giro nacional
									nMontoBoletaPesos = rGiro("tarifa_cobrada")
									nMontoBoletaDolares = 0
								end if
							end if
						end if
							
						nTotalBoletaPesos = ccur(nTotalBoletaPesos) + ccur(nMontoBoletaPesos)
						nTotalBoletaDolares = ccur(nTotalBoletaDolares) + ccur(nMontoBoletaDolares)
														
						if sMoneda = "USD" then
					%>
							<td align="right"><%=formatnumber(nMontoBoletaPesos,0)%></td>
							<td align="right"><%=formatnumber(rGiro("tipo_cambio"),2)%></td>
							<td align="right"><%=formatnumber(nMontoBoletaDolares,2)%></td>
						<%elseif sMoneda = "CLP" then%>
							<td align="right"><%=formatnumber(nMontoBoletaPesos,0)%></td>
						<%end if%>
					<% End If %>
					<!-- ********************************** FIN ************************************* -->					
					
					<td><%=rGiro("prefijo_moneda")%></td>
					<%
					If nTipo = afxGirosCartola Then
						If sCliente = Trim(rGiro("codigo_remitente")) Then 
							nMonto = cCur(rGiro("monto_giro")) * -1
						Else
							nMonto = cCur(0 & rGiro("monto_giro"))
						End If
						
					Else
						nMonto = cCur(0 & rGiro("monto_giro"))
					End If
					If sListaAgente = "ME" Then
						nMontoPesos= cCur(0 & rGiro("pesos"))
						nDiferencia = cCur(0 & rGiro("diferencia"))
					Else
						nMontoPesos= ccur(0) 
						nDiferencia = ccur(0)
					End If
					nComision = 0
					If nTipo = afxGirosEnviados Then
						'nComision = cCur(0 & rGiro("comision_captador"))
						If Isnull(rGiro("comision_captador")) Then
							nComision = 0
						Else
							nComision = cCur(rGiro("comision_captador"))							
							
							' JFMG 02-10-2008
							if rGiro("codigo_moneda") = "USD" then
								nComisionPesos = cCur(rGiro("comision_captador")) * cCur(rGiro("tipo_cambio"))
							else
								nComisionPesos = cCur(rGiro("comision_captador"))								
								if rGiro("agente_pagador") = "ME" then
									nComision = 0
								end if
							end if
							' ************************** FIN **********
						End If
					End If	
					If nTipo = afxGirosRecibidos Then
						If Isnull(rGiro("comision_pagador")) Then
							nComision = 0
						Else
							nComision = cCur(rGiro("comision_pagador"))
						End If
						'nComision = cCur(0 & rGiro("comision_pagador"))
					End If	
					%>
						<td ALIGN="right"><%=FormatNumber(nMonto, nDec, , -1)%></td>
						<td ALIGN="right"><%=FormatNumber(nMontoPesos, 0, , -1)%></td>
						<td ALIGN="right"><%=FormatNumber(nDiferencia, 0, , -1)%></td>
					<% If nTipo = afxGirosEnviados Or nTipo = afxGirosRecibidos Then %>
							<td ALIGN="right"><%=FormatNumber(nComision, nDec)%></td>
					<% End If %>
					
					<!-- JFMG 02-10-2008 -->
					<% If nTipo = afxGirosEnviados and sListaAgente = "ME" Then %>
							<td ALIGN="right"><%=FormatNumber(nComisionPesos, 0)%></td>
					<% End If %>
					<!-- ************************** FIN ********** -->
					
					<td><%=rGiro("estado")%></td>
					
				</tr>
				</a>			
			<%			
				nTotal = nTotal + cCur(0 & nMonto)
				nTotalComision = nTotalComision + FormatNumber(nComision,ndec)
				
				' JFMG 02-10-2008
				nTotalComisionPesos = nTotalComisionPesos + cCur(nComisionPesos)
				' ************************** FIN **********
				
				nTotalPesos = nTotalPesos + cCur(0 & nMontoPesos) 
				nTotalDif = nTotalDif + cCur(0 & nDiferencia)
				nCantidad = nCantidad + 1
			End If		
			codigoGiroAnt = rGiro("codigo_giro") 'MS 29-11-2013
			
			rGiro.MoveNext			
				If rGiro("codigo_moneda") <> sMn Then
				%>
					<tr style="height: 20px" CLASS="Encabezado">
						<% If nTipo = afxGirosAviso Or nTipo = afxGirosReiteraAviso  Or nTipo = afxGirosPendientes Then %>
							<td colspan="5" style="background-color: white"></td>
						<% Else %>
							<td colspan="4" style="background-color: white"></td>
						<% End If %>
						<td ALIGN="left"><b><%=FormatNumber(nCantidad, 0)%>&nbsp;Giros</b></td>			
						<td align="right"><b>Total</b></td>
						
						<!-- JFMG 12-11-2008 si son giros enviados coloca los datos de boleta de servicios-->
						<%if nTipo = afxGirosEnviados then
							if sMoneda = "USD" then%>
								<td align="right"><b><%=formatnumber(nTotalBoletaPesos,0)%></b></td>
								<td>&nbsp;</td>
								<td align="right"><b><%=formatnumber(nTotalBoletaDolares,2)%></b></td>
								<td>&nbsp;</td>
							<%elseif sMoneda = "CLP" then%>
								<td align="right"><b><%=formatnumber(nTotalBoletaPesos,0)%></b></td>
								<td>&nbsp;</td>
							<%end if%>
						<%else%>
							<td>&nbsp;</td>
						<%end if%>
						<!-- ***************************** FIN **********************************************-->
						
						<td ALIGN="right"><b><%=FormatNumber(nTotal, nDec)%></b></td>
						<td ALIGN="right"><b><%=FormatNumber(nTotalPesos, 0)%></b></td>
						<td ALIGN="right"><b><%=FormatNumber(nTotalDif, 0)%></b></td>
						<% If nTipo = afxGirosEnviados Or nTipo = afxGirosRecibidos Then %>
								<td ALIGN="right"><b><%=FormatNumber(nTotalComision, nDec)%></b></td>
						<% End If %>
						
						<!-- JFMG 02-10-2008 -->
						<% If nTipo = afxGirosEnviados and sListaAgente = "ME" Then %>
								<td ALIGN="right"><b><%=FormatNumber(nTotalComisionPesos, 0)%></b></td>
						<% End If %>
						<!-- ************************** FIN ********** -->
						
						<td></td>
					</tr>
					<tr><td><br><br><br></td></tr>
				<%
					nTotal = 0
					nTotalComision = 0
					nTotalPesos = 0
					nTotalComisionPesos = 0
					nTotalDif = 0
					nTotalBoletaPesos = 0
					nCantidad = 0
				End iF			
		Loop

'******* Temporral para Viña, Valpo, etc
		If nAgentes > 1 Then
			Select Case i
			Case 1
				Set rGiro = rGiro2
			Case 2
				Set rGiro = rGiro3
			Case 3
				Set rGiro = rGiro4
			End Select
		End If
	Next
'***************

		Set rGiro =  Nothing
		Set rGiro2 = Nothing
		Set rGiro3 = Nothing
		Set rGiro4 = Nothing
		Set afxGiro = Nothing
		
		%>
</table>
</td></tr>
</table>
</body>
<script>

	Sub FilaLista
	End Sub
	
	Sub objConsulta_OnScriptletEvent(strEventName, varEventData)
		Dim sAgtCaptador, sAgtPagador
		
		<% If nTipo = afxGirosRecibidos Then %>
				sAgtPagador = "<%=sPagador%>"
				sAgtCaptador = Trim(objConsulta.Captador) 
		<% ElseIf nTipo = afxGirosEnviados Then %>
				sAgtCaptador = "<%=sCaptador%>"
				sAgtPagador = Trim(objConsulta.Pagador)
		<% Else %>
				sAgtCaptador = "<%=sCaptador%>"
				sAgtPagador = "<%=sPagador%>"
		<% End If %>			

	   Select Case strEventName
			
			Case "Aceptar"
						window.navigate "ListaGiros.asp?Titulo=<%=sTitulo%>&Desde=" & LimpiarCampoTxt(objConsulta.Desde) & _
						"&Hasta=" & LimpiarCampoTxt(objConsulta.Hasta) & "&Cliente=<%=sCliente%>&Tipo=<%=nTipo%>" & _
						"&Captador=" & LimpiarCampoTxt(sAgtCaptador) & _
						"&NombreCliente=" & LimpiarCampoTxt(Trim(objConsulta.NombreCliente)) & _
						"&ApellidoCliente=" & LimpiarCampoTxt(Trim(objconsulta.ApellidoCliente)) & _
						"&Pagador=" & sAgtPagador & "&mn=" & Trim(objconsulta.Moneda) & _
						"&tg=" & objConsulta.TipoGiro & _
						"&st=<%=nEstadoGiro%>"
		End Select
		
	End Sub

</script>
</html>
