<%@ Language=VBScript %>
<%
    If Session("NombreUsuario") = "" Then
		response.Redirect "http:../Compartido/TimeOut.htm"
		response.end
	End If

    ' JFMG 14-11-2012
    If Request("ClienteeAFN") <> "" Then ' esto porque se está llamando desde otra aplicación (por ejemplo la de VIGO)
        Session("CodigoCliente") = Request("ClienteeAFN") 
    End If
    ' FIN 14-11-2012

 %>

<!--#INCLUDE virtual="/Compartido/TimeOut.asp" -->
<!--#INCLUDE virtual="/Compartido/Constantes.asp" -->
<!--#INCLUDE virtual="/Compartido/RutinasEncriptar.asp" -->
<%

	' JFMG 04-12-2009 datos para AFEXchangeWeb	
	dim iTipoGiroAFEXchangeWeb, sMonedaGiro, cComisionSucursal, cComisionMatriz, cMontoTarifaGiro
	' ********** FIN JFMG 04-12-2009 *************

	'Variables
	Dim rsGiro, sCodigo, afxGiro, nTipoLlamada, nTipoLista, rsTipoCambio
	Dim sAFEXchange, sAFEXpress, nEstado
	Dim nDecimales, sColorMoneda
	Dim sNumeroPIN, Correlativo, sCodigoCaptador, sCodigoPagador
	Dim nCategoria, sPaisBeneficiario, sCiudadBeneficiario
	Dim cMontoGiro, dFechaGiro
	Dim nMonedaGiro, nMontoPesos, nTarifaPesos , nIvaPesos, nTipoCambio, nComisionPesos,nComisionIva
	Dim nGastoT , nTotalPesos, sDeposito, sMensaje
	dim sCautela , sNomCliente , sApellidoCliente,sApMaternoCliente,rCliente, sPaisRemitente, sDisabled
	dim sTipoCambio
	dim nGastosTransferencia, nTarifaCobradaUSD 'APPL-2977_MS_19-08-2014
    dim sSucursalPago, nLugarPago 'INTERNO-8479 MS 09-11-2016
	
	' JFMG 05-09-2008 para la impresión del nuevo comprobante
	dim sNombreNacionalidadBeneficiario, sFechaNacimientoBeneficiario, sOcupacionBeneficiario, _
		sNombreNacionalidadRemitente, sFechaNacimientoRemitente, sOcupacionRemitente, _
		sHoraPago, sHoraCaptacion, sFechaPago, sTelefonoRemitente, sTelefonoBeneficiario, cMontoGiroPesos, nMontoPago, nMontoRecibir
	' ***************************** FIN **********************************************
    
    'APPL-278 MS 27-08-2015
    'Valida que se haya ingresado el Tipo de cambio Observado
     Dim rsObservado, nObservado
     sSql = "select dbo.MostrarTipoCambioObservado() as observado "
     set rsObservado = EjecutarSqlCliente(Session("afxCnxAFEXpress"), sSQL) 
     If Err.number <> 0 Then
	    MostrarErrorMS "Error al consultar el dólar observado."
	 End If

     if not rsObservado.EOF then
        nObservado =  rsObservado("observado")  
       if CCur(nObservado) <= ccur("0.00") then
            Response.Redirect "../Compartido/Error.asp?description=El valor del dólar observado no ha sido ingresado. Verifique con el departamento de giros si éste ha sido ingresado."
        end if  
     else
       MostrarErrorMS "Error al consultar el dólar observado."
     end if
    'FIN APPL-278 MS 27-08-2015

	' JFMG 31-05-2011 se agrega validación en AFEXChangeWeb
	'Dim sGiroenAFEXChangeWeb
	' FIN JFMG 31-05-2011
	
	' JFMG 03-06-2012
	Dim bEnviarAFEXChangeWEB
	bEnviarAFEXChangeWEB = False
	' FIN JFMG 03-06-2012
	
	Dim sMensajeFirma ' JFMG 31-07-2013, para que las sucursales impriman algún mensaje que se posicionará en la firma del comprobante de envío
	                    ' los Agentes utilizaran la Session("MensajeFirma")	
	
	
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
		set rsTipoCambio = ObtenerTipoCambioDeposito(trim(sCodigo))
	Else
		Set rsGiro = ObtenerGiroXP(sCodigo)
		set rsTipoCambio = ObtenerTipoCambioVenta(sCodigo)' agrega pss 13-08-2009
	End If
	If not rsTipoCambio.eof then
		sTipoCambio = formatnumber(rsTipoCambio("Valor"),2) 'agrega tipo de cambio de venta para agentes que pagan en moneda local pss 13-08-2009
	else 
		sTipoCambio = empty 
	end if
	
	' JFMG 04-12-2009 datos para AFEXchangeWeb	
	sMonedaGiro = rsGiro("codigo_moneda")
	cComisionMatriz = rsGiro("comision_matriz")
	cMontoTarifaGiro = rsGiro("tarifa_cobrada")
	' ********** FIN JFMG 04-12-2009 *************

	dFechaGiro = EvaluarVar(rsGiro("fecha_captacion"), "")
	If dFechaGiro = Empty Then
		dFechaGiro = Date
	End If

	If nCategoria = 3 Then	
		If IsNull(trim(rsGiro("perfilpagador")))  Then
			sPerfilImpresion =   trim(rsGiro("perfilcaptador"))
		Else
			sPerfilImpresion =trim(rsGiro("perfilpagador"))
		End If
	End If
	
	'MS 22-04-2013
	Dim sCaptacionenAFEXChangeWeb
	sCaptacionenAFEXChangeWeb = rsGiro("CaptacionenAFEXChangeWeb")
	
	Dim sPagoenAFEXChangeWeb
	sPagoenAFEXChangeWeb = rsGiro("PagoenAFEXChangeWeb")
	
    'FIN MS 22-04-2013
    
	If SperfilImpresion <> 1 then
		sCodigoCaptador = EvaluarVar(rsGiro("agente_captador"), "")
		sCodigoPagador = EvaluarVar(rsGiro("agente_pagador"), "")
		sPaisBeneficiario = EvaluarVar(rsGiro("pais_beneficiario"), "")
		sCiudadBeneficiario = EvaluarVar(rsGiro("ciudad_beneficiario"), "")
		cMontoGiro = cCur(0 & rsGiro("monto_giro"))
		TipoGiro = cInt(0 & rsGiro("tipo_giro"))
		nTipoCambio= cCur(0 & rsGiro("tipo_Cambio"))
		nIvaPesos= cCur(0 & rsGiro("monto_iva"))
		nTarifaCobradaUSD = cCur(0 & rsGiro("tarifa_cobrada"))'APPL-2977_MS_19-08-2014
		If sCodigoPagador ="ME" and rsGiro("Codigo_moneda")="CLP" then
			nMontoPesos = cMontoGiro 
			nTarifaPesos=  ccur(0 & rsGiro("tarifa_sugerida"))
			nComisionPesos =nIvaPesos
			nComisionIva = cCur(0 & rsGiro("gastos_transferencia"))
		else
		    'APPL-2977_MS_19-08-2014
			nMontoPesos = cCur(0 & round(cMontoGiro * nTipoCambio)) 
			If nCategoria = 3 Then	
			    nTarifaPesos= round(cCur(nTipoCambio * nTarifaCobradaUSD))
			    nGastosTransferencia = cCur(0 & rsGiro("gastos_transferencia"))
    		    nComisionPesos =cCur(0 & formatnumber((nTarifaCobradaUSD - nGastosTransferencia) * nTipoCambio, 0))
			else
			    nTarifaPesos= round(cCur(nTipoCambio * cCur(0 & rsGiro("tarifa_sugerida"))))
			    nComisionPesos =cCur (0 & formatnumber(nIvaPesos * nTipoCambio, 0))
			end If
			'FIN APPL-2977_MS_19-08-2014
			nComisionIva = cCur(nTarifaPesos - nComisionPesos)
		end if

		nGastot = cCur(0 & round(nTarifaPesos - nComisionIva))
		nTotalPesos = cCur(round(nMontoPesos + (nTarifapesos - nGastoT)))
		
	else
		nMontopesos =  cCur(0 & rsGiro("monto_giro"))
		nIvaPesos= cCur(0 & rsGiro("monto_iva"))
		NTARIFApESOS =  cCur(0 & rsGiro("tarifa_cobrada"))
		nToTalPesosN = cCur(0 & (nMontopesos + ntarifaPesos))
		nMontoRecibir = cCur(0 & rsGiro("MontoRecibir"))		
	end If	
	
	'ANR-7-MS-20-06-2014
	dim nAfectoIvaAP, nCategoriaPagador
	nAfectoIvaAP =rsGiro("sw_AfectoIVA")
	nCategoriaPagador=rsGiro("categoria_pagador")
	'FIN ANR-7-MS-20-06-2014
	
	' JFMG 02-09-2008 para la impresión del nuevo comprobante
	sNombreNacionalidadBeneficiario = EvaluarVar(rsGiro("nombrenacionalidadbeneficiario"), "")
	sFechaNacimientoBeneficiario = EvaluarVar(rsGiro("fechanacimientobeneficiario"), "")
	sOcupacionBeneficiario = EvaluarVar(rsGiro("ocupacionbeneficiario"), "")
	sNombreNacionalidadRemitente = EvaluarVar(rsGiro("nombrenacionalidadremitente"), "")
	sFechaNacimientoRemitente = EvaluarVar(rsGiro("fechanacimientoremitente"), "")
	sOcupacionRemitente = EvaluarVar(rsGiro("ocupacionremitente"), "")
	sHoraPago = EvaluarVar(rsGiro("hora_pago"), "")
	sHoraCaptacion = EvaluarVar(rsGiro("hora_captacion"), "")
	sFechaPago = EvaluarVar(rsGiro("fecha_pago"), "")
	nMonedaGiro = EvaluarVar(rsGiro("codigo_moneda"), "")
	nTarifaCobrada = EvaluarVar(rsGiro("tarifa_cobrada"), "")
	MontoEnvioMoneyGram = EvaluarVar(rsGiro("MontoEnvioMoneyGram"), "")
	TarifaCLPAFEX = EvaluarVar(rsGiro("TarifaCLPAFEX"), "")
	MontoRecibir = EvaluarVar(rsGiro("MontoRecibir"), "")
	TarifaMoneyGram = EvaluarVar(rsGiro("TarifaMoneyGram"), "")	
	MonedaEnvioMoneyGram = EvaluarVar(rsGiro("MonedaEnvioMoneyGram"), "")
	MonedaRecibir = EvaluarVar(rsGiro("MonedaRecibir"), "")
	GastosTransferencia = EvaluarVar(rsGiro("gastos_transferencia"), 0)
	MontoAfectoIVA =  EvaluarVar(rsGiro("monto_iva"), 0)
	sTelefonoRemitente = EvaluarVar(rsGiro("fono_remitente"), "")
	if sTelefonoRemitente = "" then
		sTelefonoRemitente = EvaluarVar(rsGiro("telefonoremitente"), "")
	end if
	sTelefonoBeneficiario = EvaluarVar(rsGiro("fono_Beneficiario"), "")
	if sTelefonoBeneficiario = "" then
		sTelefonoBeneficiario = EvaluarVar(rsGiro("telefonoBeneficiario"), "")
	end if
	cMontoGiroPesos = EvaluarVar(rsGiro("monto_pesos"), 0)
	'INTERNO-8479 MS 09-11-2016
    sSucursalPago = EvaluarVar(rsGiro("NombreSucursalPago"), "")
    nLugarPago = EvaluarVar(rsGiro("lugar_pago"), 0)
	'FIN INTERNO-8479 MS 09-11-2016
	' ***************************** FIN **********************************************
	If nMonedaGiro = "CLP" Then
    'BELE-99 MS 02-10-2015
        If cMontoGiroPesos <> "0" then
            nMontoPago = cMontoGiro
        else
    	    nMontoPago =  cMontoGiroPesos
        end if
    'FIN BELE-99 MS 02-10-2015
	else
		nMontoPago = cMontoGiro
	End If
	
	If sCodigoPagador ="ME" Then
		nMonedaGiro= "CLP"
	End IF
	
	sMensaje = Trim(rsGiro("mensaje"))
	sPaisRemitente = rsGiro("Pais_remitente")	
	
	If isnull(sPaisRemitente)  then
		sPaisRemitente = request("Pais")
		If spaisremitente <> "" then
		else
			sPaisRemitente = 1
		end if 
		sDisabled = "" 
	Else
		sDisabled = "disabled"

	End If
	
	sDeposito = rsGiro("Forma_Pago")
	Dim sFormaPago 'MS 25-04-2014
	If sDeposito = 1 Then
		sMensaje = sMensaje &   " ESTE GIRO SERA DEPOSITADO EN " & trim(rsgiro("nombre")) & " Nº" & trim(rsgiro("numero_ctacte")) & ", BANCO: " & trim(rsGiro("Descripcion"))   & ", MONEDA DEPOSITO: " & trim(rsGiro("Moneda_deposito"))
		sFormaPago = "DEPOSITO" 'MS 25-04-2014
	Else
		sDeposito = 0
		Session("Deposito") = False
		sFormaPago = "EFECTIVO" 'MS 25-04-2014
	End If
	
	If sAFEXpress = ""  And EvaluarVar(rsGiro("codigo_beneficiario"), "") <> "" Then
		CargarCliente
	End If


	nEstado = cInt(rsGiro("estado_giro"))
	If rsGiro("codigo_moneda") = Session("MonedaNacional") Then
	
		nDecimales = 0
		sColorMoneda = "DodgerBlue"
	Else
		nDecimales = 2
		sColorMoneda = "#4dc087" '"MediumSeaGreen"
	End If
	
	
	dim sId , sTipoId, sIdT, sTipoIdT,sTipoIdR,sIdR, TotalGiro, valorGiro , valortarifa
		
		If rsGiro("rut_beneficiario") <> "" Then
			sTipoId = "RUT"
			sId = FormatoRut(rsGiro("rut_beneficiario"))
		Else
			sTipoId = "PASAPORTE"
			sId = Trim(rsgiro("pasaporte_beneficiario")) & ";" & Trim(rsGiro("paispasap_beneficiario"))
		End If
		If Trim(rsGiro("pasaporte_retira")) <> "" Then
			sTipoIdT = "PASAPORTE"
			sIdT = Trim(rsGiro("pasaporte_retira")) & ";" & Trim(rsGiro("paispasap_retira"))
		Else
			sTipoIdT = "RUT"
			sIdT = FormatoRut(rsGiro("rut_retira"))
		End If
		
		If rsGiro("rut_Remitente") <> "" Then
			sTipoIdR = "RUT"
			sIdR = FormatoRut(rsGiro("rut_Remitente"))
		Else
			sTipoIdR = "PASAPORTE"
			sIdR = Trim(rsgiro("pasaporte_remitente")) & ";" & Trim(rsGiro("paispasap_remitente"))
		End If
	

    Session("NumeroReferencia") = trim(rsGiro("invoice"))
	Session("Fecha") = sFechaPago
	Session("Hora") = sHoraPago
	Session("FechaC") = dFechaGiro
	Session("HOraC") = sHoraCaptacion
	Session("CodigoAfex")= trim(sCodigo)
	Session("NombreBeneficiario") = rsGiro("nombre_retira") & " " & rsGiro("apellido_retira")
	Session("NumeroIdentificacionBeneficiario") = sID
	Session("DireccionBeneficiario") = trim(rsGiro("direccion_beneficiario"))
	Session("TipoIdentificacionBeneficiario")= sTipoID
	session("CiudadBeneficiario")= trim(rsGiro("nombre_ciudad_beneficiario"))
	Session("TelefonoBeneficiario")= "(" & trim(rsGiro("codpais_beneficiario")) & "-" & trim(rsGiro("codarea_beneficiario")) & ") " & trim(rsGiro("fono_beneficiario"))
	Session("FechaNacimientoBeneficiario") =trim(sFechaNacimientoBeneficiario)
	Session("Nacionalidadbeneficiario") =trim(sNombreNacionalidadBeneficiario)
	Session("Ocupacionbeneficiario") = trim(sOcupacionBeneficiario)
	Session("NombreRemitente") = trim(rsGiro("nombre_remitente")) & " " & trim(rsGiro("apellido_remitente"))
	Session("CiudadRemitente") = trim(rsGiro("nombre_ciudad_remitente"))
	Session("PaisRemitente") = trim(rsGiro("nombre_pais_remitente"))
	Session("DireccionRemitente") = trim(rsGIRO("Direccion_remitente"))
	Session("TipoIdentificacion") = sTipoIdR
	Session("TelefonoRemitente") = rsGiro("fono_remitente")
	Session("NacionalidadRemitente") 
	Session("estadogiro")=rsGiro("estado")
	Session("Atendidopor") = session("NombreusuarioOperador")
	Session("Agencia") = rsGiro("Nombre_pagador")
	Session("FechaNacimientoRemitente") = sFechaNacimientoRemitente
	Session("NacionalidadRemitente") = sNombreNacionalidadRemitente
	Session("OcupacionRemitente") = sOcupacionRemitente
	Session("PaisBeneficiario")= rsGiro("Nombre_pais_beneficiario")
	If rsGiro("sw_Retira") = "1" Then
		Session("NombreBeneficiario") = rsGiro("nombre_retira") & " " & rsGiro("apellido_retira")
		Session("NumeroIdentificacionBeneficiario") = sIDT
		Session("TipoIdentificacionBeneficiario")= sTipoIDT
		Session("NumeroIdentificacionRemitente")= sIdRt
		Session("Mensaje") = "Beneficiario original:  " & trim(rsgiro("Nombre_beneficiario")) & " " & trim(rsGiro("Apellido_beneficiario"))& "- " _
		& sTipoId & ": " & sId & " " & sMensaje
	ELSE
	
		Session("NombreBeneficiario") = rsGiro("nombre_beneficiario") & " " & rsGiro("apellido_beneficiario")
		Session("NumeroIdentificacionBeneficiario") = sID
		Session("TipoIdentificacionBeneficiario")= sTipoID
		Session("NumeroIdentificacionRemitente")= sIdR
		session("mensaje") = smensaje
	End If
	
	Session("MonedaMontoEnvio") = rsgiro("monto_giro")
	Session("MonedaEnvioCargo") = rsGiro("tarifa_cobrada")
	valorgiro = cCur(0 & rsGiro("monto_giro"))
	valortarifa = cCur(0 & rsGiro("tarifa_cobrada"))
	totalgiro = cCur(0 & valorgiro + valortarifa)
	session("Totalrecibir") = totalgiro
	session("TotalGiro")= rsGiro("prefijo_moneda")
	If ccur(0& rsGiro("Montorecibir")) <> 0 then
		Session("MontoRecibir") = rsGiro("Montorecibir")
		Session("MonedaRecibir") = rsGiro("monedarecibir")
	else
		Session("MontoRecibir") = rsGiro("Monto_giro")
		Session("MonedaRecibir") = rsGiro("codigo_moneda")
	end If

	IF rsGiro("Codigo_moneda")= "CLP" then
		Session("Monto") = rsGiro("prefijo_moneda")
		Session("Cargo") = rsGiro("prefijo_moneda")
		
		' JFMG 26-07-2013 mensaje legal comprobante
        If ccur("0" & rsGiro("Monto_giro")) >= ccur("0" & Session("MontoComparacionCLPMensajeFirmaEnvio")) Then
            Session("MensajeFirma") = replace(Session("MensajeFirmaEnvio"), ";", " ")
            
            
            sMensajeFirma = Session("MensajeFirmaEnvio")
            'APPL-948 MS 11-08-2015
            if nCategoria = 1 Or nCategoria = 2 then
                sMensajeFirma = replace(sMensajeFirma, "<AFEX>", "AFEX")
            else
                sMensajeFirma = replace(sMensajeFirma, "<AFEX>", session("NombreCliente"))
                Session("MensajeFirma") = Replace(sMensajeFirma,"<AFEX>",session("NombreCliente"))
            end if
            'FIN APPL-948 MS 11-08-2015
        End If
        ' FIN JFMG 26-07-2013

	else
		session("Monto") = rsGiro("codigo_moneda")
		Session("Cargo") = rsGiro("prefijo_moneda")		

		' JFMG 26-07-2013 mensaje legal comprobante
        If ccur("0" & rsGiro("Monto_giro")) >= ccur("0" & Session("MontoComparacionUSDMensajeFirmaEnvio")) Then
            Session("MensajeFirma") = replace(Session("MensajeFirmaEnvio"), ";", " ")     
            sMensajeFirma = Session("MensajeFirmaEnvio")
            'APPL-948 MS 11-08-2015
            if nCategoria = 1 Or nCategoria = 2 then
                sMensajeFirma = replace(sMensajeFirma, "<AFEX>", "AFEX")
            else
                sMensajeFirma = replace(sMensajeFirma, "<AFEX>", session("NombreCliente"))
                'Session("MensajeFirma") = Replace(sMensajeFirma,"<AFEX>", session("NombreCliente"))
                Session("MensajeFirma") = Replace(sMensajeFirma,";", " ")

            end if
            'FIN APPL-948 MS 11-08-2015
        End If
        ' FIN JFMG 26-07-2013

	end if
		
	On Error Resume Next


	Sub CargarCliente()
		Dim rsCliente, nCampo

		nCampo = afxCampoCodigoExpress
		sArgumento = rsGiro("codigo_beneficiario")
		sArgumento2 = ""
		sArgumento3 = ""
		If nCampo = 0 Then Exit Sub
					
		Set rsCliente = BuscarCliente(nCampo, sArgumento, sArgumento2, sArgumento3)
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
	
	Function AsignarPinaGiro(ByVal CodigoGiro, ByVal CodigoAgente, ByVal CodigoUsuario)
		Dim sSQL
		Dim rs
		Dim sMensajePromocion
		
		On Error Resume Next
		
		' se verifica los datos del giro para ver si se le puede asignar un pin
		If Session("CodigoAgente") = sCodigoCaptador and sCodigoPagador <> "ME" and ucase(trim(sPaisBeneficiario)) <> "CL" Then
			If Session("EstadoPromocionPinTelefonico") = 1 Then
				If Instr(Session("AgentesPromocionPinTelefonico"), CodigoAgente) > 0 Then ' INTERNO-2912 JFMG 23-11-2014 Instr(Session("CategoriaAgenteCaptadorPromocionPinTelefonico"), nCategoria) > 0 Then					
					If ccur("0" & cMontoGiro) >= ccur("0" & Session("MontoMinimoGiroPromocionPinTelefonico")) then
						
							sSQL = "exec PinTelefonico.AsignarPingaGiro " & evaluarstr(CodigoGiro) & ", " & evaluarstr(CodigoAgente) & ", " & evaluarstr(CodigoUsuario)
							set rs = ejecutarsqlcliente(Session("afxCnxAFEXpress"), sSQL)
		
							If err.number <> 0 then
								MostrarErrorMS "Obtener Número PIN Telefónico"
								
							ElseIf not rs.eof then
		
								if rs("numeropin") <> "" then
									sMensajePromocion = "PIN " & trim(rs("numeropin")) ' INTERNO-2912 JFMG 02-12-2014 & " Tel.: STGO. 7600112 - REG. (044)6900112"				
								end if
								rs.close()
							End if
		
							set rs = nothing
							
						
					End If
				End If
			End If
		End If
		
		AsignarPinaGiro = sMensajePromocion
	End Function
	
	Function BuscarHistoria(ByVal Conexion, _
							ByVal Giro)
		Dim sSQL
	      
		'Manejo de errores
		On Error Resume Next
	   
		'Crea la consulta
	    sSQL = "exec ObtenerHitoriaGiro '" & Giro & "' " 'INTERNO-2850

		'Asigna al metodo el resultado de la consulta
		Set BuscarHistoria = EjecutarSQLCliente(Conexion, sSQL)
	   
		If Err.number <> 0 Then
			MostrarErrorMS "Obtener Historia"
		End If
	End Function


    ' JFMG 25-05-2012 datos para AFEXchangeWeb
	If session("EnlaceAFEXChangeWeb") and (Session("Categoria") = 1 Or Session("Categoria") = 2)  _
	and instr("7;9",rsGiro("estado_giro")) = 0 then

	    If rsGiro("agente_captador") = Session("CodigoAgente") and sCaptacionenAFEXChangeWeb = "0" Then
		    ' ENVIO
		    iTipoGiroAFEXchangeWeb = 1 
			cComisionSucursal = rsGiro("comision_captador")
			bEnviarAFEXChangeWEB = True			 
		elseIf rsGiro("agente_pagador") = Session("CodigoAgente") and instr("1;2;4;9",rsGiro("estado_giro")) = 0 and sPagoenAFEXChangeWeb = "0" then
		    ' PAGO
			iTipoGiroAFEXchangeWeb = 2  
			cComisionSucursal = rsGiro("comision_pagador")						
			bEnviarAFEXChangeWEB = True				
		end if
	
    end if
	' ********** FIN 25-05-2012 *******************

	Response.Expires = 0

%>
<!--#INCLUDE virtual="/Compartido/Rutinas.asp" -->
<!--#INCLUDE virtual="/Compartido/Giros.asp" -->
<!--#INCLUDE virtual="/Compartido/Errores.asp" -->
<html>
<head>
<meta name="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link rel="stylesheet" type="text/css" href="../Estilos/Cliente.css">
<link rel="stylesheet" type="text/css" href="../Estilos/Reportes.css">
</head>
<script language="VBScript">
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
	
	    If <%=bEnviarAFEXChangeWEB%> Then
		    RegistrarOPAFEXchangeWeb()
		End If
	    
	    ' JFMG 14-06-2012 mensaje por anulacion de captador
	    If "<%=request("SA")%>" = "1" Then
	        msgbox "Su Petición de Anulación de Giro ha sido enviada. " & vbcrlf & _
					"ATENCION: ANTES DE REALIZAR CUALQUIER DEVOLUCIÓN DE DINERO AL CLIENTE " & vbcrlf & _
				    "VERIFIQUE EN LA HISTORIA(DEL GIRO) QUE EFECTIVAMENTE SE ENCUENTRE ANULADO O RECLAMADO, " & vbcrlf & _
					"YA QUE SERÁ SU RESPONSABILIDAD SI ESTE NO SE ALCANZA A DETENER.",,"AFEX"
	    End If	

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
				      or (Session("Categoria")) = 5 Then	
		%>
					objMenu.addchild sId, "Boleta de Servicios", "Servicios", "Principal"
					objMenu.addchild sId, "Ver Boleta", "BS", "Principal"
					objMenu.addchild sId, "Oficinas de Pago", "OfPago", "Principal"
				<% if rsGiro("estado_giro") <> 2 then%>
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
                'MS 24-04-2014: No se permite pagar giros para depósito(Forma_Pago)
				If Session("Categoria") = 1 Then
					If (rsGiro("sw_editado") = 1 Or rsGiro("tipo_giro") = 0) And Trim(sPaisBeneficiario) = Trim(Session("PaisMatriz")) and rsGiro("Forma_Pago") = 0 Then 
		%>													
						objMenu.addchild sId, "Pagar", "Pagar", "Principal"
						objMenu.addchild sId, "Pagar a Tercero", "PagarTercero", "Principal"
						objMenu.addchild sId, "Actualizar Datos Cliente", "ADC", "Principal"
						nCont = nCont + 4
		<%	
					End If 
				
				'MS 24-04-2014: No se permite pagar giros para depósito(Forma_Pago)
				ElseIf Session("Categoria") = 2 Then 
					If (rsGiro("sw_editado") = 1 Or rsGiro("tipo_giro") = 0) And Trim(sPaisBeneficiario) = Trim(Session("PaisMatriz")) and rsGiro("Forma_Pago") = 0 Then
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
						objMenu.addchild sId, "Actualizar Datos Cliente", "ADC", "Principal"
						nCont = nCont + 5
		<%			
					End If 
				ElseIf Session("Categoria") = 3 Then
					If rsGiro("sw_editado") = 1 Or rsGiro("tipo_giro") = 0 Then
					    
		%>
		
							objMenu.addchild sId, "Pagar", "Pagar", "Principal"
							objMenu.addchild sId, "Pagar a Tercero", "PagarTercero", "Principal"
							objMenu.addchild sId, "Actualizar Datos Cliente", "ADC", "Principal"
							nCont = nCont + 4
		<%	
						'End If
					End If
				ElseIf Session("Categoria") = 4 Then
					If rsGiro("agente_pagador") = Session("CodigoAgente") And rsGiro("sw_editado") = 1 Then
						If rsGiro("estado_giro") <> 1 Or Session("CodigoAgente") = Session("CodigoMoneyBroker") Then
		%>					
								objMenu.addchild sId, "Pagar", "PagarInt", "Principal"
								objMenu.addchild sId, "Avisar", "Avisar", "Principal"
								nCont = nCont + 3
		<%	
						End If
					End If
					
				End If 
		%>
		<%	
			Case 5
				If rsGiro("agente_pagador") = Session("CodigoAgente") _
				And (Session("Categoria") = 1 Or  Session("Categoria") = 2 Or Session("Categoria") = 3)or Session("Categoria") = 5 Then
		%>
					' JFMG 21-01-2009 se agrga condición para que no se puedan imprimir giros mgram
					if "<%=sCodigoCaptador%>" <> "MB" then
						objMenu.addchild sId, "Comprobante de Pago", "Comprobante", "Principal"
					end if
					
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
                ' JFMG 01-08-2013 ahora también pueden solucionar los de tipo NOTA
				If uCase(Left(rsGiro("tipo_reclamo"), 3)) = "TEL" or uCase(trim(rsGiro("tipo_reclamo"))) = "NOT2" Then
					If Session("Categoria") = 1 Then 
						If rsGiro("sw_editado") = 1 Then
			%>
							objMenu.addchild sId, "Solucionar", "Solucionar", "Principal"
							objMenu.addchild sId, "Actualizar Datos Cliente", "ADC", "Principal"
							nCont = nCont + 3
			<%	
						End If 
					ElseIf Session("Categoria") = 2 Then 						
						If rsGiro("sw_editado") = 1	Then
			%>
							objMenu.addchild sId, "Solucionar", "Solucionar", "Principal"
							objMenu.addchild sId, "Actualizar Datos Cliente", "ADC", "Principal"
							nCont = nCont + 3
			<%	
						End If 
					ElseIf Session("Categoria") = 3 Then
						If rsGiro("sw_editado") = 1 Then
							If rsGiro("estado_giro") <> 1 Or Session("CodigoAgente") = Session("CodigoMoneyBroker") Then
			%>
							objMenu.addchild sId, "Solucionar", "Solucionar", "Principal"
			<%	
							End If
						End If
					ElseIf Session("Categoria") = 4 Then
						If rsGiro("agente_captador") = Session("CodigoAgente") And rsGiro("sw_editado") = 1 Then
						
			%>
							objMenu.addchild sId, "Solucionar", "Solucionar", "Principal"
			<%	
						End If
					End If 
				End If		
				
			End Select 
		
		%>
		
		' JFMG 25-05-2012 datos para AFEXchangeWeb		
		If <%=bEnviarAFEXChangeWEB%> Then
		    objMenu.addchild sId, "OP. AFEXchangeWEB", "RAFEXCHANGEWEB", "Principal"
		End If
		' ********** FIN 25-05-2012 *******************
				
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
<script language="VBScript">
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

<!-- JFMG 05-09-2008 objeto nuevo de impresión de comprobantes -->
<object id="afxPrinterPredeterminada" style="left: 0px; top: 0px; HIGHT: 0px; width: 0px;"
classid="CLSID:A03020A4-B8C6-4F42-85FE-6BA0E4066A66"
codebase="afxImpresorComprobante.CAB#version=1,0,0,0">
</object>

<object id="afxPrinterTM295" style="left: 0px; top: 0px; HIGHT: 0px; width: 0px;"
classid="CLSID:E92A1CE3-48B8-4555-80D1-A8220F9D0AA3"
codebase="afxImpresorComprobante.CAB#version=1,0,0,0">
</object>
	
	
<!-- *********************************** FIN ******************-->

	
<table cellspacing="0" cellpadding="0" border="0" style="left: 6px; position: absolute; top: 60px">
<tr><td>
	<!-- Paso 1 -->
	<table class="borde" id="tabPaso1" cellspacing="0" cellpadding="2" border="0" height="100" width="560" style="left: 0px; position: relative; top: 0px">
		<tr height="15">
			<td colspan="5" class="Titulo">&nbsp;&nbsp;Datos del Remitente</td>		
		</tr>
		<tr height="15">
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
			
			<!-- JFMG 05-09-2008 Datos incorporados para soportar la nueva interfaz de MoneyGram -->
			<input type="hidden" id="txtNumeroIdentificacionBeneficiario" value="<%=rsGiro("numeroidentificacionbeneficiario")%>">
			<input type="hidden" id="txtTipoIdentificacionBeneficiario" value="<%=rsGiro("tipoidentificacionbeneficiario")%>">
			<input type="hidden" id="txtNumeroIdentificacionRemitente" value="<%=rsGiro("numeroidentificacionremitente")%>">
			<input type="hidden" id="txtTipoIdentificacionRemitente" value="<%=rsGiro("tipoidentificacionremitente")%>">
			<input type="hidden" id="txtNombreComunaBeneficiario" value="<%=MayMin(rsGiro("NombreComunaBeneficiario"))%>">
			<input type="hidden" id="txtNombreComunaRemitente" value="<%=MayMin(rsGiro("NombreComunaRemitente"))%>">
			
			<td colspan="3">
				<table>
					<tr>
						<td>Nombre<br>
							<input id="txtNombreR" size="25" style="height: 22px; width: 250px" value="<%=MayMin(Trim(rsGiro("remitente")))%>" disabled>
						</td>
						<td>
							País Remitente<br>
							<select name="cbxPaisR" style="width: 120px" <%=sDisabled%>>
							<%	
								CargarUbicacion 1, "", sPaisRemitente	
							%>
							</select>
						</td>				
					</tr>
				</table>
			</td>
						<!--<td colspan="3">Nombre<br><input id="txtNombreR" SIZE="25" style="HEIGHT: 22px; WIDTH: 400px" value="<%=MayMin(Trim(rsGiro("nombre_remitente")) & " " & Trim(rsGiro("apellido_remitente")))%>" disabled></td>-->
			<!-- ********************************** FIN ************************ -->
						
		</tr>
		<tr height="10">
			<td colspan="4" CLASS="Titulo">&nbsp;&nbsp;Datos del Beneficiario</td>
		</tr>
		<tr height="15">
			<td colspan="4">
				<table><tr>
				<!-- JFMG 05-09-2008 Datos incorporados para soportar la nueva interfaz de MoneyGram -->
				<td>Nombre<br><input id="txtNombres" style="height: 22px; width: 200px" value="<%=MayMin(Trim(rsGiro("nombre_beneficiario"))) & " " & MayMin(Trim(rsGiro("segundonombrebeneficiario")))%>" disabled></td>				
				<td>Apellido<br><input id="txtApellidos" style="height: 22px; width: 180px" value="<%=MayMin(Trim(rsGiro("apellido_beneficiario"))) & " " & MayMin(Trim(rsGiro("apellidomaternobeneficiario")))%>" disabled></td>				

					<% If not Isnull(Trim(rsGiro("rut_beneficiario"))) Then %>
						<td>Identificacion<br><input id="txtId" style="height: 22px; width: 100px" value="<%=FormatoRut(Trim(rsGiro("rut_beneficiario")))%>" disabled></td>
					<% Elseif not Isnull(Trim(rsGiro("pasaporte_beneficiario"))) then %>
						<td>Identificacion<br><input id="txtId" style="height: 22px; width: 100px" value="<%=Trim(rsGiro("pasaporte_beneficiario"))%>" disabled></td>
					<% Else%>
						<td>Identificacion<br><input id="txtId" style="height: 22px; width: 100px" value="<%=Trim(rsGiro("numeroidentificacionbeneficiario"))%>" disabled></td>
					<% End If %>				
				
				<!-- ********************************** FIN ************************ -->
				</tr></table>
			</td>
		</tr>
	<% If Session("Categoria") = 4 Then %>
			<tr height="15" style="display: none">
	<%	Else %>
			<tr height="15">
	<% End If %>
			<td colspan="4">
				<table><tr>
				<td COLSPAN="2">Dirección<br>
				<input id="txtDireccion" style="height: 22px; width: 350px" value="<%=MayMin(rsGiro("direccion_beneficiario"))%>" disabled></td>
						<td>País<br>
							<input style="height: 22px; width: 80px" name="txtPais" value="<%=MayMin(rsGiro("nombre_pais_beneficiario"))%>" disabled>			
						</td>
						<td>Ciudad<br>
							<input style="height: 22px; width: 110px" name="txtCiudad" value="<%=MayMin(rsGiro("nombre_ciudad_beneficiario"))%>" disabled>			
						</td>
				</tr></table>
			</td>
		</tr>
		<tr height="15">
			<td colspan="4">
				<table><tr>
				<td>Teléfono<br>
				<input disabled id="txtPaisFono" style="width: 20px" value="<%=rsGiro("codpais_beneficiario")%>">
				<input disabled id="txtAreaFono" style="width: 20px" value="<%=rsGiro("codarea_beneficiario")%>">
				<input id="txtFono" style="width: 80px" value="<%=rsGiro("fono_beneficiario")%>" disabled>
				</td>
				<td>Mensaje al Beneficiario<br><input id="txtMensaje" style="height: 21px; width: 410px" size="40" value="<%=Trim(rsGiro("mensaje"))%>">
				</tr></table>
			</td>
		</tr>
		<tr height="10">
			<td colspan="4" class="Titulo">&nbsp;&nbsp;Datos del Agente</td>			
		</tr>
	<% If Session("Categoria") = 4 Then %>
		<tr height="15" style="display: none">
	<% Else %>
		<tr height="15">
	<% End If %>
			<td colspan="4">
				<table><tr>
				<td>Agente Captador<br>
					<input style="height: 22px; width: 200px" name="txtCaptador" value="<%=MayMin(rsGiro("nombre_captador"))%>" disabled>			
				</td>			
				<td>Agente Pagador<br>
					<input style="height: 22px; width: 200px" name="txtPagador" value="<%=MayMin(rsGiro("nombre_pagador"))%>" disabled>			
				</td>			
				<td>Confirmación<br>
					<input style="height: 22px; width: 100px" name="txtConfirmacion" value="<%=rsGiro("codigo_confirmacion")%>" disabled>
				</td>
				</tr></table>
			</td>
		</tr>
		<tr height="15">
			<td colspan="4">
				<table><tr>
				<td id="tdNN" STYLE="HEIGHT: 10px; WIDTH: 140px; display: none"></td>				
				<td id="tdMoneda">Moneda de Pago<br>				
					<input style="height: 22px; width: 140px; font-weight: bold; text-color: white; background-color: <%=sColorMoneda%>" name="txtCodigoMoneda" value="<%=MayMin(rsGiro("moneda"))%>" disabled>
				</td>
				<td>Monto<br>
					<input style="height: 22px; text-align: right; width: 100px; font-weight: bold; text-color: white; background-color: <%=sColorMoneda%>" name="txtMonto" value="<%=FormatNumber(rsGiro("monto_giro"), nDecimales)%>" disabled>
				</td>
				<td style="display: none">Tarifa<br>
					<input style="height: 22px; text-align: right; width: 80px" name="txtTarifa" value="<%=FormatNumber(rsGiro("tarifa_cobrada"), nDecimales)%>" disabled>
				</td>
				<td style="display: none">Total<br>
					<input style="height: 22px; text-align: right; width: 100px" name="txtTotal" value="<%=FormatNumber(cCur(0 & rsGiro("monto_giro")) + cCur(0 & rsGiro("tarifa_cobrada")), nDecimales)%>" disabled>
				</td>
				<td>Giro<br>
					<input style="height: 22px; width: 80px" name="txtGiro" value="<%=rsGiro("codigo_giro")%>" disabled>
				</td>
				<td>Estado<br>
					<input STYLE="HEIGHT: 22px; WIDTH: 180px" NAME="txtEstado" value="<%=rsGiro("estado")%>" disabled>
				</td>
				</tr></table>
			</td>
		</tr>
		<% If sDeposito = 1 Then %>
			<tr height="15" style="display: ">
		<% Else %>
			<tr height="15" style="display: none">
		<% End If %>
		<td colspan ="4">
			<table>
			<tr>
				<td colspan="2" style="display: <%=sDisplay%>">Banco<br>
					<input style="height: 22px; width: 400px" name="txtBanco" value="<%=rsGiro("Descripcion")%>" disabled></td>
					<!-- MS 25-04-2014 -->
				<td colspan="2" style="display: <%=sDisplay%>">Forma de Pago<br>
					<input style="height: 22px; width: 100px" name="txtFormaPago" value="<%=sFormaPago%>" disabled></td>
					<!-- FIN MS 25-04-2014 -->
			</tr>
			<tr>
				<td width="246" style="display: <%=sDisplay%>">Tipo Cuenta<br>
					<input style="height: 22px; width: 200px" name="txtTipoCta" value="<%=rsGiro("Nombre")%>" disabled></td>
				<td width="152" style="display: <%=sDisplay%>">N&uacute;mero Cuenta<br>
			      <input style="height: 22px; width: 200px" name="txtNumeroCta" value="<%=rsGiro("numero_ctacte")%>" disabled></td>
			    <td width="152" style="display: <%=sDisplay%>">Moneda Depósito<br>
			      <input style="height: 22px; width: 200px" name="txtMonedaDeposito" value="<%=rsGiro("Moneda_Deposito")%>" disabled></td>
			  </tr>
			</table>		
		</td>
		</tr>
	
		<tr height="15">
			<td colspan="4">
				<table><tr>
				<td>Invoice<br>
					<input style="height: 22px; width: 100px" name="txtInvoice" value="<%=rsGiro("invoice")%>" 	disabled>
				</td>
				<td>Orden<br>
					<input style="height: 22px; width: 70px" name="txtOrden" value="<%=rsGiro("correlativo_salida")%>" disabled>
				</td>
				<td>Nº Boleta<br>
					<input style="height: 22px; text-align: right; width: 70px" name="txtBoleta" value="<%=rsGiro("numero_documento")%>" disabled>
				</td>
				<td>Mensaje al Agente Pagador<br><input id="txtNota" style="height: 21px; width: 295px" value="<%=rsGiro("nota")%>">
				</tr></table>
			</td>
		</tr>
		<tr height="15">
			<td></td>
		</tr>
		<tr height="18">
			<td colspan="4" class="Titulo">&nbsp;&nbsp;Historia</td>
		</tr>
		<tr>
			<td colspan="4">
			<table cellspacing="1" cellpadding="1" id="tbReporte" border="0" align="center" style="color: #505050; font-family: Verdana; font-size: 10px; position: relative; top: 0px">
				<tr class="Encabezado" style="height: 20px">
					<td width="90px">
						<b>Fecha</b>
					</td>
					<td width="80px">
						<b>Hora</b>
					</td>
					<td width="360px">
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
						<td><%=rsHistoria("hora") %></td><!--INTERNO-2850-->
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
	
	nTop = 309
	If Session("Categoria") = 4 Then nTop = 215
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
<table border="0" cellspacing="0" cellpadding="0" style="left: 421px; position: absolute; top: 40px">	
<tr><td>
    <object align="left" id="objMenu" style="height: 111px; left: 0px; position: relative; top: 0px; width: 160px" type="text/x-scriptlet" width="170" viewastext border="0" valign="top"><param name="Scrollbar" value="0"><param name="URL" value="http:../Scriptlets/Menu.htm"></object>
</td></tr>
</table>
</body>
<script language="VBScript">
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
					
				ElseIf Right(varEventData, 12) = "PagarTercero" Then
					If Not ValidarMBR() Then Exit Sub

					If <%=Session("IdCliente")%> = 0 Then
						window.navigate "ActualizacionCliente.asp?AFEXchange=<%=sAFEXchange%>&AFEXpress=<%=sAFEXpress%>&Tipo=DetalleGiro&Giro=<%=sCodigo%>&Accion=<%=afxAccionPagarTercero%>"
						window.close 
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
				If MsgBox("¿Está seguro que desea anular este giro?", vbYesNo+vbQuestion) <> vbYes Then Exit Sub
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
					<% If Session("ModoPrueba") Then %>
							if not ImprimirBoletaServicios then exit sub
					<% Else %>
							if not ImprimirBoletaServicios then exit sub
					<% End If %> 
					window.navigate "AtencionClientes.asp"
					window.close 

				ElseIf Right(varEventData, 3) = "ATC" Then
					window.navigate "AtencionClientes.asp?Accion=<%=afxAccionBuscar%>&Campo=<%=afxCampoCodigoExpress%>&Argumento=<%=sAFEXpress%>"
					window.close 

				ElseIf Right(varEventData, 3) = "ADC" Then
				    window.navigate "ActualizacionCliente.asp?AFEXchange=<%=sAFEXchange%>&AFEXpress=<%=sAFEXpress%>&Tipo=DetalleGiro&Giro=<%=sCodigo%>"
					window.close 
									
				ElseIf Right(varEventData, 6) = "OfPago" Then
					window.showModalDialog "OficinasPago.asp?CodPag=<%=sCodigoPagador%>&PaisB=<%=sPaisBeneficiario%>&CiudadB=<%=sCiudadBeneficiario%>", , "dialogwidth:40;dialogheight:20"
					window.close 
				ElseIf Right(varEventData, 2) = "BS" Then
					MostrarBS
				ElseIf Right(varEventData, 2) = "CP" Then
					MostrarCP
					
				' JFMG 03-12-2009 INFORMACION PARA ENCRIPTAR
				ElseIf Right(varEventData, 14) = "RAFEXCHANGEWEB" Then
				
					RegistrarOPAFEXchangeWeb()
					' ************* FIN JFMG 03-12-2009					
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
		Dim SPAISR	
		SPAISR= CBXPAISR.VALUE		
		If trim(txtId.value)= "" then 
			msgbox "Se produjeron problemas para pagar el Giro, intente nuevamente",,"Pago de giro"
			exit sub 
		End IF
		
		If sPaisr = ""  then
			msgbox "Debe agregar el país del remitente antes de pagar el giro" ,,"Pago de giro"
			exit sub 
		else

				<% If nEstado = 7 Then %>		
						frmGiro.action = "GrabarPagoGiroReclamo.asp"
				<% Else %>
						frmGiro.action = "GrabarPagoGiro.asp?eg=<%=nEstado%>"&"&Pais="&cbxpaisr.value
				<% End If %>
		End IF
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

	Function ImprimirComprobantePago()
		Dim sLinea, afxPrinter, sTipoId, sId, sIdT, sTipoIdT
		dim NomTercero, ApeTercero, beneficiario
		dim Printer

		' JFMG 05-09-2008 implementado para nuevo comprobante
		ImprimirComprobantePago = False
		If txtRut.value <> "" Then
			sTipoId = "1"
			sId = FormatoRut(txtRut.value)
		Elseif txtPass.value <> "" then
			sTipoId = "2"
			sId = Trim(txtPass.value) & ";" & Trim(txtPaisPass.value)
		Elseif txtNumeroIdentificacionBeneficiario.value <> "" then			
			sTipoId = trim(txtNumeroIdentificacionBeneficiario.value)
			sId = Trim(txtNumeroIdentificacionBeneficiario.value)
		End If
		
		NomTercero= trim(txtNombres.value)
		ApeTercero= trim(txtApellidos.value)
		sIdT = sId
		sTipoIdt = sTipoId
		If Trim(frmGiro.txtPassRetira.value) <> "" Then
			sTipoIdT = "2"
			sIdT = Trim(frmGiro.txtPassRetira.value) & ";" & Trim(frmGiro.txtPaisPassRetira.value)
			NomTercero= trim(frmGiro.txtNombresRetira.value) 
			ApeTercero= trim(frmGiro.txtApellidosRetira.value)
			beneficiario = trim(txtnombres.value ) & " " & trim(txtapellidos.value)
		Elseif frmGiro.txtRutRetira.value <> "" then
			sTipoIdT = "1"
			sIdT = FormatoRut(frmGiro.txtRutRetira.value)	
			NomTercero= trim(frmGiro.txtNombresRetira.value) 
			ApeTercero= trim(frmGiro.txtApellidosRetira.value)	
			beneficiario = trim(txtnombres.value ) & " " & trim(txtapellidos.value)
		End If
		
		' JFMG 21-07-2008 nuevo formato de impresión
		
		Dim sPrefijo, logoMoneyGram, sHoraComprobante, sFechaComprobante		
		logoMoneyGram = ""
		
		' JFMG 03-10-2008
		sHoraComprobante = "<%=sHoraPago%>"
		sFechaComprobante = "<%=sFechaPago%>"
		
		sHoraComprobante = right("0" & sHoraComprobante, 6)
		sHoraComprobante = left(sHoraComprobante, 2) & ":" & mid(sHoraComprobante, 3, 2) & ":" & right(sHoraComprobante, 2)		
		' **************** FIN *****************
			
		If "<%=sCodigoCaptador%>" = "MB" Then
			logoMoneyGram = "MoneyGRam"
			
			' JFMG 03-10-2008
			sFechaComprobante = left(sFechaComprobante, 10)
			' **************** FIN *****************
		end if		
		
		' JFMG 21-07-2008 nuevo formato de impresión
		Dim titulo, tituloSector1, tituloSector2

        titulo = "Comprobante de Pago"
        tituloSector1 = "Beneficiario / Persona que recibe:"
        tituloSector2 = "Remitente / Persona que envia:"
       		
		On Error Resume Next
		
		 If "<%=nCategoria%>" = "3" then
		 
			window.open "Comprobante_pago.asp", "noimporta", "width=800, height=600, scrollbars=NO"
		else

		
			set Printer = CreateObject("afexImpresorComprobante.Predeterminada")
			Printer.ImprimirComprobanteGiro trim(txtInvoice.value), trim(frmGiro.txtCodigoGiro.value),  NomTercero & " " & ApeTercero , _
				 trim(sIdt), trim(txtDireccion.value) & " " & trim(txtNombreComunaBeneficiario.value), sTipoIDT, trim(txtCiudad.value), "(" & trim(txtPaisFono.value) & trim(txtAreaFono.value) & ") " & trim("<%=sTelefonoBeneficiario%>"), _
				 "<%=sFechaNacimientoBeneficiario%>", "<%=sNombreNacionalidadBeneficiario%>",  left(trim("<%=sOcupacionBeneficiario%>"),35) , trim(txtNombresR.value) & " " & trim(txtApellidosR.value), txtCiudadR.value, trim(txtPaisR.value), trim(txtMensaje.value), "Beneficiario Original : "& Beneficiario & " " & stipotd & ": " & sid  , txtPagador.value, _
				 trim(txtMonto.value) & " " & "<%=nMonedaGiro%>", "<%=Session("NombreOperador")%>",  titulo, logoMoneyGram, tituloSector1, tituloSector2,  _
				 "","","", sFechaComprobante, sHoraComprobante
			if err.number <> 0 then			
				MsgBox "Se produjo un error al intentar imprimir el Comprobante de Pago. " & err.Description, , "AFEX"
			End If
		
			set Printer = Nothing
		End If 		
		ImprimirComprobantePago = True
		'*********************** Fin **********************************		
	End Function

	Function ImprimirBoletaServicios()
		Dim sLinea, afxPrinter2, sTipoId, sId, sIdR, sTipoIdR
		Dim sPromocion, sPIN, nCorrelativo, rs, Categoria		
		Dim nDec 'APPL-8606_MS_15-10-2014
		' JFMG 05-09-2008 datos para nuevo comprobante
		dim MontoGiro, MontoCargo, Total, MontoRecibir
		dim Printer1, TM295
		' ************************** FIN ************
		
		dim sMensaje2  ' pss 14-08-2009
		
		If frmGiro.txtMonedaPago.value = "<%=Session("MonedaNacional")%>" Then nDec = 0 Else nDec = 2 'APPL-8606_MS_15-10-2014
		
		If "<%=sTipoCambio%>" <> "" Then
			sMensaje2 =  "Tipo Cambio Referencial " &  "<%=sTipoCambio%>"
		else
			sMensaje = ""
		End IF
		
		ImprimirBoletaServicios = False
		
		sPIN = ""
		nCorrelativo = 0
		Categoria = "<%=nCategoria%>"
		
		'JFMG 16-12-2010
		sPromocion = "<%=AsignarPinaGiro(sCodigo, Session("CodigoAgente"), Session("NombreUsuario"))%>"
		if sPromocion = "" then sPromocion = left(trim("<%=sOcupacionRemitente%>"),35) ' INTERNO-2912 JFMG 23-11-2014 trim(txtMensaje.value)		
		'FIN JFMG 16-12-2010
		
		' JFMG implementación nuevo comprobante
		If txtRut.value <> "" Then
			sTipoId = "1"
			sId = FormatoRut(txtRut.value)
		Elseif txtPass.value <> "" then
			sTipoId = "2"
			sId = txtPass.value & txtPaisPass.value			
		
		Elseif txtNumeroIdentificacionBeneficiario.value <> "" then
			sTipoId = txtTipoIdentificacionBeneficiario.value
			sId = txtNumeroIdentificacionBeneficiario.value
		End If
		If txtRutR.value <> "" Then
			sTipoIdR = "1"
			sIdR = FormatoRut(txtRutR.value)
		Elseif txtPassR.value <> "" then
			sTipoIdR = "2"
			sIdR = txtPassR.value & txtPaisPassR.value
		Elseif txtNumeroIdentificacionRemitente.value <> "" then			
			sTipoIdR = txtTipoIdentificacionRemitente.value
			if sTipoIdR = "1" then
				sIdR = formatorut(txtNumeroIdentificacionRemitente.value)
			else
				sIdR = txtNumeroIdentificacionRemitente.value
			end if
		End If
		
		' JFMG 21-07-2008 nuevo formato de impresión
		Dim sPrefijo, logoMoneyGram, sHoraComprobante, sFechaComprobante
				
		logoMoneyGram = ""
		
		' JFMG 03-10-2008
		sHoraComprobante = "<%=sHoraCaptacion%>"
		sFechaComprobante = "<%=dFechaGiro%>"		
		
		sHoraComprobante = right("0" & sHoraComprobante, 6)
		sHoraComprobante = left(sHoraComprobante, 2) & ":" & mid(sHoraComprobante, 3, 2) & ":" & right(sHoraComprobante, 2)		
		' **************** FIN *****************
			
		If "<%=sCodigopagador%>" = "ME" Then			
			logoMoneyGram = "MoneyGram"
			
			' JFMG 03-10-2008
			sFechaComprobante = left(sFechaComprobante, 10)
			' **************** FIN *****************
			
			MontoCargo = ccur("0" & "<%=TarifaCLPAFEX%>")
			If "<%=nMonedaGiro%>" = "USD" Then				
				if trim("<%=MonedaEnvioMoneygram%>") = "USD" then
					MontoGiro = ccur("0" & "<%=nMontoGiro%>")
					sPrefijo = "USD"
				else
					MontoGiro = ccur("0" & "<%=nMontoPesos%>")
					sPrefijo = "$"
				end if
				'APPL-8606_MS_15-10-2014
				Total = sPrefijo & " " & formatnumber(ccur("0" & MontoGiro) + ccur("0" & MontoCargo), nDec)
				MontoGiro = sPrefijo & " " & formatnumber(MontoGiro,nDec)
				MontoCargo = sPrefijo & " " & formatnumber(MontoCargo,nDec)
				'APPL-8606_MS_15-10-2014
			else				
				if trim("<%=MonedaEnvioMoneygram%>") = "USD" then
					MontoGiro = ccur("0" & "<%=nMontoGiro%>")
					sPrefijo = "USD"
				else
					MontoGiro = ccur("0" & "<%=cMontoGiroPesos%>")
					sPrefijo = "$"					
				end if
				'APPL-8606_MS_15-10-2014
				Total = sPrefijo & " " & formatnumber(ccur("0" & MontoGiro) + ccur("0" & MontoCargo), nDec)
				MontoGiro = sPrefijo & " " & formatnumber(MontoGiro,nDec)
				MontoCargo = sPrefijo & " " & formatnumber(MontoCargo,nDec)
				'APPL-8606_MS_15-10-2014
			end if
			'APPL-8606_MS_15-10-2014
			if "<%=MonedaRecibir%>" = "CLP" then
				MontoRecibir = formatnumber(ccur("0" & "<%=MontoRecibir%>"),nDec) & " " & "<%=MonedaRecibir%>"
			else
				MontoRecibir = formatnumber(ccur("0" & "<%=MontoRecibir%>"),nDec) & " " & "<%=MonedaRecibir%>"
			end if			
			
			if "<%=sCodigoCaptador%>" = "AW" then 
			    GastoExentoIVA = formatnumber(ccur(0 & "<%=GastosTransferencia%>"),nDec)
			    GastoAfectoIVA = formatnumber(ccur(0 & "<%=MontoAfectoIVA%>"),nDec)
			    if trim("<%=MonedaEnvioMoneygram%>") ="CLP" then
			        TotalTarifa = formatnumber(ccur(0 & "<%=TarifaCLPAFEX%>"),nDec)
			        GastoAfectoIVAPesos = GastoAfectoIVA 
			    else
			        TotalTarifa = formatnumber(ccur(0 & "<%=nTarifaCobrada%>"),nDec)
			        GastoAfectoIVAPesos =  formatnumber(GastoAfectoIVA * ccur(0 & "<%=nTipoCambio%>"),0) 
			        sLineaItemTipoCambio = ";           T/C OBS $ " & "<%=nTipoCambio%>"
				    sLineaMontoTipoCambio = ";"
			    end if
			else
			    GastoExentoIVA = formatnumber((ccur(0 & "<%=TarifaMoneyGram%>") * 76) / 100,nDec)
			    GastoAfectoIVA = formatnumber(ccur(0 & "<%=TarifaCLPAFEX%>") - ccur(GastoExentoIVA),nDec)
			    TotalTarifa = formatnumber(ccur(0 & GastoExentoIVA) + ccur("0" & GastoAfectoIVA))
			    GastoAfectoIVAPesos = GastoAfectoIVA
			end if
			'APPL-8606_MS_15-10-2014
			GastoAfectoIVATotal = GastoAfectoIVAPesos
		else
			MontoGiro = trim(txtMonto.value)
			MontoCargo = "<%=nTarifaCobrada%>" 
			If "<%=nMonedaGiro%>" = "USD" Then
                
				sPrefijo = "USD"
				
				Total = sPrefijo & " " & formatnumber(ccur("0" & MontoGiro) + ccur("0" & MontoCargo),2)
				MontoRecibir = formatnumber(trim(txtMonto.value),2) & " " & "<%=nMonedaGiro%>"
				MontoGiro = sPrefijo & " " & formatnumber(MontoGiro,2)
				MontoCargo = sPrefijo & " " & formatnumber(MontoCargo,2)
				
				if "<%=GastosTransferencia%>" <> 0 then
				    'ANR-6-MS-18-06-2014
					if "<%=nCategoriaPagador%>" = "4" and "<%=nAfectoIvaAP%>" = "True" then
					    GastoExentoIVA = 0
					    GastoAfectoIVA = "<%=nTarifaCobrada%>"
					else
					    GastoExentoIVA = "<%=GastosTransferencia%>"
					    GastoAfectoIVA = "<%=MontoAfectoIVA%>"
					end if
					
					'FIN ANR-6-MS-18-06-2014
				else
					GastoExentoIVA = 0
					GastoAfectoIVA = "<%=nTarifaCobrada%>"
				end if
				
				GastoAfectoIVA = GastoAfectoIVA 'ANR-6-MS-18-06-2014 '"<%=MontoAfectoIVA%>"
				TotalTarifa = formatnumber(ccur(GastoExentoIVA) + ccur(GastoAfectoIVA),2)
				GastoExentoIVA = formatnumber(GastoExentoIVA, 2)
				GastoAfectoIVA = formatnumber(GastoAfectoIVA, 2)
				
				sLineaItemTipoCambio = ";           T/C OBS $ " & "<%=nTipoCambio%>"
				sLineaMontoTipoCambio = ";"
				GastoAfectoIVAPesos = formatnumber(GastoAfectoIVA * "<%=nTipoCambio%>", 0)
				GastoAfectoIVATotal = GastoAfectoIVAPesos
				
			Else			
				
				sPrefijo = "$"
			
				Total = sPrefijo & " " & formatnumber(ccur("0" & MontoGiro) + ccur("0" & MontoCargo),0)
				MontoRecibir = formatnumber(trim(txtMonto.value),0) & " " & "<%=nMonedaGiro%>"
				MontoGiro = sPrefijo & " " & formatnumber(MontoGiro,0)
				MontoCargo = sPrefijo & " " & formatnumber(MontoCargo,0)
				
				if "<%=GastosTransferencia%>" <> 0 then

                    'ANR-6-MS-18-06-2014
					if "<%=nCategoriaPagador%>" = "4" and "<%=nAfectoIvaAP%>" = "True" then
					    GastoExentoIVA = 0
					    GastoAfectoIVA = "<%=nTarifaCobrada%>"
					else
					    GastoExentoIVA = "<%=GastosTransferencia%>"
					    GastoAfectoIVA = "<%=MontoAfectoIVA%>"
					end if
					
					'FIN ANR-6-MS-18-06-2014
				else
					GastoExentoIVA = 0
					GastoAfectoIVA = "<%=nTarifaCobrada%>"
				end if
				
				GastoExentoIVA = formatnumber(GastoExentoIVA,0)
				GastoAfectoIVA = formatnumber(ccur("0" & GastoAfectoIVA),0)
				TotalTarifa = formatnumber(ccur("0" & GastoExentoIVA) + ccur("0" & GastoAfectoIVA),0)
				
				GastoAfectoIVAPesos = formatnumber(GastoAfectoIVA, 0)
				GastoAfectoIVATotal = GastoAfectoIVAPesos
				
			End If			
			
		End IF

		' JFMG 21-07-2008 nuevo formato de impresión
		Dim titulo, tituloSector1, tituloSector2
        Dim sCodMoneda, iMonto

        titulo = "Comprobante de Envio"
        tituloSector1 = "Remitente / Persona que envia:"
        tituloSector2 = "Beneficiario / Persona que recibe:"
                
		On Error Resume Next
		
		set Printer1 = CreateObject("afexImpresorComprobante.Predeterminada")
		set TM295 = CreateObject("afexImpresorComprobante.TM295")
	
		If "<%=nCategoria%>" = "3"  then 'and "<%=sPerfilImpresion%>" = "1" then
			
            window.open "Comprobante_Envio.asp", "noimporta", "width=800, height=600, scrollbars=NO"
			
		Else
            'miki SMC-29 MM 2016-03-08
            Dim rsGiro1, sPagador, sReferencia, sEtiquetaMontoRecibir, sSql1
            sPagador = txtPagador.value
            sReferencia = trim(txtInvoice.value)
            sCodMoneda = "<%=nMonedaGiro%>"
            sDireccionBeneficiario = trim(txtDireccion.value)

            if "<%=sCodigoPagador%>" = "SW" then
            <% 
                'miki SMC-29 MM 2016-03-08 FIX1
                sSql1 = "select isnull(pin,'0') pin , isnull(MontoRecibir,0) MontoRecibir, isnull(MonedaRecibir,'') MonedaRecibir from giro where codigo_giro = " & EvaluarStr(sCodigo)
                set rsGiro1 = EjecutarSqlCliente(Session("afxCnxAFEXpress"), sSql1) 
                if not rsGiro1.EOF  then
                    sReferencia = rsGiro1("pin")
                    MontoRecibir = formatnumber(rsGiro1("MontoRecibir"),0) & " " & rsGiro1("MonedaRecibir")    
                end if
                If Err.number <> 0 Then
	                MostrarErrorMS "Error al buscar datos del giro SMALLWORLD para comprobante."
	            End If
                sCodMoneda = "USD"
            %>
                
                sReferencia = "<%=sReferencia%>"
                MontoRecibir = "<%=MontoRecibir%>"
				'INTERNO-8479 MS 09-11-2016
                Dim sMensajeSW
                sMensajeSW = ""
                sPagador = "CM - " & "<%= sSucursalPago%>"

                if Trim("<%=sPaisBeneficiario%>") = "HT" then
                    sMensajeSW = "Impuesto USD $1.50"
                else
                    if Trim("<%=nLugarPago%>")= "0" and Trim("<%=sPaisBeneficiario%>") = "DO" then
                       sMensajeSW = "Forma Pago: ;Home Delivery"
                    end if
                end if
				'FIN INTERNO-8479 MS 09-11-2016
            end if
            iMonto = "<%=nMontoPago%>"

            If "<%=nMonedaGiro%>" = "USD" then
                iMonto = formatnumber(iMonto,2)
            Else
                iMonto = formatnumber(iMonto,0)
            End if 
            'FIN miki SMC-29 MM 2016-03-08

            'TECNOVA AJEMIC 25-05-2017 tec-401 Interno 6706. Agregar nombre banco y nro cuenta en comprobante si forma de pago es Depósito
			dim sNombreBanco
			dim sNroCuentaBanco
			sNombreBanco = ""
			sNroCuentaBanco = ""
			if "<%=sFormaPago%>" = "DEPOSITO" then
			    sNombreBanco = Trim(txtBanco.value)
			    sNroCuentaBanco = Trim(txtNumeroCta.value)
			End if

			'FIN TECNOVA AJEMIC 25-05-2017 tec-401 Interno 6706. Agregar nombre banco y nro cuenta en comprobante si forma de pago es Depósito

			If CajaPregunta("AFEX En Linea", "Coloque la boleta en la impresora y haga click en Aceptar") Then
				'afxPrinterPredeterminada.ImprimirComprobanteGiro 
                ' INTERNO-2912 JFMG 23-11-2014
				Printer1.ImprimirComprobanteGiro sReferencia, trim(frmGiro.txtCodigoGiro.value), trim(txtNombresR.value) & " " & trim(txtApellidosR.value), _
				trim(sIdR), trim(txtDireccionR.value) & " " & trim(txtNombreComunaRemitente.value),  sTipoIDR, trim(txtCiudadR.value), _
				"(" & trim(txtPaisFonoR.value) & trim(txtAreaFonoR.value) & ") " & trim("<%=sTelefonoRemitente%>"), _
				"<%=sFechaNacimientoRemitente%>", "<%=sNombreNacionalidadRemitente%>", sPromocion, _
				trim(txtNombres.value) & " " & trim(txtApellidos.value), txtCiudad.value, trim(txtPais.value), _
				trim(txtMensaje.value) & " " & sMensaje2, "", "<%=Session("NombreCliente")%>", _
				sCodMoneda & " " & iMonto, "<%=Session("NombreOperador")%>", titulo, logoMoneyGram, tituloSector1, tituloSector2, _
				"Cargo :;Total :;Monto a recibir :;" & MontoRecibir, MontoCargo & ";" & Total & ";;", _
				"", sFechaComprobante, sHoraComprobante, sPagador, "", "<%=sMensajeFirma%>", "", , _
                sMensajeSW, _ 
                "", _
                sDireccionBeneficiario, sNombreBanco, sNroCuentaBanco
                'TECNOVA AJEMIC 25-05-2017 tec-401 Interno 6706. Agregar sNombreBanco y sNroCuentaBanco si forma de pago es Depósito
                 ' Tecnova rperez 21-12-2016 TEC-176 INTERNO-7963
                'INTERNO-9798 MS 24-02-2017
				
                ' FIN INTERNO-2912
				if err.number <> 0 then			
					MsgBox "Se produjo un error al intentar imprimir el Comprobante de Envío. " & err.Description, , "AFEX"
					exit function
				end if			
			End If
		
			If CajaPregunta("AFEX En Linea", "Coloque la boleta de servicios en la impresora TM y haga click en Aceptar") Then
				TM295.ImprimirBoletaServiciosGiro "", frmGiro.txtCodigoGiro.value & ";Gasto exento Envio " & trim(sPrefijo) & " " & GastoExentoIVA & ";Gasto afecto IVA   " & trim(sPrefijo) & " " & GastoAfectoIVA & ";TOTAL TARIFA       " & trim(sPrefijo) & " " & TotalTarifa & ";" & sLineaItemTipoCambio,";;" & "$ " & GastoAfectoIVAPesos & ";;" & sLineaMontoTipoCambio, "$ " & GastoAfectoIVATotal
				if err.number <> 0 then			
					MsgBox "Se produjo un error al intentar imprimir la Boleta de servicios. " & err.Description, , "AFEX"				
					exit function
				end if
			End If
		End If 		
		set Printer1 = nothing
		set TM295 = nothing

		
		ImprimirBoletaServicios = True
		'*********************** Fin **********************************		
	End Function
	
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
						"&Mensaje=" & txtMensaje.value  
					
		 'APPL-8183_MS_24-09-2014
		if "<%=SCodigoPagador%>" = "ME" then 		   
			sDetalle = sDetalle & " &Monto=" & "<%=nMonedaGiro%>" & " " &  "<%=nMontoPesos%>" & _
						"&Gastos=" & "<%=nMonedaGiro%>" & " " & "<%=nComisionIva%>" & _
						"&Total=" & "<%=nMonedaGiro%>" & " " & "<%=nTotalPesos%>" & _						
						"&Comision=" & "<%=nMonedaGiro%>" & " " & "<%=nGastoT%>" & _ 
						"&tarifa=" & "<%=nMonedaGiro%>" & " " & "<%=nComisionIva+nGastoT%> " & _
						"&totalCliente=" & "<%=nMonedaGiro%>" & " " & "<%=nMontoPesos+nComisionIva+nGastoT%> " 
		Else
			sDetalle = sDetalle & "&Monto=" & frmGiro.txtMonedaPago.value & "  " & FormatNumber(cCur(0 & txtMonto.value), nDec) & _
			            "&Gastos=" & frmGiro.txtMonedaPago.value & "  " & FormatNumber(cCur(0 & trim(txtGastos.value)), nDec) & _
			            "&Comision=" & frmGiro.txtMonedaPago.value & "  " & FormatNumber(cCur(0 & trim(txtIva.value)), nDec) & _
			            "&Total=" & frmGiro.txtMonedaPago.value & "  " & Formatnumber(cCur(0 & trim(txtMonto.value)) +  cCur(0 & trim(txtGastos.value)), nDec) & _
			            "&MontoEquivalente=" & frmGiro.txtMontoEquivalente.value & _ 			           
			            "&tarifa=" & frmGiro.txtMonedaPago.value & "  " & Formatnumber(cCur(cCur(0 & trim(txtGastos.value))+cCur(0 & trim(txtIva.value))), nDec) & _
			            "&totalCliente=" & frmGiro.txtMonedaPago.value & "  " & Formatnumber(cCur(0 & trim(txtMonto.value))+cCur(0 & trim(txtGastos.value)+cCur(0 & trim(txtIva.value))), nDec) 
		End IF
		
		IF "<%=SCodigoPagador%>" <> "ME" then 
			If frmGiro.txtMonedaPago.value <> "<%=Session("MonedaNacional")%>" Then
				sDetalle = sDetalle & "&TipoCambio=" & FormatNumber(cCur(0 & trim(txtTipoCambio.value)), nDec)
				sDetalle = sDetalle & "&TotalNacional=" & "<%=Session("MonedaNacional")%>" & "  " & FormatNumber(cCur(0 & trim(txtIva.value))*cCur(0 & trim(txtTipoCambio.value)), 0)'APPL-8606_MS_15-10-2014
			End If		  
		else
		    If "<%=nMonedaGiro%>" <> "CLP" then
			    sDetalle = sDetalle & "&TipoCambio=" & FormatNumber(cCur(0 & trim(txtTipoCambio.value)), nDec)
			    sDetalle = sDetalle & "&TotalNacional=" & "<%=Session("MonedaNacional")%>" & "  " & FormatNumber(cCur(0 & trim(txtIva.value))*cCur(0 & trim(txtTipoCambio.value)), 0)'APPL-8606_MS_15-10-2014
			else
			    If frmGiro.txtMonedaPago.value <> "<%=Session("MonedaNacional")%>" Then
			        sDetalle = sDetalle & "&TipoCambio=" & FormatNumber(cCur(0 & trim(txtTipoCambio.value)), nDec)
			    end if
		    end if
		End If
		'FIN APPL-8183_MS_24-09-2014
		'msgbox sdetalle 
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
	
	sub RegistrarOPAFEXchangeWeb()
	
	    ' JFMG 31-05-2011 validación para que no aparezca la pantalla si ya está registrado en AFEXChangeWeb
	    'IF "<%=sGiroenAFEXChangeWeb%>" = "1" Then exit sub
	    ' FIN JFMG 31-05-2011
	    
		<%					

						function RemplazaLetra(Cadena)							
							cadena = replace(cadena, "Ñ", "%c3%91")
							RemplazaLetra = replace(cadena, "ñ", "%c3%b1")
						end function
	
						dim CuentaUsuarioEncriptado
						dim ContrasenaUsuarioEncriptado
						dim CuentaCuentaSucursalEncriptado
						dim ContrasenaSucursalEncriptado
						dim CodigoAFEXEncriptado
						dim TipoGiroEncriptado
						dim TipoNacionalidadGiroEncriptado
						dim sFeaturesURL
						
						dim MonedaGiroEncriptado
						dim MontoGiroEncriptado
						dim ComisionMatrizEncriptado
						dim ComisionSucursalEncriptado
						dim MontoTarifaGiroEncriptado
						
						CuentaUsuarioEncriptado = RemplazaLetra(EncriptarCadena(trim(Session("NombreUsuarioOperador"))))
						ContrasenaUsuarioEncriptado = RemplazaLetra(EncriptarCadena(trim(Session("ContrasenaOperador"))))
						CuentaCuentaSucursalEncriptado = RemplazaLetra(EncriptarCadena(trim(Session("NombreUsuarioAgente"))))
						ContrasenaSucursalEncriptado = RemplazaLetra(EncriptarCadena(trim(Session("ContrasenaAgente"))))
						CodigoAFEXEncriptado = RemplazaLetra(EncriptarCadena(trim(sCodigo)))
						TipoGiroEncriptado = RemplazaLetra(EncriptarCadena(iTipoGiroAFEXchangeWeb))
						TipoNacionalidadGiroEncriptado = RemplazaLetra(EncriptarCadena(TipoGiro))
						
						MonedaGiroEncriptado = RemplazaLetra(EncriptarCadena(trim(sMonedaGiro)))
						MontoGiroEncriptado = RemplazaLetra(EncriptarCadena(ccur(cMontoGiro)))
						If ucase(sMonedaGiro) = "CLP" then
						    ComisionMatrizEncriptado = RemplazaLetra(EncriptarCadena(ccur(formatnumber(ccur(cComisionMatriz),0))))
						    ComisionSucursalEncriptado = RemplazaLetra(EncriptarCadena(ccur(formatnumber(ccur(cComisionSucursal),0))))
						else 
						    ComisionMatrizEncriptado = RemplazaLetra(EncriptarCadena(ccur(cComisionMatriz)))
						    ComisionSucursalEncriptado = RemplazaLetra(EncriptarCadena(ccur(cComisionSucursal)))
						end if
						MontoTarifaGiroEncriptado = RemplazaLetra(EncriptarCadena(ccur(cMontoTarifaGiro)))
						
						If trim(sMonedaGiro) = trim(Session("MonedaNacional")) Then
                            'CUM-505 MS 02-02-2016
                            If nSucursalUsaBE = "Falso" Then
							    sFeaturesURL = "height=100,width=350,left=500,top=500, title=eAFEX"
                            else
                                sFeaturesURL = "height=250,width=505,left=500,top=500, title=eAFEX"
                            End if
                             'FIN CUM-505 MS 02-02-2016
						else
							sFeaturesURL = "height=540,width=710,left=250,top=250, title=eAFEX"
						end if
					
					
					    Dim sURLAfexChangeWeb
					    
					    'If InStr(Session("ListaAgentesAFEXChangeWebQA"), Session("CodigoAgente")) > 0 Then
					        sURLAfexChangeWeb = Session("URLAFEXChangeWEB")
					    'Else
					    '    sURLAfexChangeWeb = "laurel:91"
					    'End If
					
					%>	 	 
	
						window.open "http://" & "<%=sURLAfexChangeWeb%>" & "/iniciogiro.aspx?CuentaUsuario=<%=CuentaUsuarioEncriptado%>" & _
								"&ContrasenaUsuario=<%=ContrasenaUsuarioEncriptado%>" & _
								"&CuentaSucursal=<%=CuentaCuentaSucursalEncriptado%>" & "&CodigoAFEX=<%=CodigoAFEXEncriptado%>" & _
								"&TipoGiro=<%=TipoGiroEncriptado%>" & _
								"&ContrasenaSucursal=<%=ContrasenaSucursalEncriptado%>" & _
								"&MonedaGiro=<%=MonedaGiroEncriptado%>" & _
								"&MontoGiro=<%=MontoGiroEncriptado%>" & _
								"&ComisionMatriz=<%=ComisionMatrizEncriptado%>" & _
								"&ComisionSucursal=<%=ComisionSucursalEncriptado%>" & _
								"&IP=" & "<%=request.servervariables("REMOTE_ADDR")%>" & _
								"&MontoTarifaGiro=" & "<%=MontoTarifaGiroEncriptado%>" & _
								"&TipoNacionalidad=" & "<%=TipoNacionalidadGiroEncriptado%>", "","<%=sFeaturesURL%>"
								
	
	end sub
	
-->
</script>
<!--#INCLUDE virtual="/Compartido/Rutinas.htm" -->
</html>
