<%@ LANGUAGE = VBScript %>
<%
	'Se asegura que la página no se almacene en la memoria cache
	Response.Expires = 0
	If Session("CodigoCliente") = "" Then
		response.Redirect "http:../Compartido/TimeOut.htm"
		response.end
	End If
         
%>

<%	
	Dim afxGiro, Giro
	Dim sAFEXchange, sAFEXpress
	Dim sCodigo, nTipoCliente, bExtranjero
	Dim sPagador, nAgente, sNombreBeneficiario, sApellidoBeneficiario , sNombreFinal
	Dim sMonedaGiro, sMonedaPago 
	Dim nFormaPago 
    Dim nTipoCambioVentaUSD 
    
	'INTERNO-8479 MS 11-11-2016
    Dim nLugarPago
	nLugarPago = 1
	'FIN INTERNO-8479 MS 11-11-2016
	sNombreBeneficiario = Trim(Request.Form("txtNombreB"))
	sNombreBeneficiario = Replace(sNombreBeneficiario, "'", "")
	sApellidoBeneficiario = Trim(Request.Form ("txtapellidoB"))
	sApellidoBeneficiario = Replace(sApellidoBeneficiario,"'","")
	
	sNombreFinal= sNombrebeneficiario  & " " & sApellidoBeneficiario
	
	DIM MONTOLOCAL
	
	On Error	Resume Next
	
	If Request.Form("optMG") = "on" Then
		sPagador = Session("CodigoMGEnvio")
		
	Else
		sPagador = Request.Form("cbxPagador")
		sMonedaGiro = Request.Form("cbxMonedaGiro")
		sMonedaPago = Request.Form("cbxMonedaPago")
	End If
	sAFEXpress = Request.Form("txtExpress")
	sAFEXchange = Request.Form("txtExchange")
	If Session("PaisCliente") <> Session("PaisMatriz") Then		
		bExtranjero = True
	Else
		bExtranjero = False
		
	End If
	
	If Not bExtranjero Then
		If sAFEXpress = "" Then
			If Request.Form("optPersona") = "on" Then
				nTipoCliente = 1
			Else
				nTipoCliente = 2
			End If
			sAFEXpress = AgregarClienteXP
		                    End If
		                End If

	bGiroNacional = ((Session("PaisMatriz") = Request.Form("cbxPais")) And (Request.Form("cbxPais") = Request.Form("cbxPaisB")))
	
	
	'Dim sCodigoBeneficiario
	If bGiroNacional Then
		dim sSql , rsCliente
		
		'sSql = "Select Codigo_cliente,rut_cliente,pasaporte_cliente,paispasap_cliente " & _
		'		" from cliente c " & _
		'		" inner join giro g on g.codigo_beneficiario = c.codigo_cliente " & _
		'		" where nombres_cliente = " & EvaluarStr(Request.Form("txtnombreb")) & _
		'		" and apellidos_cliente = " & EvaluarStr(Request.Form("txtApellidob") ) & _
		'		" and g.codigo_Remitente = " & EvaluarStr(sAFEXpress)
		
		sSQL = "execute BuscarBeneficiario " & 	EvaluarStr(Request.Form("txtnombreb")) & ", " & _
				EvaluarStr(Request.Form("txtApellidob") ) & ", " & EvaluarStr(sAFEXpress)
				
		set rsCliente = EjecutarSqlCliente(Session("afxCnxAFEXpress"), sSQL) 
		
		If Err.number <> 0 Then
			'Set afxGiro = Nothing
			MostrarErrorMS "Consulta Cliente "
		End If
		
		Dim sCodigoBeneficiario , sRutBeneficiario , sPasaportebeneficiario , sPaisPasapBeneficiario
		If not rsCliente.eof Then		
		
			sCodigoBeneficiario = evaluarvar(rsCliente("Codigo_cliente"),"")
			sRutBeneficiario = evaluarvar(rsCliente("Rut_cliente"),"")
			sPasaporteBeneficiario = evaluarvar(rsCliente("pasaporte_cliente"),"")
			sPaisPasapBeneficiario = evaluarvar(rsCliente("paispasap_cliente"),"")
		Else	
			If trim(sNombreFinal) = trim(Request.Form("txtNombreCompletob")) Then
				
				scodigoBeneficiario  = evaluarvar(Request.Form("txtCodigob"),"")
				sRutBeneficiario = evaluarvar(Request.Form ("txtRutB"),"")
				sPasaporteBeneficiario = evaluarvar(Request.form ("txtPasapB"),"")
				sPaispasapBeneficiario = evaluarvar(Request.form ("txtPaisPasB"),"")
			Else
				scodigoBeneficiario  = Empty
				sRutBeneficiario = Empty
				sPasaporteBeneficiario = Empty
				sPaispasapBeneficiario = Empty
			End If
		End IF
		
		If scodigobeneficiario = Empty Then			
		   sCodigoBeneficiario = AgregarBeneficiarioXP
		end if

	End If
	
	ValidacionGiro
	
	if err.number <> 0 then
		MostrarErrorMS "Grabar Envio de Giro. 1"
	end if	

	
	AgregarGiro
	
	'Set afxGiro = Nothing	
		
	If Trim(Giro) = "" Then
		MostrarErrorMS "MG 5, Invoice Duplicado"
	Else				
		If Not bExtranjero Then
		
			response.Redirect "DetalleGiro.asp?Codigo=" & Giro & "&Cliente=" & sAFEXpress
		Else
			Response.Redirect "AtencionClientes.asp"
		End If
		If sAFEXchange <> "" Then
			Response.Redirect "AtencionClientes.asp?Accion=1&Campo=5&Argumento=" & sAFEXchange
		End If
		If sAFEXpress <> "" Then
			Response.Redirect "AtencionClientes.asp?Accion=1&Campo=6&Argumento=" & sAFEXpress
		End If
	End If					

	'Valiad q' giro no exista
	Sub ValidacionGiro
		'INTERNO-8479 MS 11-11-2016
        dim sLugarPago         
        sLugarPago = 1
        if Trim(Request.Form("cbxLugarPago")) <> "" then
            sLugarPago = Request.Form("cbxLugarPago")
        end if
        
        'miki APPL-23946 MM 2016-05-19
        'se valida que haya tarifa/comisión existente para el agente/pais/ciudad/monto del giro a enviar.
        'response.Write "yyy" & "|" & Session("CodigoAgente") & "," &  sPagador & "," &  Request.Form("cbxPaisB") & "," &  Request.Form("cbxCiudadB") & "," &  sMonedaGiro & "," &  sMonedaGiro & "," & Request("txtMonto") & "," &  nTarifa & "," &  nGastoTransfer & "," &  nComisionCaptador & "," &  nComisionPagador & "," & nComisionMatriz & "," &  nAfectoIva & "|"
        ObtenerTarifaGiros Session("CodigoAgente"), sPagador, Request.Form("cbxPaisB"), Request.Form("cbxCiudadB"), sMonedaGiro, sMonedaPago, _
											Request("txtMonto"), 0, 0, 0, 0, 0, 0, sLugarPago
		'FIN miki APPL-23946 MM 2016-05-19
		'FIN INTERNO-8479 MS 11-11-2016
		' codigo nuevo PSS 30-12-2009
		If Session("Perfil")= 1 then
			dim sMoneda 
			sMoneda = "CLP"
			If Not ValidarCredito(session("CodigoAgente"), sMoneda, request("txtPesos")) then
			End If
			sMontovalida = request("txtPesos")		
		else		
			If Not ValidarCredito(Session("CodigoAgente"), Request("cbxMonedaGiro"), Request("txtMonto")) Then
			End If
			sMontovalida = request("txtMonto")
		end if 
		
		'If Session("Categoria") <> 3 And Session("Categoria") <> 4 Then	
		If Session("Categoria") <> 4 Then
			If ValidarGiro(cCur(0 & Request.Form("txtBoleta")), sAFEXpress, cCur(0 & sMontoValida), sNombreBeneficiario, sApellidoBeneficiario) Then 
				Err.Raise 2003, "AfexWeb.GrabarGiros", "Existe un giro creado con datos similares al que ahora desea grabar o la boleta de servicios ya existe. Por razones de seguridad no se creará este giro.<br><br>Para poder realizarlo, es necesario anular el giro creado con datos similares.<br><br>Si necesita mayor información comuníquese con el departamento de giros." ' Tecnova rperez 29-08-2016 - Sprint 2 : Agrega texto adicional a mensaje
				MostrarErrorMS "Grabar Envio de Giro"
			End If
		End If

		If sPagador = Session("CodigoMGEnvio") Or Session("Categoria") = 3 Or Session("Categoria") = 4 Then	
			If ValidarInvoice(Session("afxCnxAFEXpress"), Trim(Request.Form("txtInvoiceMG")), Session("CodigoAgente")) Then	
				response.Redirect "http:../compartido/informacion.asp?detalle=El Invoice se encuentra duplicado. No se puede agregar el giro."
			End If
			'mostrarerrorms Session("afxCnxAFEXpress") & ", " & Trim(Request.Form("txtInvoiceMG")) & ", " & Session("CodigoMGEnvio")
		End If
	End Sub

'Métodos	
	Sub AgregarGiro()
		Dim nPrioridad, bValidaInvoice
		Dim sDireccion, sNombreB, sApellidoB
		Dim sBancoBR, sAgenciaBR, sCtaCteBR, sCpfBR, cRealesBR, i, sGiroBarsil
		Dim sSQL, rsGiro , nMonto
		Dim sTipoCuenta, nGasto, MontoDolarRecibir
		Dim nComisionCaptador,nComisionPagador,nComisionMatriz,nAfectoIva,nTarifaSugerida , nTarifaCobrada,nCambio
		Dim sBancoSW, sTipoCtaSW, sNroCuentaSW, sPin ,sCodigoSucursalPago, nMontoRecibir, nRateMontoRecibir, sMonedaRecibir
    	
		If sPagador = Session("CodigoMGEnvio") Then
			bValidaInvoice = False
			nPrioridad = 1
			sMonedaGiro="USD"
			nMonto=replace(ccur(Request.Form("txtMontoDolar")),",",".")
			
		Else
			bValidaInvoice = False
			nPrioridad = 0
			nMonto = FormatoNumeroSQL(cCur(Request.Form("txtMonto")))
		End If
		
		'miki SMC-9 MM 2015-11-30 
		If Request.Form("cbxdeposito") = 1 or sPagador = "SW" Then
			nFormaPago = 1
			sDatosDeposito = Request.Form("txtDatosDeposito")
			i = instr(sDatosDeposito, ";")
			if i > 0 then
				sBancoBR = left(sDatosDeposito, i - 1) 
				sDatosDeposito = mid(sDatosDeposito, i + 1)
				i = instr(sDatosDeposito, ";")
				sTipoCuenta = left(sDatosDeposito, i - 1) 
				sDatosDeposito = mid(sDatosDeposito, i + 1)
				i = instr(sDatosDeposito, ";")
				sCtaCteBR= left(sDatosDeposito, i - 1) 
				sDatosDeposito = mid(sDatosDeposito, i + 1)
				i = instr(sDatosDeposito, ";")
				sMonedaDeposito = left(sDatosDeposito, i - 1) 'INTERNO-8479 MS 11-11-2016		
		    End If
		Else
		   nFormaPago = 0
		End IF
    
		sDireccion = Trim(Request.Form("txtDireccionB"))
		sDireccion = Replace(sDireccion, "'", "")
		sNombreB = Trim(Request.Form("txtNombreB"))
		sNombreB = Replace(sNombreB, "'", "")
		sApellidoB = Trim(Request.Form("txtApellidoB"))
		sApellidoB = Replace(sApellidoB, "'", "")
		
		'mostrarerrorms Request.Form("txtGiroBrasil")		
		
		sGiroBrasil = Request.Form("txtGiroBrasil")
		i = instr(sGiroBrasil, ";")
		if i > 0 then
			sBancoBR = left(sGiroBrasil, i - 1) 
			sGiroBrasil = mid(sGiroBrasil, i + 1)
			i = instr(sGiroBrasil, ";")
			sAgenciaBR = left(sGiroBrasil, i - 1) 
			sGiroBrasil = mid(sGiroBrasil, i + 1)
			i = instr(sGiroBrasil, ";")
			sCtaCteBR = left(sGiroBrasil, i - 1) 
			sGiroBrasil = mid(sGiroBrasil, i + 1)
			i = instr(sGiroBrasil, ";")
			sCpfBR = left(sGiroBrasil, i - 1) 			
			sGiroBrasil = mid(sGiroBrasil, i + 1)
			i = instr(sGiroBrasil, ";")
			cRealesBR = sGiroBrasil
			cRealesBR = FormatoNumeroSQL(cCur(cDbl(cRealesBR)))
		
		End if	
		if cRealesBR = empty then cRealesBR = "null"
		
		If sPasaporteBeneficiario = " " Then
			sPasaporteBeneficiario = trim(Request.Form("txtIdBeneficiario"))
		
		End IF
		
		
		If Session("Perfil")= 1 And not bGiroNacional Then
			sMonedaGiro = "CLP"
			nMonto = cCur(0& Request.Form ("txtPesos"))
			nCambio  = cCur(0 & Request.form("txtTipoCambioOP"))
			nComisionCaptador = FormatNumber(cCur(0 & Request.Form("txtComisionCaptador")) * nCambio,0)
			nCOmisionPagador = formatnumber(cCur(0 & Request.Form ("txtComisionPagador")) * nCambio,0)
			nComisionmatriz = FormatNUmber(cCur(0 & Request.Form("txtComisionMatriz"))* nCambio,0)
			nAfectoIva = FormatNumber(cCur(0 & Request.Form("txtAfectoIva"))* nCambio ,0)
			nTarifaSugerida = cCur(0& Request.Form ("txtTarifaSugeridaPesos"))
			ntarifaCobrada = cCur(0 & Request.Form("txtTarifaCobradaPesos"))
			nGasto = FormatNumber(cCur(0 & Request.Form("txtGasto"))* nCambio,0)
			nMontoDOlarRecibir = FormatNumber(cCur(0 & Request.Form ("txtMonto")),0)
			sMonedaPago = Request.Form("cbxMonedaPago")
		Else			
		
			nComisionCaptador = Request.Form("txtComisionCaptador")
			nCOmisionPagador =  Request.Form ("txtComisionPagador")
			nComisionmatriz = Request.Form("txtComisionMatriz")
			nAfectoIva = Request.Form("txtAfectoIva")
			nTarifaSugerida =  Request.Form ("txtTarifaSugerida")
			ntarifaCobrada = Request.Form("txtTarifaCobrada")
			nGasto = Request.Form("txtGasto")
			sMonedaPago = null
			
			' JFMG 25-05-2012			
			If request.Form("chkOtraMonedaPago") = "on" Then 
			    sMonedaPago = request.Form("cbxMonedaPago")			
			    nMontoDOlarRecibir = FormatNumber(cCur("0" & Request.Form("txtMontoPagar")),0)
			    If request.Form("cbxDeposito") = 1 Then sMonedaDeposito = request.Form("cbxMonedaPago")
			End If
			' FING JFMG 25-05-2012
		End If
			'Response.Write ngasto
			'Response.End 
			
		' JFMG 29-04-2011 se agregan datos para mensajeria cliente
		Dim sMensajeriaCliente		
		sMensajeriaCliente = Request.Form("txtMensajeriaCliente")
		' FIN JFMG 29-04-2011
		'INTERNO-8479 MS 11-11-2016
        Dim IdSucursalPago
        'miki SMC-9 MM 2015-11-30
        If sPagador = "SW" Then
            nFormaPago = Request.Form("cbxdeposito")
            
            If nFormaPago = 2 then
                nFormaPago = 0
                nLugarPago = 0
            End if

            sMonedaPago = sMonedaDeposito
            sBancoSW = sBancoBR
            'Si es deposito para haiti o dominicana Trae codigo de la agencia pagadora
            if (Trim(Request.Form("cbxPaisB")) = "HT" or Trim(Request.Form("cbxPaisB") = "DO")) and nFormaPago = "1" then
                IdSucursalPago = BuscaIdSucursalPago(sBancoSW)
            else
                IdSucursalPago = sBancoSW
            end if

            sTipoCtaSW = sTipoCuenta 
            sNroCuentaSW = sCtaCteBR 
            'busca pin
            sPin = GenerarPin(sPagador, IdSucursalPago)

            'busca codigo sucursal de pago
            if (Trim(Request.Form("cbxPaisB")) = "HT" or Trim(Request.Form("cbxPaisB") = "DO")) and nFormaPago = "1" then
                sCodigoSucursalPago = sBancoSW
                sBancoBR = IdSucursalPago
            else
            sCodigoSucursalPago = BuscaCodigoSucursalPago(sPagador,  nFormaPago, sBancoSW, sTipoCtaSW)
            end if
            'buscar paridad
            Dim tasa
            nRateMontoRecibir = BuscaParidad(sPagador, IdSucursalPago, sMonedaPago)
            nMontoRecibir = cDbl(nMonto) * cDbl(nRateMontoRecibir)
             
        End if
        'miki SMC-9 MM 2015-11-30 FIN

		    sSQL = " execute enviargiro " & _
			    EvaluarSTR(Session("CodigoAgente")) & ", " & EvaluarSTR(sPagador) & ", " & FormatoNumeroSQL(nMonto) & ", " & _
			    FormatoNumeroSQL(cCur(cDbl(nTarifaCobrada))) & ", " & nPrioridad & ", " & nLugarPago & ", " &  nFormaPago & "," & EvaluarSTR(sMonedaGiro) & ", " & _
			    EvaluarSTR(sMonedaGiro) & ", " & EvaluarSTR(Request.Form("txtMensajeB")) & ", " & EvaluarSTR(Request.Form("txtMsjPagador")) & "," & EvaluarSTR(sRutBeneficiario) & ", " & _
			    EvaluarStr(trim(sPasaporteBeneficiario)) & ", " & _ 
			    evaluarstr(sPaisPasapBeneficiario ) & ", " & EvaluarSTR(sNombreB) & ", " & EvaluarSTR(sApellidoB) & ", " & EvaluarSTR(sDireccion) & ", " & EvaluarSTR(Request.Form("cbxCiudadB")) & ", " & _
			    " NULL, " & EvaluarSTR(Request.Form("cbxPaisB")) & ", " & cInt(0 & Request.Form("txtPaisFonoB")) & ", " & cInt(0 & Request.Form("txtAreaFonoB")) & ", " & _
			    cCur(0 & Request.Form("txtFonoB")) & ", " & EvaluarSTR(Request.Form("txtRut")) & ", " & EvaluarSTR(Request.Form("txtPasaporte")) & ", " & _
			    EvaluarSTR(Request.Form("cbxPaisPasaporte")) & ", " & EvaluarSTR(Trim(Request.Form("txtNombres")) & Trim(Request.Form("txtRazonSocial"))) & ", " & _
			    EvaluarSTR(Request.Form("txtApellidos")) & ", " & EvaluarSTR(Request.Form("txtDireccion")) & ", " & EvaluarSTR(Request.Form("cbxCiudad")) & ", " & _
			    EvaluarSTR(Request.Form("cbxComuna")) & ", " & EvaluarSTR(Request.Form("cbxPais")) & ", " & _
			    cInt(0 & Request.Form("txtPaisFono")) & ", " & cInt(0 & Request.Form("txtAreaFono")) & ", " & cCur(0 & Request.Form("txtFono")) & ", " & _
			    EvaluarSTR(Session("NombreUsuarioOperador")) & ", " & EvaluarSTR(sCodigoBeneficiario) & ", " & _
			    EvaluarSTR(sAFEXpress) & ", " & EvaluarSTR(Request.Form("txtInvoiceMG")) & ", " & cCur(0 & Request.Form("txtBoleta")) & ", " & FormatoNumeroSQL(cCur(cDbl(nTarifaSugerida))) & ", " & _
		        FormatoNumeroSQL(cCur(cDbl(nGasto))) & ", " & FormatoNumeroSQL(cCur(cDbl(nComisionCaptador))) & ", " & _
			    FormatoNumeroSQL(cCur(cDbl(nComisionPagador))) & ", " & FormatoNumeroSQL(cCur(cDbl(nComisionMatriz))) & ", " & _
			    FormatoNumeroSQL(cCur(cDbl(nAfectoIva))) & ", " & _			
			    EvaluarStr(sBancoBR) & ", " & EvaluarStr(sAgenciaBR) & ", " & EvaluarStr(sCtaCteBR) & ", " & EvaluarStr(sCpfBR) & ", " & FormatoNumeroSQL(cRealesBR)  _
			    & ", Null, " & FormatoNumeroSQL(ccur(nTipoCambioVentaUSD)) & ", " & FormatoNumeroSQL(ccur(0 & Request.Form("txtMontoPesos")) )  _	    
			    & ", " &  FormatoNumeroSQL(ccur(0 & Request.form("txtTarifaPesos"))) & ", " & EvaluarStr(trim(sTipoCuenta)) & ", " & _
			    EvaluarStr(trim(sMonedaDeposito))  & ", " & FormatoNumeroSQL(cCur(cDbl(nMontoDOlarRecibir))) & ", " & EvaluarStr(sMonedaPago) & _
			    ", NULL, NULL, NULL, " & EvaluarStr(sMensajeriaCliente) & _ 
                ", 0 , " & EvaluarSTR(sPin) & ", " & EvaluarSTR(sCodigoSucursalPago) & _
                ", " & FormatoNumeroSQL(ccur(0 & nRateMontoRecibir)) & ", " & FormatoNumeroSQL(ccur(0 & nMontoRecibir))

		'FIN INTERNO-8479 MS 11-11-2016

		set rsGiro = EjecutarSqlCliente(Session("afxCnxAFEXpress"), sSQL)
		
		If Err.number <> 0 Then
			'Set afxGiro = Nothing
			MostrarErrorMS "Grabar Envio Giro 2"
		End If
		Giro = rsGiro("codigogiro")
		'Response.write rsGiro("codigogiro")
		'Response.End 
		
		'If afxGiro.ErrNumber <> 0 Then
		'	MostrarErrorAFEX afxGiro, "Grabar Envio Giro 3"
		'End If
		If Trim(Giro) = "" Then
			Set rsGiro = Nothing
			MostrarErrorMS "Invoice Duplicado 3"
		End If		
		
		set rsGiro = nothing

        'miki SMC-9 MM 2015-11-30
        If sPagador = "SW" Then
            ActualizaPin sPagador, sBancoSW, sPin
        End if
	End Sub

    'miki SMC-9 MM 2015-11-30
    Sub ActualizaPin (Byval sPagador, Byval sBanco, Byval sPin)
        dim sSql
        dim rsPin
        
		sSql =" UPDATE SucursalPago SET UltimoPinSucursalPago = '" & sPin & "' WHERE codigo_agente = '" & sPagador & "' and  IdSucursalPago = " & sBanco & " "
		set rsPin = EjecutarSqlCLiente (Session("afxCnxAFEXpress") ,sSql )
		If Err.number <> 0 Then
			MostrarErrorMS "Actualizar Pin SmallWorld"
		End IF
    End Sub

    'miki SMC-9 MM 2015-11-30
    Function BuscaCodigoSucursalPago (Byval sPagador, ByVal sFormaPago, Byval sBanco, Byval sTipoCta)
		Dim sSql
        Dim rsSucursalPago
        Dim nTipoCta   
    
        Select Case sTipoCta
            Case "AH" nTipoCta = "1"
            Case "CC" nTipoCta = "2"
            Case Else nTipoCta = "0"
        End Select 
        
        sSql = "SELECT t.CodigoSucursalPago FROM SucursalPago s inner join TipoPagoSucursalPago t on (s.IdSucursalPago = t.IdSucursalPago) " & _
	           " WHERE s.codigo_agente = '" & sPagador & "' and t.IdSucursalPago = " & sBanco & " and  t.FormaPago = " & Cstr(sFormaPago) & " and t.TipoCuentaPago = " & Cstr(nTipoCta) & " "
		set rsSucursalPago = EjecutarSqlCLiente (Session("afxCnxAFEXpress") ,sSql )
        If Not rsSucursalPago.EOF then
            BuscaCodigoSucursalPago = rsSucursalPago("CodigoSucursalPago")
        Else
            BuscaCodigoSucursalPago = "XXXXXXXX"
        End if
		If Err.number <> 0 Then
			MostrarErrorMS "Obtener Codigo Sucursal de Pago"
		End IF
	End Function

     'INTERNO-8479 MS 11-11-2016
	 Function BuscaIdSucursalPago (Byval CodigoSucursalPago)
		Dim sSql
        Dim rsSucursalPago
        Dim nTipoCta   
        
        sSql = "SELECT top 1 s.IdSucursalPago  FROM SucursalPago s inner join TipoPagoSucursalPago t on (s.IdSucursalPago = t.IdSucursalPago) " & _
	           " WHERE t.CodigoSucursalPago = '" & CodigoSucursalPago & "'"
		set rsSucursalPago = EjecutarSqlCLiente (Session("afxCnxAFEXpress") ,sSql )
        If Not rsSucursalPago.EOF then
            BuscaIdSucursalPago = rsSucursalPago("IdSucursalPago")
        Else
            BuscaIdSucursalPago = "0"
        End if
		If Err.number <> 0 Then
			MostrarErrorMS "Obtener Codigo Sucursal de Pago"
		End IF
	End Function
	'FIN INTERNO-8479 MS 11-11-2016

    'miki SMC-9 MM 2015-11-30
    Function GenerarPin (Byval sPagador, Byval sBanco)
		dim sSql
        dim rsPin
        
		sSql =" SELECT dbo.GenerarPin('" & sPagador & "', " & sBanco & ") AS ProximoPin"
		set rsPin = EjecutarSqlCLiente (Session("afxCnxAFEXpress") ,sSql )
        If Not rsPin.EOF then
            GenerarPin = rsPin("ProximoPin")
        Else
            GenerarPin = "XXXXXXXXXX"
        End if
		If Err.number <> 0 Then
			MostrarErrorMS "Obtener Pin SmallWorld"
		End IF
	End Function

    'miki SMC-9 MM 2015-11-30 
	'INTERNO-8479 MS 11-11-2016
    Function BuscaParidad (Byval sPagador, Byval sBanco, Byval sMoneda)
		dim sSql
        dim rsParidad
        sSql = "SELECT pa.ValorParidadAgenteServicio FROM costocobertura.agenteinternacional ai inner join costocobertura.agenteservicio ags " & _
                " on (ai.IdAgenteInternacional = ags.IdAgenteInternacional) " & _
                " inner join costocobertura.ParidadAgenteServicio pa on (ags.IdAgenteServicio = pa.IdAgenteServicio)" & _
                " inner join Corporativa.dbo.moneda mo on (pa.IdMonedaValorParidadAgenteServicio = mo.IdMoneda)" & _
                " WHERE ai.CodigoAgenteInternacional = '" & sPagador & "' and pa.IdSucursalPago = " & sBanco & " and pa.IdTipoParidadAgenteServicio = 1" & _
                " and pa.IdEstadoParidadAgenteServicio = 1 " & _
                " and mo.codigo_moneda = '" & sMoneda & "' " 
        'FIN miki INTERNO-7701 MM 2016-08-25
		set rsParidad = EjecutarSqlCLiente (Session("afxCnxAFEXpress") ,sSql )
        If Not rsParidad.EOF then
            BuscaParidad = rsParidad("ValorParidadAgenteServicio")
        Else
            BuscaParidad = "1"
        End if
		If Err.number <> 0 Then
			MostrarErrorMS "Obtener Paridad SmallWorld"
		End IF
	End Function
	'FIN INTERNO-8479 MS 11-11-2016

	Function AgregarBeneficiarioXP
		Dim afxClienteXP
		
		Set afxClienteXP = Server.CreateObject("AfexClienteXP.Cliente")
		AgregarBeneficiarioXP = afxClienteXP.Agregar(Session("afxCnxAFEXpress"), "", Session("CodigoAgente"), _
								 , , , _
								 request.Form("txtNombreB"), _
								 request.Form("txtApellidoB"),, _
								 request.Form("txtDireccionB"), , _
								 request.Form("cbxCiudadB"), _
								 request.Form("cbxPaisB"), _
	 							 CInt(0 & request.Form("txtPaisFonoB")), CInt(0 & request.Form("txtAreaFonoB")), CCur(0 & request.Form("txtFonoB")))
	 							 
		If Err.number <> 0 Then
			Set afxClienteXP = Nothing
			MostrarErrorMS "Agregar Cliente Giros 1"
		End If
		If afxClienteXP.ErrNumber <> 0 Then
			MostrarErrorAFEX afxClienteXP, "Agregar Cliente Giros 2"
		End If
		Set afxClienteXP = Nothing

	End Function
 	
%>
<!--#INCLUDE virtual="/Compartido/Constantes.asp" -->
<!--#INCLUDE virtual="/Compartido/Rutinas.asp" -->
<!--#INCLUDE virtual="/Compartido/Errores.asp" -->
<!--#INCLUDE virtual="/Compartido/Giros.asp" -->
