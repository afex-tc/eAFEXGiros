

<%@ Language=VBScript %>
<!--#INCLUDE virtual="/Compartido/TimeOut.asp" -->
<!--#INCLUDE virtual="/Compartido/Constantes.asp" -->
<!--#INCLUDE virtual="/Compartido/Rutinas.asp" -->
<!--#INCLUDE virtual="/Compartido/RutinasEncriptar.asp" -->
<%

	Dim nAccion, sPais, sCiudad, sPagador, sMoneda, nMonto
	Dim nTarifaSugerida, nTarifaCobrada, bCliente, bExtranjero
	Dim nTarifa, nGastoTransfer, nComisionCaptador, nComisionPagador
	Dim nComisionMatriz, nAfectoIva, nDDIpais, nDDICiudad
	Dim sNombreB, sApellidoB, sDireccionB, sFonoB, bGiroAnterior, sDisplay
	Dim bGiroNacional, nDec, sColorMoneda
	Dim sMensaje, rs, sNota, sMaximo, sRecomendacion, sRequerimiento
	Dim bCargarMensaje
	dim rsTc, rsAgente
	Dim cTC
	Dim sSQL
	Dim cReales
	Dim nMontoDolar, nMontoPesos, nTipoCambio, nTarifaPesos, nTotalPesos
	Dim nCambioMG, bDeposita, bBanco, sBanco, sTipo, sCliente, sCuenta, sMonedaDeposito, nCambioOp
	Dim sRutBeneficiario , sCodigobeneficiario , sPasaporteBeneficiario,sPaisPasapBeneficiario, sNombreCompletob
	Dim nMontoPesosOP, nTarifaSugeridaPesos, nTarifaCobradaPesos', nTotalPesos
	dim nMontoD, nTarifaSugeridaDolar, nTarifaCobradaDolar, nTotal, nPOs1
	Dim bOtrasMonedasPago, rsOtrasMonedasPago, sMonedaPago, sTipoCambioMonedaPagador, nMontoMonedaPagador ' JFMG 24-05-2012
	Dim sMsjNombresBeneficiario, sMsjApellidosBeneficiario, sMsjTelefono 'INTERNO-3855 MS 24-04-2015
    Dim sPaisPasaporteRemitente, sPasaporteRemitente, sRutR, nAutorizacionOperarPasaporte, nPasaporteVerificado', nAutorizado 'CUM-505 MS 03-02-2016
    Dim bPagaDomicilio 'INTERNO-8479 MS 09-11-2016
	nMontoDolar = Request.Form("txtMontoDolar") 
	nTipoCambio = Request.Form("txtTipoCambio")
	sTipo = Request.Form("cbxDeposito")
	
	nMontopesos = Request.Form ("txtPesos")
	
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

	' busca perfil de impresion del agente para ver el como se mostrara el envio de giro 20-01-2009
	sSql = "execute BuscarPerfilAgente " & EvaluarStr(session("CodigoAgente"))
	set rsAgente = EjecutarSqlCliente(Session("afxCnxAFEXpress"), sSQL) 
		
		If Err.number <> 0 Then
			MostrarErrorMS "Consulta Agente "
		End If
	If rsAgente("Perfil_impresion") = 0 or isnull(rsAGente("Perfil_impresion")) Then
		Session("Perfil") = 0
	Else
		Session("Perfil") = evaluarvar(rsAgente("Perfil_impresion"),"")
	End If
	
	bCargarMensaje = False
	nAccion = cInt(0 & Request("Accion"))
	nTarifa = cCur(0)
	nTarifaCobrada = cCur(0)
	nTarifaCobradaDolar = cCur(0)
	ntarfiaSugeridaDolar = cCur(0)
	nMonto = cCur(0)
	nGastoTransfer = cCur(0)
	nComisionCaptador = cCur(0)
	nComisionPagador = cCur(0)
	nComisionMatriz = cCur(0)
	nAfectoIva = cCur(0)
	nPOs1 = cCur(0)
	'CUM-505 MS 15-02-2016
    if trim(Request.Form("txtPasapRemitente"))<>"" then
	    sPasaporteRemitente = trim(Request.Form("txtPasapRemitente"))
        sPaisPasaporteRemitente = trim(Request.Form("txtPaisPasaporteRemitente"))
        sRutR = ""
    else 
        sPasaporteRemitente = trim(Request("sPasap"))
        sPaisPasaporteRemitente = trim(Request("sPaisPass"))
    end if
    'FIN CUM-505 MS 15-02-2016

	If Session("Categoria") = 4 Then		
		bExtranjero = True
		sDisplay = "none"
	Else
		bExtranjero = False
		sDisplay = ""
	End If
	If Request.Form("txtExchange") <> "" Or Request.Form("txtExpress") <> "" Then
		bCliente = True
	Else
		bCliente = False
	End If
	bGiroAnterior = False
	sCliente = Request.Form("txtNombres")

	Select Case nAccion
	Case afxAccionPais, afxAccionCiudad, afxAccionPagador
		If nAccion = afxAccionPagador Then
			Set rs = BuscarInformacionPagador(Session("afxCnxAFEXpress"), Request.Form("cbxPagador"), _
											  Request.Form("cbxPaisB"), Request.Form("cbxCiudadB"))
			
			If Err.Number <> 0 Then
				Set rs = Nothing
				
				Response.Redirect "../Compartido/Error.asp?description=" & err.Description
			End If
		
			sMensaje = Empty
		
			If Not rs.EOF Then
				bCargarMensaje = True
				sNota = rs("nota")
				If IsNull(rs("maximo")) Then
					sMaximo = ""
				Else
					sMaximo = "MONTO MAXIMO DIARIO U$" & rs("maximo")
				End If
				sRecomendacion = rs("recomendacion")
				sRequerimiento = rs("requerimiento")
			End If
			Set rs = Nothing
		End If
		Cargar
		'Response.Write  bExtranjero
		If Not bExtranjero Then	CalcularTarifa			
		
	Case afxAccionMonedaPago
			Cargar
			CalcularTarifaUltimoGiro
			sCodigoBeneficiario = request("Codigob")
			sRutBeneficiario = request("Rutb")
			sPasaporteBeneficiario = request("PasaporteB")
			sPaisPasapBeneficiario = request("PaispasB")
			sNombreCompletoB = Request.Form("txtnombreCompletoB")
			
	Case afxAccionMonto	
			Cargar
			If Not bExtranjero  Or sPagador = "ME" Then	CalcularTarifa		
			CalcularPesos	
			If Request.form("txtCodigob")= "" Then		
    	
				sCodigoBeneficiario = request("Codigob")
				sRutBeneficiario = request("Rutb")
				sPasaporteBeneficiario = request("PasaporteB")
				sPaisPasapBeneficiario = request("PaispasB")
				sNombreCompletoB = Request.Form("txtnombreCompletoB")

			Else
				sCodigoBeneficiario = Request.form("txtCodigob")
				sRutBeneficiario = Request.Form ("txtRutb")
				sPasaporteBeneficiario = Request.Form ("txtPasapB")
				sPaisPasapBeneficiario = Request.Form ("txtPaispasB")
				sNombreCompletoB = Request.Form ("txtnombreCompletoB") 

			End if
            'CUM-505 MS 12-02-2016
            if trim(Request.Form("txtAccionCumplimiento")) = "0" then
                Response.Redirect "../Compartido/Error.asp?description=Por el monto de la operación y la identificación del cliente, no es posible realizar el envío del giro, debe registrar al cliente con la información solicitada."
            elseif trim(Request.Form("txtAccionCumplimiento")) = "2" then
                    Response.Redirect "../Compartido/Error.asp?description=Si no es posible identificar al cliente con el RUT, favor comunicarse con el Departamento de Cumplimiento."
            elseif trim(Request.Form("txtAccionCumplimiento")) = "3" then
                    Response.Redirect "../Compartido/Error.asp?description=<b>Validación de antecedentes en espera, vuelva a intentar enviar el giro en unos minutos...</b>.<br/><br/>Si tiene cualquier duda comuníquese con el Departamento de Cumplimiento."
            elseif trim(Request.Form("txtAccionCumplimiento")) = "4" then
                    Response.Redirect "../Compartido/Error.asp?description=<b>Si ya actualizó los datos del cliente con rut, vuelva a intentar enviar el giro...</b>.<br/><br/>Si tiene cualquier duda comuníquese con el Departamento de Cumplimiento."
            'SMC-53 MS 17-03-2016
            elseif trim(Request.Form("txtAccionCumplimiento")) = "5" then
                    Response.Redirect "../Compartido/Error.asp?description=<b>El cliente debe estar registrado en giros para realizar esta operación...</b>.<br/><br/>Si tiene cualquier duda comuníquese con el Departamento de Cumplimiento."
            end if
            'FIN SMC-53 MS 17-03-2016
            'FIN CUM-505 MS 12-02-2016
	Case afxAccionTarifa
            'CUM-505 MS 07-03-2016
            if trim(Request.Form("txtAccionCumplimiento")) = "0" then
                Response.Redirect "../Compartido/Error.asp?description=Por el monto de la operación y la identificación del cliente, no es posible realizar el envío del giro, debe registrar al cliente con la información solicitada."
            elseif trim(Request.Form("txtAccionCumplimiento")) = "2" then
                    Response.Redirect "../Compartido/Error.asp?description=Si no es posible identificar al cliente con el RUT, favor comunicarse con el Departamento de Cumplimiento."
            elseif trim(Request.Form("txtAccionCumplimiento")) = "3" then
                    Response.Redirect "../Compartido/Error.asp?description=<b>Validación de antecedentes en espera, vuelva a intentar enviar el giro en unos minutos...</b>.<br/><br/>Si tiene cualquier duda comuníquese con el Departamento de Cumplimiento."
            elseif trim(Request.Form("txtAccionCumplimiento")) = "4" then
                    Response.Redirect "../Compartido/Error.asp?description=<b>Si ya actualizó los datos del cliente, inténtelo nuevamente...</b>.<br/><br/>Si tiene cualquier duda comuníquese con el Departamento de Cumplimiento."
            'SMC-53 MS 17-03-2016
            elseif trim(Request.Form("txtAccionCumplimiento")) = "5" then
                    Response.Redirect "../Compartido/Error.asp?description=<b>El cliente debe estar registrado en giros para realizar esta operación...</b>.<br/><br/>Si tiene cualquier duda comuníquese con el Departamento de Cumplimiento."
            end if
            'FIN SMC-53 MS 17-03-2016
            'FIN CUM-505 MS 07-03-2016
            
			Cargar
			CalcularTarifaCobrada
			
	Case AfxAccionNuevo 
			Cargar
			
			If Request.Form("OptPesos")= "on" then
				CalcularNuevoMonto
				CalcularTarifa
			
				nTarifaSugeridaPesos = cCur(0 &  nTarifaSugeridadolar)* cCur(0 & nCambioOP)
				ntarifacobradapesos = formatNumber(cCur(0 & ntarifaCObradaDOlar) * cCur(0 & nCambioOP),0)
				ntotalpesos =  ntarifacobradapesos + nmontopesos 
				If cCur(0 & Request.Form("txtTarifacobradapesos"))<> 0  then
					ntarifaCobradaPesos= formatNumber(cCur(0 & Request.Form("txttarifacobradapesos")),0)
					nTotalPesos = FormatNumber( nTarifaCobradaPesos + nMontoPesos,0)
				End IF
				CalcularComisionDOlar
				
				'Response.Write Ccur(0 &  Request("nPos"))
				If CcUR(0 & Request("nPos"))> 0 Then 
					nPos1= instr(nMonto,"," )
					If nPos1 > 0 then
						nMonto = formatNumber(nMonto,0)
						calculartarifa
						CalcularNuevoMontoPesos			
					End IF
				end if				
				
			else
				'Response.Write Request.Form ("txtTarifaCobrada")
				If Request.Form ("txtTarifaCobrada") <> 0 Then
					ntarifaCobradaDOlar = Request.Form("txtTarifaCobrada")
				End IF
				CalcularTarifa
				CalcularNuevoMonto
			end if
        
	Case Else
			If Not CargarUltimoEnvio() Then
				If bExtranjero Then
					sPais = Session("PaisMatriz")
					sCiudad = Session("CiudadMatriz")
					sPagador = Session("CodigoMatriz")
					nDDIPais = ObtenerDDI(1, sPais)
					nDDICiudad = ObtenerDDI(2, sCiudad)
				End If
				sMoneda = Session("MonedaExtranjera")
				'INTERNO-8479 MS 09-11-2016
                sMonedaPago = trim(Request.Form("cbxMonedaPago"))
                if trim(sMonedaPago) = "" then
                    sMonedaPago = sMoneda
                End if
				'FIN INTERNO-8479 MS 09-11-2016
			Else
				bGiroAnterior = True
			End If
				
	End Select
	bGiroNacional = (TRIM(sPais) = Session("PaisMatriz") And TRIM(Request.Form("cbxPais")) = Session("PaisMatriz"))	
	'INTERNO-3855 MS 26-04-2015
	if  Request.Form("cbxPagador") ="FP" then
	    sMsjNombresBeneficiario=" *Ingresar dos nombres"
	    sMsjApellidosBeneficiario=" *Ingresar dos apellidos"
	    sMsjTelefono=" *"
    else
        sMsjNombresBeneficiario=""
	    sMsjApellidosBeneficiario=""
	    sMsjTelefono=""
    End IF
    'FIN INTERNO-3855 MS 26-04-2015
	If Trim(sMoneda) = "" Then 
		If bGiroNacional Then
			sMoneda = Session("MonedaNacional")
		Else
			sMoneda = Session("MonedaExtranjera")
		End If
	End If
	If bExtranjero Then sPagador = Session("CodigoMatriz")
	If bGiroNacional Then sPagador = Session("CodigoMatriz")
	
	If sMoneda = Session("MonedaNacional")  or sPagador = session("CodigoMGEnvio")Then
			nDec = 0
			sColorMoneda = "DodgerBlue"
			
	Elseif sMOneda = Session("MonedaExtranjera") Then
			nDec = 2
			sColorMoneda = "#4dc087"
	End If
   


   If sPagador =session("CodigoMGEnvio")  Then
  		Dim rsTCambio
		Set rsTCambio = TCambioMG (session("afxCnxAFEXpress") , session("CodigoMGEnvio"))
		If Not rsTCambio.eof Then
			nCambioMG = rsTCambio("Valor")
		End If

    End If 
    
    If Session("Perfil") = 1 Then
		Dim rsTCambioOP
		Set rsTCambioOP = TCambioOperacion(session("afxCnxAFEXpress") )
		If Not rsTCambioOP.eof Then
			nCambioOp = formatNumber(rsTcambioOP("Valor"),2)
		End If
    End If
    
    bOtrasMonedasPago = False ' JFMG 24-05-2012
   	If sPagador <> Empty Then
		dim rsPago 
		Set rsPago = PagadorDeposita( spagador , session("afxCnxAFEXpress"), sPais ) 'INTERNO-8479 MS 09-11-2016
		If Not rsPago.EOF Then
			bdeposita = True
		Else
			bDeposita = False
			sTipo = 0
			bBanco = False
		End If
		
        'INTERNO-8479 MS 09-11-2016
		dim rsPagoHomeDelivery
		Set rsPagoHomeDelivery = PagaHomeDelivery( spagador , session("afxCnxAFEXpress"), sPais )
		If Not rsPagoHomeDelivery.EOF Then
			bPagaDomicilio = True
		Else
			bPagaDomicilio = False			
		End If

		' JFMG 24-05-2012 para cargar mas monedas de pago		
        sMonedaPago = request.Form("cbxMonedaPago")       
        Set rsOtrasMonedasPago = PagadorDiferentesMonedas(sPagador , session("afxCnxAFEXpress"), sPais)
        'FIN INTERNO-8479 MS 09-11-2016
        If Not rsOtrasMonedasPago.EOF Then
			bOtrasMonedasPago = True
    		
            If sMonedaPago = "" Then 
                sMonedaPago = rsOtrasMonedasPago("moneda_pago")
            End If
            
            Dim rsTCPagador
            SET rsTCPagador = TCambioPagador(session("afxCnxAFEXpress"), sPagador, sMonedaPago)
            sTipoCambioMonedaPagador = "0"
            If NOT rsTCPagador.EOF Then sTipoCambioMonedaPagador = rsTCPagador("valor")
            SET rsTCPagador = nothing
            
            nMontoMonedaPagador = cint("0" & nMonto) * cint("0" & sTipoCambioMonedaPagador)
        Else
             sMonedaPago = ""
        End If
		' FIN JFMG 24-05-2012
		
	End if
	
    'CUM-480 MS 24-08-2015
    Dim nMontoSumatoriaCumplimiento, rsCumplimiento, nGiroAutorizado 'APPL-278 MS 27-08-2015
    
    if Request.Form("cbxPaisB") = "CO" then
        sSql = "select Cumplimiento.MostrarMontoTotalGirosUSD(" & EvaluarStr(trim(Request.Form("txtExpress"))) &  ", " & replace(cstr(nObservado),",",".") & ", 30, NULL) as MontoTotalGirosUSD " 'CUM-480 MS 02-10-2015
        set rsCumplimiento = EjecutarSqlCliente(Session("afxCnxCorporativa"), sSQL) 

        if not rsCumplimiento.EOF then
            nMontoSumatoriaCumplimiento =  rsCumplimiento("MontoTotalGirosUSD")    
        else
            nMontoSumatoriaCumplimiento =  0
        end if
    

    end if
    'FIN CUM-480 MS 24-08-2015

    'CUM-505 MS 02-02-2016
    if  trim(sPasaporteRemitente) <> "" AND trim(sPaisPasaporteRemitente) = "CO" then
        if trim(Request.Form("txtExchange")) <> "" then
            'Verifica si el cliente tiene la historia de autorización de pasaporte: "AUTORIZADO CLIENTE CON PASAPORTE"
            sSql = "select isnull(CONVERT(char(1),Cumplimiento.VerificarAutorizacionPasaporte(" & trim(Request.Form("txtExchange"))  & ")),'0') as AutorizacionOperarPasaporte "
            set rsCumplimiento = EjecutarSqlCliente(Session("afxCnxCorporativa"), sSQL) 

            if not rsCumplimiento.EOF then
                nAutorizacionOperarPasaporte = Cint(rsCumplimiento("AutorizacionOperarPasaporte"))
            else
                nAutorizacionOperarPasaporte =  0
            end if

            'Verifica si está ingresado y autorizado el pasaporte del cliente
            sSql = "select isnull(CONVERT(char(1),Cumplimiento.VerificarAutorizacionDocumentoCliente(" & trim(Request.Form("txtExchange"))  & ",36)),0) as AntecedenteVerificado "
            set rsCumplimiento = EjecutarSqlCliente(Session("afxCnxCorporativa"), sSQL) 

            if not rsCumplimiento.EOF then
                nPasaporteIngresadoYAprobado = Cint(rsCumplimiento("AntecedenteVerificado"))
            else
                nPasaporteIngresadoYAprobado =  0
            end if
        else
            nAutorizacionOperarPasaporte =  2 'el cliente no existe en corporativa
            nPasaporteIngresadoYAprobado =  0
        End if
    end if

     Dim sURLIngresarClienteCorporativoAfexWeb, sURLIngresarDocumentoClienteCorporativoAfexWeb
     sURLIngresarClienteCorporativoAfexWeb = Session("URLIngresarClienteCorporativoAfexWeb")
     sURLIngresarDocumentoClienteCorporativoAfexWeb = Session("URLIngresarDocumentoClienteCorporativoAfexWeb")

    'FIN CUM-505 MS 02-02-2016

	Function CargarUltimoEnvio()
		Dim rsUG
		
		CargarUltimoEnvio = False		
		If Trim(Request.Form("txtExpress")) = "" Then Exit Function
		
		Set rsUG = ObtenerUltimosGiros(3, Request.Form("txtExpress"), "", 1)
		
		If rsUG Is Nothing Then Exit Function
		If Not rsUG.EOF Then
			sNombreB = MayMin(rsUG("nombre_beneficiario"))
			sApellidoB = MayMin(rsUG("apellido_beneficiario"))
			sNombreCompletoB = snombreB & " " & sApellidoB
			sDireccionB = MayMin(rsUG("direccion_beneficiario"))
			sFonoB = rsUG("fono_beneficiario")
			sPais = TRIM(evaluarvar(rsUG("pais_beneficiario"),""))
			sCiudad = evaluarvar(rsUG("ciudad_beneficiario"),"")
			sPagador = rsUG("agente_pagador")
			sMoneda = rsUG("codigo_moneda")
			nDDIPais = ObtenerDDI(1, sPais)
			nDDICiudad = ObtenerDDI(2, sCiudad)
			sPasaporteBeneficiario = evaluarvar(rsUG("pasaporte_beneficiario"),"")
			spaisPasapBeneficiario = evaluarvar(rsUG("paispasap_beneficiario"),"")
			sRutBeneficiario = evaluarvar(rsUG("Rut_beneficiario"),"")
			sCodigoBeneficiario = evaluarVar(rsUG("Codigo_beneficiario"),"")
		
			CargarUltimoEnvio = True
		End If
		Set rsUG = Nothing
	End Function
	
	Sub CalcularPesos
		
		nMontoDolar = cCur(0 & FormatNumber(Request.Form("txtMontoDolar"),2))
		nTipoCambio = cCur(0 & FormatNumber(Request.Form("txtTipoCambio"),4))
		nMontoPesos = cCur(0 & Round(nMontoDolar * nTipoCambio ))
		nCambioMG = cCur (0 & nCambioMG)

		If ntipocambio<> 0 then
			nTarifaPesos= (cCur(0 & nTarifa) * cCur(0 & Request.Form("txtTipoCambio")))
			nTarifaPesos = round(nTarifaPesos)
			nTotalPesos = round(nMontoPesos + nTarifapesos)
		End if
		
	End Sub
	
	sub CalcularNuevoMontoPesos
		Dim nCOmisionCP, nComisionSP, nPesosNew
	
	' calcular monto monto nuevo en pesos
		nPesosNew = FormatNumber(cCur(0 & nMonto) * cCur(0 & nCambioOp),0)				
		
		nComisionCP = FormatNumber(cCur(0 & ntarifaCobradaDolar) * cCur(0 & nCambioOp),0)	
		nComisionSP = FormatNumber(cCur(0 & ntarifaSugeridaDOlar) * cCur(0 & nCambioOp),0)	
		'nComisionCP = FormatNumber(cCur(0 & Request.form("txtTarifaCobrada")) * cCur(0 & nCambioOp),0)	
		'nComisionSP = FormatNumber(cCur(0 & Request.Form("txtTarifaSugerida")) * cCur(0 & nCambioOp),0)

		nTarifaCobradaPesos = cCur(0 & nComisionCP)					
		nTarifaSugeridaPesos = cCur(0 & nComisionSP)
		nMontoPesos = FormatNumber(cCur(0 & nMonto) * cCur(0 & nCambioOp),0)	
		nTotalPesos = nMontoPesos + nTarifaCobradaPesos
		
		'Response.Write nPesosNew
	end Sub
	
	Sub CalcularComisionDolar
		Dim nComisionNew, nComisionNew1

		nComisionNew = FormatNumber(cCur(0 & Request.Form ("txtTarifaCobradaPesos")) / cCur(0 & nCambioOp),2)				
		nComisionNew1 = FormatNumber(cCur(0 & Request.Form ("txtTarifaSugeridaPesos")) / cCur(0 & nCambioOp),2)				
		nComisionNew = FormatNumber(cCur(0 & ntarifaCobradapesos) / cCur(0 & nCambioOp),2)				
		nComisionNew1 = FormatNumber(cCur(0 & ntarifasugeridapesos) / cCur(0 & nCambioOp),2)				
		
		ntarifaCobradaDolar = Ccur(0 & nComisionNew)
		ntarifaSugeridaDOlar = cCur(0 & nComisionNew1)

	End Sub
	
	Sub CalcularNuevoMonto
		If Request.Form ("txttarifaCobrada") <> 0 Then
			nTarifaCobradaDolar = cCur(0 & Request.Form ("txttarifaCobrada"))
		end if
		If cCur(0 & Request.Form("txttarifaCobradapesos")) <> 0 Then
			nTarifaCobradapesos = FormatNumber(cCur(0 & Request.Form ("txttarifaCobradapesos")),0)
			
		end if
		ncambioOP = ccur(0 & formatnumber(Request.form("txttipocambioOP"),2))
		
		If Request.Form("OptPesos")= "on" then
			nMontoPesos = cCur(0 & formatnumber(Request.Form("txtpesos")))	
			nMOnto = cCur(0 & (nMontoPesos / ncambioOP))
			nPos1 = Instr(Request.form("txtMonto"), ",")
			
		Else
			ntarifacobrada= ccur(0 & Request.Form("txttarifacobrada"))
			nMonto = ccur(0 & formatNumber(Request.Form("txtmonto"),2))
			nMontoPesos = ccur (0 & (nMonto))* ccur(0 & nCambioOp)
			ntarifaCobradaPesos = formatNumber(ccur(0 & nTarifaCObradaDOlar ) * cCur(0 & nCambioOp),0)
			ntarifaSugeridaPesos = cCur(0 & nTarifaSugeridadolar) * cCur(0 & nCambioOP)
			nTotalPesos = formatNumber(cCur(0 & (nMontoPesos)) + cCur(0 & (nTarifaCobradaPesos)),0)			
			
		End If
	
	End Sub
	
		
	Sub Cargar
		sPais = Request("PaisB")
		sCiudad = Request("CiudadB")

		if sPais = "" then
			sPais = Request.Form("cbxPaisB")
			sCiudad = Request.Form("cbxCiudadB")
		end if
		if sCiudad = "" then sCiudad = Request.Form("txtCiudadBeneficiario")
		
		sPagador = Trim(Request("APagador"))
		If sPagador = "" then
		' ** cambio PSS 10-02-2010***************
			'If Request.Form("optMG") <> "on" Then
				sPagador = Request.Form("cbxPagador")
			'Else
			'	sPagador = Session("CodigoMGEnvio")
			'	bgironacional=true
			'End If
		End if
		
		sMoneda = Request("MonedaGiro")
		if sMoneda = "" then
			sMoneda = Request.Form("cbxMonedaGiro")
		end if
		
		'INTERNO-8479 MS 09-11-2016
        sMonedaPago = Request.Form("cbxMonedaPago")
		if sMonedaPago = "" then
			sMonedaPago = sMoneda
		end if
		'FIN INTERNO-8479 MS 09-11-2016
		
	'	Response.Write nmonto & ";" & Request.Form("txtMonto")
		nMonto = cDbl(0 & Request.Form("txtMonto"))		

		nDDIPais = ObtenerDDI(1, sPais)
		nDDICiudad = ObtenerDDI(2, sCiudad)
		sApellidos = Request.Form("txtApellidos")
		sNombreB = Request.Form("txtNombreB")
		sApellidoB = Request.Form("txtApellidoB")
		sDireccionB = Request.Form("txtDireccionB")
		sFonoB = Request.Form("txtFonoB")	    

	End Sub

	'INTERNO-8479 MS 09-11-2016
	Sub CalcularTarifa
        Dim sLugarPago
        sLugarPago = "1"
		' giro brasil
		If sPais = "BR" And (sPagador <> "ME" And sPagador <> Empty) Then sPagador = "AF"
		
		If sPagador = "ME" Then
			nMonto = FormatNumber(Request.Form("txtMontoDolar"),2)
			smoneda = Session("MonedaExtranjera")
		End If
		
		If  nMontoDolar<> 0 and session("perfil")= 1 then
			nmonto = nMontoDolar
		end if
		
		If sPais <> "" And sCiudad <> "" And sPagador <> "" And sMoneda <> "" And nMonto <> 0 Then
			if sMonedaPago = "" then
                sMonedaPago = sMoneda
            End if
			
            if Request.Form("cbxDeposito") = "2" then
                sLugarPago = "0"
            end if
            ObtenerTarifaGiros Session("CodigoAgente"), sPagador, sPais, sCiudad, sMoneda, sMonedaPago, _
											nMonto, nTarifa, nGastoTransfer, nComisionCaptador, nComisionPagador, _
											nComisionMatriz, nAfectoIva, sLugarPago
             
			nTarifaCobrada = nTarifa
			
			If Session("Perfil") = 1 Then
				nTarifaCobradaDolar=  nTarifa
				nTarifaSugeridaDOlar =  nTarifaCobradaDolar

			End If
		End If
	
	End Sub

	Sub CalcularTarifaCobrada
        Dim sLugarPago
        sLugarPago = "1"
		If sPais <> "" And sCiudad <> "" And sPagador <> "" And sMoneda <> "" And nMonto <> 0 Then
            if sMonedaPago = "" then
                sMonedaPago = sMoneda
            End if
			
            if Request.Form("cbxDeposito") = "2" then
                sLugarPago = "0"
            end if

			ObtenerTarifaGiros Session("CodigoAgente"), sPagador, sPais, sCiudad, sMoneda, sMonedaPago, _
											nMonto, nTarifa, nGastoTransfer, nComisionCaptador, nComisionPagador, _
											nComisionMatriz, nAfectoIva, sLugarPago
		End If
		nTarifaCobrada = CDbl(0 & Request.Form("txtTarifaCobrada"))
		
		If Session("Perfil") = 1 Then
			nTarifaCobradaDolar= nTarifaCobrada
			
		End If
	End Sub
 
	Sub CalcularTarifaUltimoGiro
        Dim sLugarPago
        sLugarPago = "1"
		If sPais <> "" And sCiudad <> "" And sPagador <> "" And sMoneda <> "" And nMonto <> 0 Then								
			'Response.redirect "../compartido/error.asp?description=" & Session("CodigoAgente") & ", " & sPagador & ", " & sPais & ", " & sCiudad & ", " & "USD" & ", " & sMoneda  & ", " &  _
			'								nMonto & ", " & nTarifa
            if sMonedaPago = "" then
                sMonedaPago = sMoneda
            End if

            if Request.Form("cbxDeposito") = "2" then
                sLugarPago = "0"
            end if

			ObtenerTarifaUltimoGiro Session("CodigoAgente"), sPagador, sPais, sCiudad, "USD", sMonedapago, _
											nMonto, nTarifa, nGastoTransfer, nComisionCaptador, nComisionPagador, _
											nComisionMatriz, nAfectoIva, sLugarPago
			nTarifaCobrada = nTarifa
		End If
				
	End Sub
	
	' JFMG 24-05-2012		
	Sub CargarOtrasMonedasPago(ByVal Seleccionado)
        Dim str
		Do While Not rsOtrasMonedasPago.EOF
            if rsOtrasMonedasPago("moneda_pago") = Seleccionado then
                str = "SELECTED"
            Else
                str = "" 
            End If
		    Response.write "<option value=" & rsOtrasMonedasPago("moneda_pago") & " " & str & " >" & rsOtrasMonedasPago("nombre_moneda") & "</option> "
		    
		    rsOtrasMonedasPago.MoveNext
		Loop
		
	End Sub
	' FIN JFMG 24-05-2012
	'FIN INTERNO-8479 MS 09-11-2016

%>
<!--#INCLUDE virtual="/Compartido/Rutinas.asp" -->
<!--#INCLUDE virtual="/Compartido/Errores.asp" -->
<html>
<head>
    <meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
    <link rel="stylesheet" type="text/css" href="../Estilos/Cliente.css">
</head>
<script LANGUAGE="VBScript">
<!--

	Sub window_onLoad()
		Dim sCadena
		
		If Not CalcularTotal Then Exit Sub
		ActivarFoco
		<%If sPagador = Session("CodigoMGEnvio") Then %>
			'frmGiro.optMG.value = "on"
			''rmGiro.optMG.checked = true
			'frmGiro.optOtroAgt.value = ""
			'frmGiro.optOtroAgt.checked = false
			'frmGiro.cbxPagador.disabled = true
		<%Else%>
			'frmGiro.optMG.value = ""
			'frmGiro.optMG.checked = false
			frmGiro.optOtroAgt.value = "on"
			frmGiro.optOtroAgt.checked = true
		<% End If %>
		colormoneda "<%=sColorMoneda%>"
		
		If "<%=bCargarMensaje%>" Then
			If frmGiro.cbxPagador.value <> Empty Then
				MsgBox "<%=sNota%>" & vbCrLf & "<%=sMaximo%>" & vbCrLf & "<%=sRecomendacion%>" & vbCrLf & "<%=sRequerimiento%>", vbInformation, "Recordatorio"
			End If
		End If
	 
		' giro brasil		
		If frmGiro.cbxPaisB.value = "BR"  And ("<%=sPagador%>" <> "ME" And "<%=sPagador%>" <> Empty) Then
			frmGiro.cbxPagador.value = "AF" 
			frmGiro.cbxPagador.disabled = True
		Else
			frmGiro.cbxPagador.disabled = False
		End If
		
		<% If session("Perfil") = 1 and Not bGiroNacional and session("CodigoMGEnvio")<> sPagador then %>
			frmGiro.txttotalpesos.value = "<%=FormatNumber(nTotalPesos,0)%>"
			frmGiro.txttarifacobradapesos.value ="<%=FormatNumber(ntarifacobradaPesos,0)%>"

 			<% If request.form("optDOlar") = "on" then %>
				frmGiro.OptDolar.checked = true
				frmGiro.Optpesos.checked = false
				
			<% End if %>
			<% If request.form("optPesos") ="on" Then %>
				frmGiro.OptPesos.checked = True 
				frmGiro.OptDolar.checked = false	
				
			<% End IF %>	
		<% End If %>
		
		
		' JFMG 24-05-2012
		<%If bOtrasMonedasPago Then %>
			'INTERNO-8479 MS 09-11-2016
            if "<%=sPagador %>" <> "SW" Then
		    window.filaOpcionMonedaPago.style.display = "block"		
		    frmGiro.cbxMonedaPago.disabled = False
		    If "<%=request.form("chkOtraMonedaPago")%>" = "on" Then
		        frmGiro.chkOtraMonedaPago.checked = True
		        chkOtraMonedaPago_onClick()
		    End If
            end if
			'FIN INTERNO-8479 MS 09-11-2016
		<%End If %>
		' FIN JFMG 24-05-2012
		
		'MS 26-07-2013
		If <%=Session("Categoria")%> > 2 Then	
			frmGiro.txtTarifaCobrada.disabled = true
			'frmGiro.txtTarifaCobradaPesos.disabled = true
			frmGiro.txtTarifaSugerida.disabled = true
			'frmGiro.txtTarifaSugeridaPesos.disabled = true
		End If
		'MS FIN 26-07-2013
	End Sub
	
	' JFMG 25-05-2012
	Sub chkOtraMonedaPago_onClick()
	    If frmGiro.chkOtraMonedaPago.checked Then
	        window.filaMonedaPago.style.display = "block"
	    else
	        window.filaMonedaPago.style.display = "none"
	    end if 
	    
	End Sub
	' FIN JFMG 25-05-2012
	
	Sub ActivarFoco()
		<% If bGiroAnterior Then %>
			<% If Session("Perfil") = 1 Then %>
			
					<% If request.Form("optPesos") = "on"   Then %>
						frmGiro.txtpesos.focus
						frmGiro.txtpesos.select
					<% end if %>
					<% If request.Form("optDolar") = "on"   Then %>
						frmGiro.txtmonto.focus
						frmGiro.txtmonto.select
						
					<% End If %>
				
			<% else %>
					<% If sPagador="ME" then %>
						frmGiro.txtMontoDolar.focus
						frmGiro.txtMontoDolar.select
					<% Else %>
						frmgiro.txtMonto.focus
						frmGiro.txtMonto.select

					<% End If %>	
			<% End If %>	
		<% End If %>			
	Exit Sub
	
		
		<% Select Case nAccion %>
		<%	Case afxAccionPais %>
				frmGiro.cbxCiudadB.focus 
				
		<%	Case afxAccionCiudad %>
				frmGiro.txtfonoB.select
				
		<%	Case afxAccionPagador %>
			<% If sPagador="ME" then %>
				frmGiro.txtMontoDolar.select
			<% Else %>
				frmGiro.txtMonto.select
			<% End If %>	
		
			
		<%	Case afxAccionMonto %>
			
				<% If bExtranjero Or request.Form("optMG") = "on" Then %>
					<% If nTipoCambio <>0  Then %>
					
						frmGiro.txtInvoiceMG.select
					<% Else %>
						'frmGiro.txtTipoCambio.select 
					<% End If %>
				<% Else %>
						frmGiro.txtBoleta.select
				<% End If %>
			
		<%	Case afxAccionTarifa %>
				<% If bExtranjero Or request.Form("optMG") = "on" Then %>
						frmGiro.txtTipoCambio.select
						frmGiro.txtTipoCambio.focus	
						'frmGiro.txtInvoiceMG.select				
				<% Else %>
						frmGiro.txtBoleta.select
				<% End If %>
				
		<% Case afxAccionDeposito %>
			frmGiro.txtmonto.select 
			
		<% Case Else %>
				<% If Session("Categoria") = 4 Then %>
						frmGiro.txtNombres.select
				<% Else %>
						frmGiro.txtNombreB.select
				<% End If %>
		<% End Select %>		
	End Sub 
	
	Sub txtTipoCambio_onBlur()
		If "<%=sPagador%>" = "<%=session("CodigoMGEnvio")%>" and "<%=session("Categoria")%>" <> 3 Then
			If cCur(0 & frmGiro.txtTipoCambio.value)  < cCur(0 & "<%=nCambioMG%>") Then
					frmGiro.txtTipoCambio.value = cCur(0 & frmGiro.txtCambioMG.value)
					MsgBox "El Tipo CambioMG no puede ser mayor al Tipo de cambio Ingresado. El giro no puede ser enviado ", ,"AFEX"
					
			End If
		End If
		If cCur(0 & frmGiro.txtTipoCambio.value) > 0 Then
		'msgbox "<%=nMontoPesos%>" & "," & "<%=nTarifaPesos%>" & "," & "<%=nTotalPesos%>"
			frmGiro.txtMontoPesos.value = "<%=nMontoPesos%>"
			frmGiro.txtTarifaPesos.value = "<%=nTarifaPesos%>"
			frmGiro.txtTotalPesos.value = "<%=formatNumber(nTotalPesos,0)%>"
			frmGiro.txtInvoiceMG.Select 
		End If
		
			HabilitarControles
			frmGiro.action = "EnviarGiro.asp?Accion=<%=afxAccionMonto%>"
			frmGiro.submit 
			frmGiro.action = ""
	End Sub
	
	Sub txtMontoDolar_onBlur()
		Dim nPos
		
		If cCur(0 & Trim(frmGiro.txtMontoDolar.value)) = 0 Then
			frmGiro.txtMontoDolar.value = "<%=FormatNumber(0, nDec)%>"
			frmGiro.txtTipoCambio.value = "<%=FormatNumber(0, nDec)%>"
			
		Else
			nPos = Instr(frmGiro.txtMontoDolar.value, ",")
			
		End If
			HabilitarControles
			frmGiro.action = "EnviarGiro.asp?Accion=<%=afxAccionMonto%>"
			frmGiro.submit 
			frmGiro.action = ""
		
	End Sub
	
	Sub txtMonto_onBlur()
		Dim nPos
		
		If cCur(0 & Trim(frmGiro.txtMonto.value)) = 0 Then
			frmGiro.txtMonto.value = "<%=FormatNumber(0, nDec)%>"
			frmGiro.txtTarifaSugerida.value = "<%=FormatNumber(0, nDec)%>"
			frmGiro.txtTarifaCobrada.value = "<%=FormatNumber(0, nDec)%>"
			frmGiro.txtTotal.value = "<%=FormatNumber(0, nDec)%>"
			Exit Sub
		Else
			nPos = Instr(frmGiro.txtMonto.value, ",")
			If nPos > 0 Then
				If cCur(0 & Mid(frmGiro.txtMonto.value, nPos)) > 0 Then
					msgbox "El monto del giro no puede incluir decimales"
					frmGiro.txtMonto.select
					Exit Sub
				End If
			End If
            'CUM-480 MS 24-08-2015
            if trim(frmGiro.cbxPaisB.value) = "CO" then
                if frmGiro.cbxMonedaGiro.value = "USD" then
                    if cCur(0 & Trim(frmGiro.txtMonto.value)) >= <%= session("MontoTopeGirosColombia")%> then 'CUM-505 MS 02-02-2016  
                        window.showModalDialog("Mensaje.asp?Mensaje=El monto del giro EXCEDE el umbral de Control")
                    else
                        nMontoSumatoriaCumplimiento = frmGiro.txtAcumuladoCumplimiento.value
                        if cCur(0 & nMontoSumatoriaCumplimiento) + cCur(0 & Trim(frmGiro.txtMonto.value)) >= <%= session("MontoTopeGirosColombia")%> then 'CUM-505 MS 02-02-2016
                            window.showModalDialog("Mensaje.asp?Mensaje=La SUMA de los giros excede el límite mensual de operación")
                        end if 
                    end if
                end if
            end if
            'FIN CUM-480 MS 24-08-2015
            
            'CUM-505 MS 02-02-2016
            'SMC-53 MS 17-03-2016
            if Trim(frmGiro.txtPaisPasaporteRemitente.value) = "CO" Then 'pasaporte Pais colombiano
                'Cliente presenta pasaporte colombiano
                dim sMensajeCumplimiento, sPregunta, sRespuestaPregunta, nMontoMayorTope
                dim CuentaUsuarioEncriptado, ContrasenaUsuarioEncriptado, CuentaSucursalEncriptado, ContrasenaSucursalEncriptado
                dim sCodClienteGiros, sCodClienteCorp, sMonedaGiroEncriptada, sMontoGiroEncriptado, sTarifaGiroEncriptada
                dim sNacionalidad, sPaisDestino, sPasaporteCliente

                CuentaUsuarioEncriptado =  RemplazaLetra("<%=EncriptarCadena(trim(Session("NombreUsuarioOperador")))%>") 
				ContrasenaUsuarioEncriptado = RemplazaLetra("<%=EncriptarCadena(trim(Session("ContrasenaOperador")))%>")
    			CuentaSucursalEncriptado = RemplazaLetra("<%=EncriptarCadena(trim(Session("NombreUsuarioAgente")))%>")
				ContrasenaSucursalEncriptado = RemplazaLetra("<%=EncriptarCadena(trim(Session("ContrasenaAgente")))%>")
                sCodClienteGiros = trim(frmGiro.txtExpress.value)
                sCodClienteCorp = frmGiro.txtExchange.value
                sMontoGiroEncriptado = cCur(0 & Trim(frmGiro.txtMonto.value))
                sTarifaGiroEncriptada = cCur(0 & Trim(frmGiro.txtTarifaCobrada.value))
                sMonedaGiroEncriptada = Trim(frmGiro.cbxMonedaGiro.value)
                If frmGiro.cbxPaisB.value <>"CL" then
                    sNacionalidad = 1
                Else
                    sNacionalidad = 2
                End If
                sPaisDestino =  Trim(frmGiro.cbxPaisB.value)
                sPasaporteCliente = Trim(frmGiro.txtPasapRemitente.value)

                sFeaturesURL = "height=600,width=730,left=100,top=50,resizable=no,scrollbars=yes, title=eAFEX"
                
                if cCur(0 & Trim(frmGiro.txtMonto.value)) > <%= session("MontoTopeGirosPasapColombiano")%> then 
                    'Cliente opera por mas de 300USD
                    sMensajeCumplimiento = "<b>Origen Cumplimiento</b><br/><br/>Para esta operación se debe agregar RUT del cliente.<br/>Si no es posible, debe comunicarse con el departamento de Cumplimiento<br/>"
                    sPregunta="<center>¿Desea actualizar ahora?</center>"

                    If  CInt(frmGiro.txtAutorizadoOperarPasaporte.value) = 1 then 
                        'Cliente se encuentra activo en la corporativa y con el pasaporte autorizado
                        frmGiro.txtAccionCumplimiento.value = "1"
                    else
                        if window.showModalDialog("MensajeCumplimiento.asp?mensajeCump=" & sMensajeCumplimiento & "&pregunta=" & sPregunta & "&MostrarBotones=display")  = "1" then
                            'Cajero decide actualizar datos de cliente en corporativa
                            frmGiro.txtAccionCumplimiento.value = "1"
                      
                            window.open "<%=sURLIngresarClienteCorporativoAfexWeb%>" & "?CuentaUsuario=" & CuentaUsuarioEncriptado & _
                            "&ContrasenaUsuario=" & ContrasenaUsuarioEncriptado & _
                            "&CuentaSucursal=" & CuentaSucursalEncriptado & "&ContrasenaSucursal=" & ContrasenaSucursalEncriptado & _								
                            "&IP=" & "<%=request.servervariables("REMOTE_ADDR")%>" & "&codCteGiro=" &  sCodClienteGiros & "&codCteCorp=" & sCodClienteCorp & "&PasaporteCliente=" & sPasaporteCliente, "",sFeaturesURL
                            frmGiro.txtAccionCumplimiento.value = "4"
                            Pause(5)
                        else
                            'Cajero decide no actualizar datos de cliente en corporativa                        
                            frmGiro.txtAccionCumplimiento.value = "2"
                            Pause(2)
                        End if

                    end if
                    
                else
                    if trim(frmGiro.txtExpress.value) <> "" then
                        'monto es menor o igual a 300 USD
                        sMensajeCumplimiento = "<b>Origen Cumplimiento</b><br/><br/>Debe actualizar al cliente para poder operar<br/>"
                        sPregunta="<center>¿Desea actualizarlo ahora?</center>"                
                            sFeaturesURL = "height=600,width=650,left=100,top=50,resizable=no,scrollbars=yes, title=eAFEX"
                        if cInt(frmGiro.txtAutorizadoOperarPasaporte.value) = 2 then 'Cliente no existe en corporativa
                            if window.showModalDialog("MensajeCumplimiento.asp?mensajeCump=" & sMensajeCumplimiento & "&pregunta=" & sPregunta & "&MostrarBotones=display")  = "1" then
                                frmGiro.txtAccionCumplimiento.value = "1"
                      
                                window.open "<%=sURLIngresarClienteCorporativoAfexWeb%>" & "?CuentaUsuario=" & CuentaUsuarioEncriptado & _
                                "&ContrasenaUsuario=" & ContrasenaUsuarioEncriptado & _
                                "&CuentaSucursal=" & CuentaSucursalEncriptado & "&ContrasenaSucursal=" & ContrasenaSucursalEncriptado & _								
                                "&IP=" & "<%=request.servervariables("REMOTE_ADDR")%>" & "&codCteGiro=" &  sCodClienteGiros & _
                                "&codCteCorp=" & sCodClienteCorp & "&PasaporteCliente=" & sPasaporteCliente, "",sFeaturesURL
                                frmGiro.txtAccionCumplimiento.value = "4"
                                Pause(5)
                            else
                                frmGiro.txtAccionCumplimiento.value = "0"
                                Pause(2)
                            End if
                        else                                                  
                            'Verifica si tiene ingresado y autorizado el pasaporte 
                            if frmGiro.txtPasaporteIngresadoYAprobado.value = "0" then
                                sMensajeCumplimiento = "<b>Origen Cumplimiento</b><br/><br/>Cliente identificado con pasaporte, debe escanear este documento y adjuntar la imagen de los antecedentes del cliente.<br/>"
                                sPregunta="<center>¿Desea ingresarlo ahora?</center>"
                                
                                if window.showModalDialog("MensajeCumplimiento.asp?mensajeCump=" & sMensajeCumplimiento & "&pregunta=" & sPregunta & "&MostrarBotones=display")  = "1" then
                                    'Abre la ventana para cargar antecedentes...
                                    window.open "<%=sURLIngresarDocumentoClienteCorporativoAfexWeb%>"  & "?CuentaUsuario=" & CuentaUsuarioEncriptado & _
                                    "&ContrasenaUsuario=" & ContrasenaUsuarioEncriptado & _
                                    "&CuentaSucursal=" & CuentaSucursalEncriptado & "&ContrasenaSucursal=" & ContrasenaSucursalEncriptado & _								
                                    "&IP=" & "<%=request.servervariables("REMOTE_ADDR")%>" & "&codCteGiro=" &  sCodClienteGiros & "&codCteCorp=" & sCodClienteCorp & _
                                    "&monedaGiro=" & sMonedaGiroEncriptada & "&montoGiro=" & sMontoGiroEncriptado & "&tarifaGiro=" & sTarifaGiroEncriptada & _
                                    "&origen=1&paisDestino=" & sPaisDestino & "&nacionalidad=" & sNacionalidad, "",sFeaturesURL
                                    Pause(5)
                                    frmGiro.txtAccionCumplimiento.value = "3"
                                else
                                    Pause(2)
                                    frmGiro.txtAccionCumplimiento.value = "0"
                                end if
                            else
                                frmGiro.txtAccionCumplimiento.value = "1"
                            End if 
                        End if
                    else
                        'MsgBox "Cliente no existe en giros. Debe ingresarlo primero para continuar con la operación."
                        frmGiro.txtAccionCumplimiento.value = "5"
                    end if
                End if
            else 
                'Cliente no opera con pasaporte Colombiano
                frmGiro.txtAccionCumplimiento.value = "1"
            end if
            'FIN SMC-53 MS 17-03-2016
            'FIN CUM-505 MS 02-02-2016

		End If
		
		<% If Session("Perfil") = 1  and not bGiroNacional Then %>
					HabilitarControles
					frmGiro.action = "EnviarGiro.asp?Accion=<%=afxAccionNuevo%>" 
					frmGiro.submit 
					frmGiro.action = ""
		<% else %>		
				<% If not bExtranjero  Then %>
					HabilitarControles
					frmGiro.action = "EnviarGiro.asp?Accion=<%=afxAccionMonto%>"
					frmGiro.submit 
					frmGiro.action = ""
				<% End IF %>
		<% End IF %>
			
	End Sub 
	'CUM-505 MS 02-02-2016
    function RemplazaLetra(Cadena)							
	    cadena = replace(cadena, "Ñ", "%c3%91")
	    RemplazaLetra = replace(cadena, "ñ", "%c3%b1")
    end function

    sub Pause(nSegundos)
        starttime = Time()
        Do until Datediff("s",starttime,Time(),0,0) > nSegundos

        Loop
    end sub

    'FIN CUM-505 MS 02-02-2016

	Sub txtTarifaCobradaPesos_onBlur()
		
		HabilitarControles
		frmGiro.action = "EnviarGiro.asp?Accion=<%=afxAccionNuevo%>"
		frmGiro.submit 
		frmGiro.action = ""
		
	End Sub
	
	Sub txtPesos_onBlur()
	Dim nPos
	
		If ccur(0 & trim(frmGiro.txtpesos.value))=0 Then
			frmGiro.txtpesos.value ="<%=FormatNumber(0, 0)%>"
			frmGiro.txttarifasugeridapesos.value = "<%=FormatNumber(0, 0)%>"
			frmGiro.txttarifacobradapesos.value = "<%=FormatNumber(0, 0)%>"
			frmGiro.txtTotalPesos.value = "<%=FormatNumber(0, 0)%>"
			
		End IF
		Dim n
		HabilitarControles
		frmGiro.action = "EnviarGiro.asp?Accion=<%=afxAccionNuevo%>"
		frmGiro.submit 
		frmGiro.action = ""
		
		nPOs = instr(frmGiro.txtmonto.value,",")
		
		If nPOs > 0 then
			msgbox  "El monto en Dolares no puede tener decimales será aproximado " & nPos
			HabilitarControles
			frmGiro.action = "EnviarGiro.asp?Accion=<%=afxAccionNuevo%>&nPOs=" & nPOs 
			'frmGiro.action = "EnviarGiro.asp?Accion=<%=afxAccionMonedaPago%>&PaisB=" & aGiro(3) & "&CiudadB=" & aGiro(4)
			frmGiro.OptDolar.checked = true 
			frmGiro.submit 
			frmGiro.action = ""
			
		End If 
		
	End Sub

	Sub txtTarifaSugerida_onfocus()
		frmGiro.txtTarifaCobrada.focus
	End Sub
		
	Sub txtTarifaCobrada_OnBlur()
        '28-08-2015 INTERNO-4280 MM
        If cCur(frmGiro.txtTarifaCobrada.value) > cCur(frmGiro.txtTarifaSugerida.value) Then
            frmGiro.txtTarifaCobrada.value = FormatNumber(frmGiro.txtTarifaCobrada.value, 2)
            frmGiro.txtTarifaCobrada.style.backgroundColor = "#FF0000"
            Exit Sub
        else
            frmGiro.txtTarifaCobrada.style.backgroundColor = "#4dc087"
        End If  

		If cCur(0 & Trim(frmGiro.txtMonto.value)) = 0 Then
			frmGiro.txtMonto.value = "<%=FormatNumber(0, nDec)%>"
			frmGiro.txtTarifaSugerida.value = "<%=FormatNumber(0, nDec)%>"
			frmGiro.txtTarifaCobrada.value = "<%=FormatNumber(0, nDec)%>"
			frmGiro.txtTotal.value = "<%=FormatNumber(0, nDec)%>"
			Exit Sub
		ElseIf Trim(frmGiro.txtTarifaCobrada.value) = "" Then
			frmGiro.txtTarifaCobrada.value = "<%=FormatNumber(0, nDec)%>"
		Else
			
		End If

		<% If session("Perfil") = 1 Then %>
				HabilitarControles
				frmGiro.action = "EnviarGiro.asp?Accion=<%=afxAccionNuevo%>"
				frmGiro.submit 
				frmGiro.action = ""
		<% else %>
		
			<% If Not bExtranjero Then %>
				HabilitarControles
				frmGiro.action = "EnviarGiro.asp?Accion=<%=afxAccionTarifa%>"
				frmGiro.submit 
				frmGiro.action = ""
			<% End IF %>
		<% End IF %>
	End Sub 

	Sub txtTotal_onFocus()
		frmGiro.txtBoleta.focus
	End Sub
	
	Sub imgCalcular_onClick()
		CalcularTotal
	End Sub

	Function CalcularTotal()
		Dim nCobrada, nMonto
		
		CalcularTotal=False
		
		<% If session("Perfil")= 1 and Not bGiroNacional and session("CodigoMGEnvio") <> sPagador then %>
			frmGiro.txtTarifaSugeridapesos.value = "<%=FormatNumber(nTarifaSugeridaPesos, 0)%>"
			'frmGiro.txttarifaCobradaPesos.value ="<%=FormatNumber(nTarifaCobradaPesos, 0)%>"
			frmGiro.txtTotalPesos.value = "<%=FormatNumber(nMontoPesos + nTarifaCobradaPesos, 0)%>"
			frmgiro.txttarifasugerida.value ="<%=FormatNumber(nTarifaSUgeridaDolar, nDec)%>"
			frmGiro.txttotal.value = "<%=FormatNumber(nMonto + nTarifaCobradaDOlar, nDec)%>"
			
			If frmGiro.OptPesos.checked = True Then
				<% If nTarifaCobradaPesos = 0 and nMontoPesos > 0 then %>
					MsgBox "La tarifa cobrada debe ser mayor que cero", ,"AFEX"
					Exit Function						  
				<% End If %>
			End If
			If frmGiro.OptDolar.checked = true Then
				<% If nTarifaCobradadolar = 0 and nMonto > 0 then %>
					MsgBox "La tarifa cobrada debe ser mayor que cero", ,"AFEX"
					Exit Function
			<% End If %>
			End If
		<% Else	%>	
			<% IF spagador = "ME" Then %>
				frmGiro.txtTarifaSugerida.value = "<%=FormatNumber(nTarifa, 2)%>"
			<% Else %>	
				frmGiro.txtTarifaSugerida.value = "<%=FormatNumber(nTarifa, nDec)%>"
			<% End If %>
			frmGiro.txtTotal.value = "<%=FormatNumber(nMonto + nTarifaCobrada, nDec)%>"
		<% End If %>	
		
		<%	If Not bGiroNacional and session("Perfil") <> 1 Then
				If nTarifaCobrada = 0 And nMonto > 0 Then %>
					MsgBox "La tarifa cobrada debe ser mayor que cero", ,"AFEX"
					Exit Function
			<%	ElseIf nTarifaCobrada < nGastoTransfer Then %>
					MsgBox "La tarifa cobrada <%=nTarifaCobrada%> no debe ser menor que los gastos de transferencia <%=nGastoTransfer%>", ,"AFEX"
					Exit Function
		<%		End If %>
		<% Else
				If session("Perfil")= 1  Then %>
					frmGiro.cbxmonedagiro.disabled= false
			<%	End IF
			End If	%>
		CalcularTotal = True
	End Function
 
	Sub cbxPaisB_onblur()
		Dim sCiudad

		If frmGiro.cbxPaisB.value = "" Then Exit Sub
		If frmGiro.cbxPaisB.value = "<%=sPais%>" Then Exit Sub
		HabilitarControles
		if sPagador="ME" Then
			frmGiro.txtMontoDolar.value = "<%=FormatNumber(0, 2)%>"
			frmGiro.txtTipoCambio.Value = "<%=FormatNumber(0, nDec)%>"
		Else
			frmGiro.txtMonto.value = "<%=FormatNumber(0, nDec)%>"
		End if
		frmGiro.cbxCiudadB.value = ""
		frmGiro.cbxPagador.value = ""
		frmGiro.cbxMonedaPago.value = ""
		frmGiro.action = "EnviarGiro.asp?Accion=<%=afxAccionPais%>"
		frmGiro.submit 
		frmGiro.action = ""
		
	End Sub
	
	Sub cbxCiudadB_onblur()
		Dim sComuna
		
		If frmGiro.cbxCiudadB.value = "" Then Exit Sub
		If frmGiro.cbxCiudadB.value = "<%=sCiudad%>" Then Exit Sub		
		HabilitarControles
		If sPagador="ME" Then
			frmGiro.txtMontoDolar.value = "<%=FormatNumber(0, 2)%>"
			frmGiro.txtTipoCambio.value = "<%=FormatNumber(0, nDec)%>"
		Else
			frmGiro.txtMonto.value = "<%=FormatNumber(0, nDec)%>"
		End IF
		frmGiro.cbxPagador.value = ""
		frmGiro.cbxMonedaPago.value = ""
		frmGiro.action = "EnviarGiro.asp?Accion=<%=afxAccionCiudad%>"
		frmGiro.submit 
		frmGiro.action = ""
	End Sub

	Sub cbxPagador_onChange() 'JFMG 24-05-2012 _onblur()
		Dim sComuna, sCadena
				
		If frmGiro.cbxPagador.value = "" Then Exit Sub
		If frmGiro.cbxPagador.value = "<%=sPagador%>" Then Exit Sub
		
		if frmGiro.cbxPagador.value ="<%=Session("CodigoMGEnvio")%>" then
			Colormoneda "DodgerBlue"
		else
			colormoneda "#4dc087"
		End IF

        'miki SMC-9 MM 2015-11-30
        If frmGiro.cbxPagador.value = "SW" then
            MsgBox "Monto mínimo de envío para este agente es US$ 25",vbInformation,"Recordatorio"
        End if
		
		HabilitarControles
		frmGiro.cbxMonedaPago.value = ""
		frmGiro.action = "EnviarGiro.asp?Accion=<%=afxAccionPagador%>"
		frmGiro.submit 
		frmGiro.action = ""
		
	End Sub
	
	Sub ColorMoneda (byval color)
	
	End sub
 
	Sub cbxMonedaPago_onChange() '_onblur()
		Dim sComuna
		
		If frmGiro.cbxMonedaPago.value = "" Then Exit Sub
		If NOT <%=bOtrasMonedasPago%> then If frmGiro.cbxMonedaPago.value = "<%=sMoneda%>" Then Exit Sub		
		HabilitarControles
		frmGiro.action = "EnviarGiro.asp?Accion=<%=afxAccionMonedaPago%>"
		frmGiro.submit 
		frmGiro.action = ""
	End Sub
	
    'INTERNO-8479 MS 09-11-2016
	Sub cbxdeposito_onChange()
        if (frmGiro.cbxdeposito.value = "1" or frmGiro.cbxdeposito.value = "2")  and Trim(frmGiro.cbxPagador.value) = "SW" and Trim(frmGiro.cbxpaisb.value) = "DO" then
            frmGiro.cbxmonedapago.value = "DOP"        
        end if 
           HabilitarControles
		frmGiro.action = "EnviarGiro.asp?Accion=<%=afxAccionMonedaPago%>"
		frmGiro.submit 
		frmGiro.action = ""
	End Sub
	'FIN INTERNO-8479 MS 09-11-2016

	Sub cbxMonedaGiro_onBlur()
		Dim sColor
		If frmGiro.cbxMonedaGiro.value = "<%=Session("MonedaNacional")%>" and frmGiro.cbxmonedagiro.value ="ME" Then
			sColor = "DodgerBlue"
			frmGiro.txtMontoDolar.style.backgroundColor=sColor
			frmGiro.txtTipoCambio.style.backgroundColor=sColor
		Else		
			sColor = "#4dc087"
		End If
		If frmGiro.cbxmonedagiro.value ="<%=Session("MonedaNacional")%>" and frmGiro.cbxPaisB.value ="CL" Then
			sColor = "DodgerBlue"
		END IF
		frmGiro.cbxMonedaGiro.Style.backgroundColor = sColor
		frmGiro.txtMonto.Style.backgroundColor = sColor
		frmGiro.txtTarifaSugerida.Style.backgroundColor = sColor
		frmGiro.txtTarifaCobrada.Style.backgroundColor = sColor
		frmGiro.txtTotal.Style.backgroundColor = sColor
		
	End Sub
	
	'CP INICIO 04/04/2012. Para que recotice cuando se cambia la moneda de envío
	Sub cbxMonedaGiro_onChange()
		HabilitarControles
		frmGiro.action = "EnviarGiro.asp?Accion=<%=afxAccionMonto%>"
		frmGiro.submit 
		frmGiro.action = ""
	End Sub
	'CP FIN 
	
	Sub imgAceptar_OnClick()
		Dim sGiroBrasil
		dim i
		dim sBancoBR, sAgenciaBR, sCtaCteBR, sCpfBR
		
		If Not CajaPregunta("AFEX En Linea", "Está seguro que desea enviar el giro?") Then
			Exit Sub
		End If

		If Not CalcularTotal Then
			Exit Sub
		End If
		
		'Validaciones 
		If Not ValidarDatos Then
			Exit Sub
		Else
			
			' si el giro es para brasil solicita otros datos
			If trim(frmGiro.cbxPaisB.value) = "BR" And Trim(frmGiro.cbxPagador.value) = "AF" then
				IF 	trim(frmGiro.cbxPaisB.value) = "BR" THEN
					frmGiro.txtGiroBrasil.value = window.showModalDialog("datosgirobrasil.asp?Dolares=" & frmGiro.txtMonto.value)				
					If trim(frmGiro.txtGiroBrasil.value) = "" Then
						msgbox "Debe ingresar los datos antes solicitados."					
						exit sub
					End If
					
					' Jonathan Miranda G. 04-06-2007
					sGiroBrasil = frmGiro.txtGiroBrasil.value
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
						
					end if
					
					    If trim(sBancoBR) = "" or trim(sAgenciaBR) = "" or trim(sCtaCteBR) = "" or trim(sCpfBR) = "" Then
						    msgbox "Debe ingresar todos los datos solicitados para el envío.2",,"AFEX"
						    exit sub
					    End If
				END IF
				'-------------------- Fin -------------------------------
				
			END IF 	
		<% If  session("Perfil")<> 1 Then %>
			<% If sPagador <> session("CodigoMGEnvio")  then %>
                'miki SMC-9 MM 2015-11-30
				'INTERNO-8479 MS 09-11-2016
				If (frmGiro.cbxDeposito.value = 1) or (Trim(frmGiro.cbxPagador.value) = "SW")  Then
                    dim monedaPago, sFormaPago                  
                    monedaPago = ""
                    sFormaPago = frmGiro.cbxdeposito.value
                    If Trim(frmGiro.cbxPagador.value) = "SW"  then
                        If Trim(frmGiro.cbxMonedaPago.value) <> "" then
                            monedaPago = Trim(frmGiro.cbxMonedaPago.value)
                        Else
                            monedaPago = Trim(frmGiro.cbxMonedaGiro.value)
                        End if
                        IF frmGiro.cbxpaisB.value = "CO" Then
                            monedaPago = "COP"
                        End If

                    End If
                    
                    'MsgBox "forma de pago= " & sFormaPago

					frmGiro.txtDatosDeposito.value = window.showModalDialog( "GiroDeposito.asp?Pagador=" & frmGiro.cbxPagador.value & "&FormaPago=" & sFormaPago & "&Pais=" & frmGiro.cbxpaisB.value & "&MonedaPago=" & monedaPago & "&TelefonoB=" & frmGiro.txtfonob.value & "&DireccionB=" & frmGiro.txtdireccionb.value)
					If trim(frmGiro.txtDatosDeposito.value) = "" Then
                        if frmGiro.cbxdeposito.value  = "1" then
						msgbox "Debe ingresar los datos de Depósito.",, "AFEX"					
                        else
                            msgbox "Debe ingresar todos los datos solicitados para el envío.",, "AFEX"                            
                        end if
						Exit sub
					End If
					 
                        			
					 sDatosDeposito = frmGiro.txtDatosDeposito.value
					i = instr(sDatosDeposito,  ";")
    
					if i > 0 then
						sBancoBR = left(sDatosDeposito, i - 1) 
						sDatosDeposito = mid(sDatosDeposito, i + 1)
						i = instr(sDatosDeposito, ";")
						sTipoCuenta = left(sDatosDeposito, i - 1) 
						sDatosDeposito = mid(sDatosDeposito, i + 1)
						i = instr(sDatosDeposito, ";")
						sCtaCteBR = left(sDatosDeposito, i - 1)
						sDatosDeposito = mid(sDatosDeposito,i + 1)
						i = instr(SDatosDeposito,";")
						sMonedaDeposito = left(sDatosDeposito,i - 1)
                        sDatosDeposito = mid(sDatosDeposito,i + 1)
                        i = instr(SDatosDeposito,";")
						
                        if frmGiro.cbxdeposito.value  = "2" then
						    frmGiro.txtdireccionb.value = left(sDatosDeposito, i - 1)
                            sDatosDeposito = mid(sDatosDeposito, i + 1)
						    frmGiro.txtFonoB.value = sDatosDeposito'mid(sDatosDeposito)
                        end if
                        'MsgBox "tipo  cta->" & sTipoCuenta & ", nro Cta->" & sCtaCteBR & ", moneda dep->" & sMonedaDeposito & ", banco->" & sBancoBR
					end if
					'FIN INTERNO-8479 MS 09-11-2016
					' JFMG 25-05-2012
					If "<%=sPagador%>" = "IK" Then
					    If trim(sBancoBR) = "" or trim(sTipoCuenta) = "" or trim(sCtaCteBR) = ""Then
						    msgbox "Debe ingresar todos los datos solicitados para el envío.1",,"AFEX"
						    exit sub
					    End If
					ELse	
					' FIN JFMG 25-05-2012
					   'INTERNO-3855 MS 26-04-2015
						If "<%=sPagador%>" = "FP" Then
					        If trim(sBancoBR) = "" or trim(sCtaCteBR) = ""Then
					            msgbox "Debe ingresar todos los datos solicitados para el envío.1",,"AFEX"
					            exit sub
					         else
					            if  instr(1, ";13;14;15;",  ";" & sBancoBR & ";")<=0 then
					                If trim(sTipoCuenta) = "" Then
					                    msgbox "Debe ingresar todos los datos solicitados para el envío.2",,"AFEX"
					                    exit sub
					                end if
					            end if 
					         end if 
						 else
                            If "<%=sPagador%>" <> "SW" Then
                                If trim(sBancoBR) = "" or trim(sTipoCuenta) = "" or trim(sCtaCteBR) = "" or trim(sMonedaDeposito) = "" Then
						            msgbox "Debe ingresar todos los datos solicitados para el envío.",,"AFEX"
						            exit sub
						        end if
                            end if
					    End If
					   'FIN INTERNO-3855 MS 26-04-2015
					   
					 ' JFMG 25-05-2012
					 End If
				     ' FIN JFMG 25-05-2012
					    
				End If	
				<% end if %>
			<% end if %>
			HabilitarControles
			
			' JFMG 27-04-2011 mensajeria cliente
			If <%=Session("SolicitudMensajeriaClienteActiva")%> Then
				If "<%=Session("AgentePagadorMensajeriaCliente")%>" = frmGiro.cbxPagador.value Then
					dim sMensajeriaCliente
					sMensajeriaCliente = window.showModalDialog("solicitarmensajeriacliente.asp")
					frmGiro.txtMensajeriaCliente.value = sMensajeriaCliente
					'msgbox frmGiro.txtMensajeriaCliente.value
				End If
			End If
			' FIN JFMG 27-04-2011
	
		<% If Session("ModoPrueba") Then %>
				frmGiro.action = "GrabarEnvioGiro.asp"
				
		<% Else %>
				frmGiro.action = "GrabarEnvioGiro.asp"
		<% End If %>
			frmGiro.submit 
			frmGiro.action = ""
		End If	
		
	End Sub

	Function ValidarDatos()
		
		ValidarDatos = False
		<% If Request.Form("optPersona") = "on" Then %>
			If Trim(frmGiro.txtNombres.value) = "" Then
				MsgBox "Debe ingresar el nombre del remitente",,"AFEX"
				Exit Function
			End If
			If Trim(frmGiro.txtApellidos.value) = "" Then
				MsgBox "Debe ingresar apellidos del remitente",,"AFEX"
				Exit Function
			End If
		<% Else %>
			If Trim(frmGiro.txtRazonSocial.value) = "" Then
				MsgBox "Debe ingresar el nombre del remitente",,"AFEX"
				Exit Function
			End If			
		<% End If %>
		If Trim(frmGiro.txtNombreB.value) = "" Then
			MsgBox "Debe ingresar el nombre del beneficiario",,"AFEX"
			Exit Function
		End If
		If Trim(frmGiro.txtApellidoB.value) = "" Then
			MsgBox "Debe ingresar apellidos del beneficiario",,"AFEX"
			Exit Function
		End If
		If Trim(frmGiro.cbxPaisB.value) = "" Then
			MsgBox "Debe ingresar el pais del beneficiario",,"AFEX"
			Exit Function
		End If
		If Trim(frmGiro.cbxCiudadB.value) = "" Then
			MsgBox "Debe ingresar la ciudad del beneficiario",,"AFEX"
			Exit Function
		End If
		If Trim(frmGiro.txtboleta.value) = "" and "<%=Session("Categoria")%>" =2 or _
		Trim(frmGiro.txtboleta.value) = "" and "<%=Session("Categoria")%>" =1 Then
				MsgBox "Debe ingresar el número de boleta",,"AFEX"
				Exit Function
		End If

		If "<%=sPagador%>" = "<%=session("CodigoMGEnvio")%>"  and "<%=session("Categoria")%>" <> 3 Then
			If cCur(0 & frmGiro.txtTipoCambio.value) < cCur(0 & "<%=nCambioMG%>") Then
					MsgBox "El Tipo CambioMG no puede ser mayor al Tipo de cambio Ingresado. El giro no puede ser enviado ", ,"AFEX"
					Exit Function
			End If
		End If
		If frmGiro.optOtroAgt.value = "on" Then
			
		End If
		If frmGiro.optOtroAgt.value = "on" Then
			If frmGiro.cbxPaisB.value = "CL" Then
				If frmGiro.cbxCiudadB.value = "SCL" Then
				
					If Len(Trim(frmGiro.txtfonoB.value)) < 7  Then
						Msgbox "El número de teléfono para Santiago de Chile debe ser de siete dígitos",, "AFEX"
						Exit Function
					End If
				Else
					If Len(Trim(frmGiro.txtfonoB.value)) < 6 Then
						Msgbox "El número de teléfono para Regiones de Chile debe ser de seis dígitos",, "AFEX"
						Exit Function
					End If			
				End If
			End If
			
		End If
			
		If trim(frmGiro.txtfonob.value)="1234567" or trim(frmGiro.txtfonob.value)="0000000" OR _
				left(trim(frmGiro.txtfonob.value),4)="1111" or trim(frmGiro.txtfonob.value)="1111111" or _ 
				left(trim(frmGiro.txtfonob.value),4)="2222" or trim(frmGiro.txtfonob.value)="2222222" or _
				left(trim(frmGiro.txtfonob.value),4)="3333" or trim(frmGiro.txtfonob.value)="3333333" or _
				left(trim(frmGiro.txtfonob.value),4)="4444" or trim(frmGiro.txtfonob.value)="4444444" or _
				left(trim(frmGiro.txtfonob.value),4)="5555" or trim(frmGiro.txtfonob.value)="5555555" or _
				left(trim(frmGiro.txtfonob.value),4)="6666" or trim(frmGiro.txtfonob.value)="6666666" or _
				left(trim(frmGiro.txtfonob.value),4)="7777" or trim(frmGiro.txtfonob.value)="7777777" or _
				left(trim(frmGiro.txtfonob.value),4)="8888" or trim(frmGiro.txtfonob.value)="8888888" or _
				left(trim(frmGiro.txtfonob.value),4)="9999" or trim(frmGiro.txtfonob.value)="9999999" then
					msgbox "El teléfono ingresado no es válido, Ingrese nuevamente" ,, "AFEX"
					Exit Function 
		End If		
			
		'If frmGiro.optMG.value <> "on" Then
			If Trim(frmGiro.cbxPagador.value) = "" Then
				MsgBox "Debe seleccionar el agente pagador",,"AFEX"
				Exit Function
			End If
		'End If
		If Trim(frmGiro.cbxMonedaPago.value) = "" Then
			MsgBox "Debe seleccionar la moneda de pago",,"AFEX"
			Exit Function
		End If
		If cCur(0 & Trim(frmGiro.txtMonto.value)) = 0 Then
			MsgBox "Debe ingresar el monto del giro",,"AFEX"
			Exit Function
		End If
		If  frmGiro.cbxMonedaGiro.value = "USD" And cCur(0 & Trim(frmGiro.txtMonto.value)) > 3000 and <%=bGiroNacional%> Then	
				 MsgBox "No puede enviar más de USD 3.000",,"AFEX"
				 Exit Function
	    End If
	    
	    If  frmGiro.cbxMonedaGiro.value = "CLP" And cCur(0 & Trim(frmGiro.txtMonto.value)) > 1500000 and <%=bGiroNacional%> Then	
				 MsgBox "No puede enviar más de CLP 1.500.000",,"AFEX"
				 Exit Function
	    End If
		
		<% If Not bGiroNacional Then %>
			If cCur(0 & frmGiro.txtTarifaCobrada.value)  < cCur(0 & "<%=nGastoTransfer%>") Then
				MsgBox "La tarifa cobrada no debe ser menor que " & FormatNumber("<%=nGastoTransfer%>", 2), ,"AFEX"
				Exit Function
			End If
		<% End If %>
		'If frmGiro.optMG.value = "on" Then			
		'	If Len(Trim(frmGiro.txtInvoiceMG.value)) <> 8 Then
		'		MsgBox "El invoice debe ser de 8 dígitos", ,"AFEX"
		'		Exit Function
		'	End If				
		'End If
		'INTERNO-3855 MS 25-04-2015
		If "<%=sPagador%>" = "FP"  and "<%=session("Categoria")%>" <> 3 Then
			If Len(Trim(frmGiro.txtfonoB.value)) < 4  Then
				Msgbox "Debe ingresar un número telefónico válido.",, "AFEX"
				Exit Function
			End If
		End If
		'FIN INTERNO-3855 MS 25-04-2015

		'28-08-2015 INTERNO-4280 MM
        If cCur(frmGiro.txtTarifaCobrada.value) > cCur(frmGiro.txtTarifaSugerida.value) Then
		    MsgBox "Tarifa Cobrada no debe ser superior a la Tarifa Sugerida",,"AFEX"
		    frmGiro.txtTarifaCobrada.select
		    frmGiro.txtTarifaCobrada.focus
            frmGiro.txtTarifaCobrada.style.backgroundColor = "#FF0000"
            Exit Function
        End If  

        'INTERNO-8479 MS 09-11-2016
		If "<%=sPagador%>" = "SW"  Then
			If Len(Trim(frmGiro.txtfonoB.value)) < 6  Then
				Msgbox "Debe ingresar el Teléfono del Beneficiario.",, "AFEX"
				Exit Function
			End If

            
            If Trim(frmGiro.cbxpaisb.value) = "DO"  Then
                If Trim(frmGiro.txtdireccionb.value) = ""  Then
				    Msgbox "Debe ingresar la dirección del beneficiario.",, "AFEX"
				    Exit Function
        End If  
            End If
		End If
		'FIN INTERNO-8479 MS 09-11-2016

		ValidarDatos = True
	End Function
	
    Sub UltimosGiros()
		Dim sString, aGiro, sNombre, sCliente
	
		If Trim(frmGiro.txtExpress.value) = "" Then
			sCliente = Trim(frmGiro.txtExchange.value)
		Else
			sCliente = Trim(frmGiro.txtExpress.value)
		End If
		
		If frmGiro.optPersona.value = "on" Then
			sNombre = Trim(Trim(frmGiro.txtnombres.value) & " " & Trim(frmGiro.txtApellidoP.value) & " " & Trim(frmGiro.txtApellidoM.value))
		Else
			sNombre = Trim(Trim(frmGiro.txtRazonSocial.value))
		End If
		
		sString = Empty
		sString = window.showModalDialog("../Compartido/UltimosGiros.asp?CodigoCliente=" & sCliente & _
																	"&NombreCliente=" & sNombre & _
																	"&CodigoMoneda=" & frmGiro.cbxMonedaPago.value & _
																	"&TipoGiro=<%=afxListaGirosEnviados%>")
		If sString <> Empty Then
			' crea un arreglo con los datos de la transferencia seleccionada
			aGiro = Split(sString, ";", 15)
			
			' asigna los datos al envio
			window.frmGiro.txtNombreB.value = aGiro(0)
			window.frmGiro.txtApellidoB.value = aGiro(1)
			window.frmGiro.txtDireccionB.value = aGiro(2)
			window.frmGiro.txtPaisFonoB.value = aGiro(5)
			window.frmGiro.txtAreaFonoB.value = aGiro(6)
			window.frmGiro.txtFonoB.value = aGiro(7)
			window.frmGiro.txtMonto.value = aGiro(8)
			frmGiro.cbxMonedaPago.value = aGiro(9)
			frmGiro.cbxMonedaGiro.value = aGiro(9)
			window.frmGiro.txtcodigoB.value = aGiro(11)
			window.frmGiro.txtRutB.value = aGiro(12)
			window.frmGiro.txtPasapB.value = aGiro(13)
			window.frmGiro.txtPaisPasB.value = aGiro(14)
			
			HabilitarControles
			frmGiro.cbxCiudadB.value = ""
			frmGiro.cbxPagador.value = ""
			frmGiro.action = "EnviarGiro.asp?Accion=<%=afxAccionMonedaPago%>&PaisB=" & aGiro(3) & "&CiudadB=" & aGiro(4) & _
																		   "&APagador=" & aGiro(9) & "&MonedaPago=" & aGiro(10) & _
																		   "&CodigoB=" & aGiro(11) & "&rutb=" & aGiro(12) & _
																		   "&pasaporteb=" & aGiro(13) & "&Paispasb=" & aGiro(14) & _
																		   "&Beneficiario=" & aGiro(0) 			
			frmGiro.submit 
			frmGiro.action = ""
		End If		
	End Sub
	
	Sub OptDolar_OnClick()
		if frmGiro.OptDolar.value ="on" then
			If frmGiro.txtMonto.value = 0 Then
				frmGiro.txtmonto.value = "<%=FormatNumber(0, 2)%>"
				frmGiro.txttarifacobrada.value = "<%=FormatNumber(0, 2)%>"
				frmGiro.txttarifaSugerida.value = "<%=FormatNumber(0, 2)%>"
				frmGiro.txttotal.value = "<%=FormatNumber(0, 2)%>"
			End If 
			
			frmGiro.OptPesos.checked = false
			frmGiro.txtmonto.disabled= false
			frmGiro.txtmonto.select
			frmGiro.txtpesos.disabled= true
			frmGiro.txtTarifaSugeridapesos.disabled = true 
			frmGiro.OptDolar.checked =true
			frmGiro.txttarifacobradapesos.disabled= false
			frmGiro.txtmonto.value ="<%=FormatNumber(0, 2)%>"
			frmGiro.txtpesos.value=0
			frmGiro.txttarifacobrada.value= "<%=FormatNumber(0, 2)%>"
			frmGiro.txttarifasugerida.value= "<%=FormatNumber(0, 2)%>"
			frmGiro.txttarifacobradapesos.value ="<%=FormatNumber(0, 0)%>"
			frmGiro.txttarifasugeridapesos.value ="<%=FormatNumber(0, 0)%>"
			frmGiro.txttotalpesos.value = "<%=FormatNumber(0, 0)%>"
			frmGiro.txttotal.value = "<%=FormatNumber(0, 2)%>"
		end if
		
	End Sub
	
	Sub OptPesos_onClick()
		If frmGiro.OptDolar.value = "on" Then
			frmGiro.OptDolar.checked =false
			frmGiro.txtpesos.disabled= false
			frmGiro.txtpesos.select
			frmGiro.txtmonto.disabled= true 
			'frmGiro.txttarifacobrada.disabled= true 
			frmGiro.OptPesos.checked = true 
			frmGiro.txtpesos.disabled= false
			'frmGiro.txtTarifacobradapesos.disabled = false
		end If
		
		'frmGiro.txtpesos.value = 0
		'frmgiro.txtmonto.value = 0 
	End Sub
	
	Sub optMG_onClick()
		frmGiro.optOtroAgt.checked = false
		frmGiro.cbxPagador.disabled = true
		frmGiro.txtInvoiceMG.disabled = false		
		
		HabilitarControles
		frmGiro.cbxMonedaPago.value = ""
		frmGiro.action = "EnviarGiro.asp?Accion=<%=afxAccionPagador%>&APagador=<%=Session("CodigoMGEnvio")%>"
		frmGiro.submit 
		frmGiro.action = ""
		
	End Sub
	
	Sub optOtroAgt_onClick()
'		frmGiro.optMG.checked = false
		frmGiro.cbxPagador.disabled = false
		frmGiro.txtInvoiceMG.disabled = true
		frmGiro.txtInvoiceMG.value = ""
		HabilitarControles
		frmGiro.cbxMonedaPago.value = ""
		frmGiro.action = "EnviarGiro.asp?Accion=<%=afxAccionPagador%>"
		frmGiro.submit 
		frmGiro.action = ""		
	End Sub

	sub ListaBeneficiarios()
		dim sBeneficiario
		dim i
		
		sBeneficiario = window.showModalDialog("listabeneficiarios.asp?Cliente=" & frmGiro.txtexpress.value)
		if trim(sBeneficiario) = "" then exit sub
		
		i = instr(sBeneficiario, ";")
		frmGiro.txtnombreb.value = left(sBeneficiario, i - 1)
		sBeneficiario = mid(sBeneficiario, i + 1)
		i = instr(sBeneficiario, ";")
		frmGiro.txtapellidob.value = left(sBeneficiario, i - 1)
		sBeneficiario = mid(sBeneficiario, i + 1)
		i = instr(sBeneficiario, ";")
		frmGiro.cbxpaisb.value = trim(left(sBeneficiario, i - 1))		
		sBeneficiario = mid(sBeneficiario, i + 1)
		i = instr(sBeneficiario, ";")
		frmGiro.cbxciudadb.value = trim(sBeneficiario)	
		frmGiro.txtciudadbeneficiario.value = trim(sBeneficiario)
		
		if frmGiro.cbxciudadb.value <> "" then cbxCiudadB_onBlur()
		if frmGiro.cbxpais.value <> "" then cbxPaisB_onBlur()
	end sub
	
	' JFMG 02-02-2010 se agrega para que no se pueda cambiar el tipo de cambio
	sub txtTipoCambioOP_onchange()
		frmGiro.txtTipoCambioOP.value = "<%=nCambioOp%>"
		msgbox "El tipo de cambio no se puede alterar.", , "AFEX"
	end sub
	' ********** FIN JFMG 02-02-2010 *****************
	
    ' Tecnova rperez 29-08-2016 - Sprint 2

    Function EsLetra(k)
        Dim retorno
        retorno = False
        
        If (k >= 65 And k <= 90) Or k = 192 Or k = 8 Or k = 32 or k = 9 Then 'APPL-41034 MS 24-03-2015
           '     A           Z           Ñ     Backspace  Space    Tab
            retorno = True
        End If

        EsLetra = retorno
    End Function

    Sub txtNombreB_OnKeyDown(e)
        e.returnValue = EsLetra(e.keyCode)
    End Sub

    Sub txtApellidoB_OnKeyDown(e)
        e.returnValue = EsLetra(e.keyCode)
    End Sub


    ' Tecnova rperez 29-09-2016 - TEC-50

    Sub SoloLetras(ByRef formObject)
        Dim regEx
        Set regEx = New RegExp

        regEx.Pattern = "[^\w ñÑ]"
        regEx.Global = True
        
        formObject.value = Trim(formObject.value)
        formObject.value = regEx.Replace(formObject.value, "")
    End Sub

    Sub txtNombreB_OnBlur()
        SoloLetras frmGiro.txtNombreB
    End Sub

    Sub txtApellidoB_OnBlur()
        SoloLetras frmGiro.txtApellidoB
    End Sub

-->
</script>

<body>
    <table><tr><td>
                <marquee STYLE="HEIGHT: 55px; LEFT: 5px; POSITION: absolute; TOP: 1px; WIDTH: 200px" BEHAVIOR="slide" DIRECTION="down" SCROLLAMOUNT="2000" SCROLLDELAY="1">		
	<h1 STYLE="COLOR: #cfcfcf; FONT-SIZE: 25px">Enviar un Giro</h1>
</marquee>
                <marquee STYLE="HEIGHT: 55px; LEFT: 4px; POSITION: absolute; TOP: 0px; WIDTH: 200px" BEHAVIOR="slide" DIRECTION="down" SCROLLAMOUNT="2000" SCROLLDELAY="1">		
	<h1 STYLE="COLOR: steelblue; FONT-SIZE: 25px">Enviar un Giro</h1>
</marquee>
</td></tr></table>
    <!--<marquee STYLE="HEIGHT: 400; LEFT: 4px; POSITION: absolute; TOP: 16px; WIDTH: 573px" BEHAVIOR="slide" DIRECTION="up" SCROLLAMOUNT="2000" SCROLLDELAY="1">-->
    <form id="frmGiro" method="post">
    <input type="hidden" name="txtExchange" value="<%=Request.Form("txtExchange")%>">
    <input type="hidden" name="txtExpress" value="<%=Request.Form("txtExpress")%>">
    <input type="hidden" name="optPersona" value="<%=Request.Form("optPersona")%>">
    <input type="hidden" name="optEmpresa" value="<%=Request.Form("optEmpresa")%>">
    <input type="hidden" name="txtApellidoP" value="<%=Request.Form("txtApellidoP")%>">
    <input type="hidden" name="txtApellidoM" value="<%=Request.Form("txtApellidoM")%>">
    <input type="hidden" name="txtDireccion" value="<%=Request.Form("txtDireccion")%>">
    <input type="hidden" name="cbxComuna" value="<%=Request.Form("cbxComuna")%>">
    <input type="hidden" name="cbxCiudad" value="<%=Request.Form("cbxCiudad")%>">
    <input type="hidden" name="cbxPais" value="<%=Trim(Request.Form("cbxPais"))%>">
    <input type="hidden" name="txtPaisFono" value="<%=Request.Form("txtPaisFono")%>">
    <input type="hidden" name="txtAreaFono" value="<%=Request.Form("txtAreaFono")%>">
    <input type="hidden" name="txtFono" value="<%=Request.Form("txtFono")%>">
    <input type="hidden" name="txtRut" value="<%=Request.Form("txtRut")%>">
    <input type="hidden" name="txtPasaporte" value="<%=Request.Form("txtPasaporte")%>">
    <input type="hidden" name="cbxPaisPasaporte" value="<%=Request.Form("cbxPaisPasaporte")%>">
    <input type="hidden" name="txtGasto" value="<%=nGastoTransfer%>">
    <input type="hidden" name="txtComisionCaptador" value="<%=nComisionCaptador%>">
    <input type="hidden" name="txtComisionPagador" value="<%=nComisionPagador%>">
    <input type="hidden" name="txtComisionMatriz" value="<%=nComisionMatriz%>">
    <input type="hidden" name="txtAfectoIva" value="<%=nAfectoIva%>">
    <!-- APPL-9009 -->
    <input type="hidden" name="cbxSexo" value="<%=Trim(Request.Form("cbxSexo"))%>">
    <input type="hidden" name="txtNumeroCelular" value="<%=Trim(Request.Form("txtNumeroCelular"))%>">
    <!-- FIN APPL-9009 -->
    
    
    <!-- Paso 1 -->
    <table class="borde" id="tabPaso1" cellspacing="0" cellpadding="1" style="position: relative;
        top: 10; left: 6;">
        <tr height="10">
            <td colspan="5" class="Titulo">
                Datos del Remitente
            </td>
        </tr>
        <tr>
            <td colspan="5">
                <table cellspacing="0" cellpadding="0">
                    <tr height="15">
                        <td>
                        </td>
                        <% If bCliente Then	%>
                        <% If Request.Form("optPersona") = "on" Then %>
                        <td>
                            Nombres<br>
                            <input name="txtNombres" size="25" style="height: 22px; width: 300px" disabled value="<%=Request.Form("txtNombres")%>">
                        </td>
                        <td>
                            Apellidos<br>
                            <input name="txtApellidos" size="25" style="height: 22px; width: 200px" disabled
                                value="<%=Trim(Request.Form("txtApellidoP") & " " & Request.Form("txtApellidoM"))%>">
                        </td>
                        <% Else %>
                        <td>
                            Razon Social<br>
                            <input name="txtRazonSocial" size="25" style="height: 22px; width: 300px" disabled
                                value="<%=Request.Form("txtRazonSocial")%>">
                        </td>
                        <% End If %>
                        <% Else %>
                        <td>
                            Nombres<br>
                            <input name="txtNombres" size="25" style="height: 22px; width: 300px" onkeypress="IngresarTexto(2)"
                                onblur="window.frmGiro.txtNombres.value=MayMin(Trim(window.frmGiro.txtNombres.value))"
                                value="<%=Request.Form("txtNombres")%>">
                        </td>
                        <td>
                            Apellidos<br>
                            <input name="txtApellidos" size="25" style="height: 22px; width: 200px" onkeypress="IngresarTexto(2)"
                                onblur="window.frmGiro.txtApellidos.value=MayMin(Trim(window.frmGiro.txtApellidos.value))"
                                value="<%=Request.Form("txtApellidos")%>">
                        </td>
                        <% End If %>
                        <!--			<td COLSPAN="2">Apellidos<br><input NAME="txtApellidoR" SIZE="25" style="HEIGHT: 22px; WIDTH: 200px"></td>			-->
                        <% If Not bExtranjero Then %>
                        <td>
                            <img border="0" id="imgTransferencia" src="../images/Transferencia.jpg" style="cursor: hand"
                                width="19" height="22" onclick="UltimosGiros">
                        </td>
                        <td style="cursor: hand" onclick="UltimosGiros">
                            Ultimos<br>
                            Giros
                        </td>
                        <% End If %>
                    </tr>
                </table>
            </td>
        </tr>
        <tr height="10">
            <td colspan="5" class="Titulo">
                Datos del Beneficiario
            </td>
        </tr>
        <tr>
            <td colspan="5">
                <table cellspacing="0" cellpadding="1">
                    <tr height="15">
                        <td>
                        </td>
                        <td colspan="2">
                            Nombres
                            <!--INTERNO-3855 MS 26-04-2015-->
                            <b>
                                <label id="lblNombresBen" runat="server">
                                    <%= sMsjNombresBeneficiario%>
                                </label>
                            </b>
                            <!--INTERNO-3855 MS 26-04-2015-->
                            <br />
                            <input name="txtNombreB" size="25" style="height: 22px; width: 300px" onkeypress="IngresarTexto(2)"
                                onblurs="window.frmGiro.txtNombreB.value=MayMin(Trim(window.frmGiro.txtNombreB.value))"
                                value="<%=sNombreB%>">
                        </td>
                        <td colspan="2">
                            Apellidos
                            <!--INTERNO-3855 MS 26-04-2015-->
                            <b>
                                <label id="lblApellidosBen" runat="server">
                                    <%= sMsjApellidosBeneficiario%>
                                </label>
                            </b>
                            <!--INTERNO-3855 MS 26-04-2015-->
                            <br />
                            <input name="txtApellidoB" size="25" style="height: 22px; width: 200px" onkeypress="IngresarTexto(2)"
                                onblurs="window.frmGiro.txtApellidoB.value=MayMin(Trim(window.frmGiro.txtApellidoB.value))"
                                value="<%=sApellidoB%>">
                        </td>
                        <% If Not bExtranjero Then %>
                        <td>
                            <img border="0" id="imgTransferencia" src="../images/Transferencia.jpg" style="cursor: hand"
                                width="19" height="22" onclick="ListaBeneficiarios">
                        </td>
                        <td style="cursor: hand" onclick="ListaBeneficiarios">
                            Beneficiarios
                        </td>
                        <% End If %>
                    </tr>
                </table>
            </td>
        </tr>
        <tr>
            <td colspan="5">
                <table border="0" cellspacing="0" cellpadding="1">
                    <tr height="15">
                        <td>
                        </td>
                        <td colspan="2">
                            Dirección<br>
                            <input name="txtDireccionB" size="50" style="height: 22px; width: 280px" onkeypress="IngresarTexto(3)"
                                onblurs="window.frmGiro.txtDireccionB.value=MayMin(Trim(window.frmGiro.txtDireccionB.value))"
                                value="<%=sDireccionB%>">
                        </td>
                        <td>
                            Pais<br>
                            <select name="cbxPaisB" style="width: 120px">
                                <%	
						CargarUbicacion 1, "", sPais 	
                                %>
                            </select>
                        </td>
                        <td colspan="1">
                            Ciudad<br>
                                <select name="cbxCiudadB" style="width: 160px">
                                    <%	
						CargarCiudadesPais sPais, sCiudad 
                                    %>
                                </select>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td colspan="5">
                    <table border="0" cellspacing="0" cellpadding="1">
                        <tr height="15">
                        <td>
                        </td>
                        <td colspan="1">
                            Teléfono
                            <!--INTERNO-3855 MS 26-04-2015-->
                            <b>
                                <label id="lblTelefonoBen" runat="server">
                                    <%=sMsjTelefono%></label></b>
                            <!--INTERNO-3855 MS 26-04-2015-->
                            <br>
                            <input disabled name="txtPaisFonoB" style="width: 40px" value="<%=nDDIPais%>">
                            <input disabled name="txtAreaFonoB" style="width: 40px" value="<%=nDDICiudad%>">
                            <input name="txtFonoB" size="10" maxlength="10" style="width: 90px" onkeypress="IngresarTexto(1)"
                                value="<%=sFonoB%>">
                        </td>
                        <td>
                            Mensaje al Beneficiario<br>
                            <input name="txtMensajeB" style="font-family: verdana; font-size: 9pt; height: 21px;
                                width: 380px" size="255" onblurs="window.frmGiro.txtMensajeB.value=MayMin(Trim(window.frmGiro.txtMensajeB.value))"
                                onkeypress="IngresarTexto(3)" value="<%=Request.Form("txtMensajeB")%>">
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
        <tr height="15">
            <td colspan="5" class="Titulo">
                Datos del Agente Pagador
            </td>
        </tr>
        <tr>
            <td colspan="5">
                <table cellspacing="1" cellpadding="1" id="tbPagador" border="0">
                    <tr height="15" style="display: <%=sDisplay%>">
                        <td colspan="3">
                            Agente Pagador?<br>
                            <%	
				Dim sDisabled 
				If bGiroNacional Then sDisabled = "disabled"
                            %>
                            <input type="radio" name="optOtroAgt" checked style="border: 0" <%=sDisabled%>>
                            <select name="cbxPagador" style="width: 250px" <%=sDisabled%>>
                                <%	If sCiudad <> "" Then
						If sPagador <> Session("CodigoMGEnvio") Then
							CargarAgentePagador sPais, sCiudad, sPagador, 0
						Else
							CargarAgentePagador sPais, sCiudad, "", 0
						End If
					                        End If				
                                        %>
                            </select>
                            <img border="0" id="imgAgtPagador" src="../images/Transferencia.jpg" style="cursor: hand"
                                width="19" height="22" onclick="window.open 'http:ListaAgtPagador.asp?pa=<%=sPais%>&ci=<%=sCiudad%>'"><br>
                        </td>
                            <td width="130" valign="bottom" <% If sPagador = "SW" Then  %> style="display:none;" <% End If %> >
                                    Moneda<br />
                            <% If bGiroNacional   Then %>
                            <select name="cbxMonedaGiro" style="width: 130px; background-color: <%=sColorMoneda%>;
                                color: white; font-weight: bold">
                                <% Else %>
                                <select name="cbxMonedaGiro" style="width: 130px; background-color: <%=sColorMoneda%>;
                                    color: white; font-weight: bold" disabled>
                                    <% End If %>
                                    <%	If sPagador <> "" Then
							CargarMonedaGiro sPagador, sPais, sCiudad, sMoneda
						            End If					
                                %>
                                    </select>
                            </td>
                            
                            <!--INTERNO-8479 MS 11-11-2016--> 
                            <% If sPagador = "SW" Then  %>
                             <td valign="top">
                                Moneda de Pago<br />
                                <% If sPais = "DO" and (sTipo = "1" or sTipo = "2") Then %>
                                    <select name="cbxMonedaPago" style="width: 130px" disabled>
                                <% else %>
                                     <select name="cbxMonedaPago" style="width: 130px"  >
                                <% end if %>
                                
                                <%	
                                    'If bOtrasMonedasPago Then , sPais, sTipo
							            CargarOtrasMonedasPago sMonedaPago
							        'Else
							        '    CargarMonedaGiro sPagador, sPais, sCiudad, sMoneda
							        'End If
						        %>
                                </select>
                        </td>
                            <% End If %>
                            
                            <!--FIN INTERNO-8479 MS 11-11-2016-->
                    </tr>
                    <% If session("Perfil") = 1  and not bGiroNacional  and session("CodigoMGEnvio") <> spagador then %>
                    <tr>
                        <td colspan="2" width="120">
                            <input type="radio" name="OptPesos" checked style="border: 0">Pesos
                        </td>
                        <td width="120">
                            <input type="radio" name="OptDolar" style="border: 0">Dolar
                        </td>
                    </tr>
                    <tr height="10">
                        <td>
                            Monto en Pesos
                            <br>
                            <input style="height: 22px; text-align: right; width: 120px; background-color: DodgerBlue;
                                color: white; font-weight: bold" name="txtpesos" size="15" onkeypress="IngresarTexto(1)"
                                value="<%=FormatNumber(nMontopesos,0)%>">
                        </td>
                        <td>
                            <img style="display: <%=sDisplay%>" border="0" height="22" id="imgCalcular1" name="imgCalcular"
                                onmouseover="imgCalcular1.style.cursor ='Hand'" src="../images/BotonFlechaDerecha.jpg"
                                style="left: 0px; position: relative; top: 5px" width="21">
                        </td>
                        <td style="display: <%=sDisplay%>">
                            Tarifa Sugerida Pesos<br>
                            <input style="height: 22px; text-align: right; width: 120px; background-color: DodgerBlue;
                                color: white; font-weight: bold" name="txtTarifaSugeridaPesos" size="15" value="<%=formatnumber(nTarifaSugeridaPesos,0)%>"
                                sdisabled>
                        </td>
                        <td style="display: <%=sDisplay%>">
                            Tarifa Cobrada Pesos<br>
                            <input style="height: 22px; text-align: right; width: 120px; background-color: DodgerBlue;
                                color: white; font-weight: bold" name="txtTarifaCobradaPesos" size="15" value="<%=formatnumber(ntarifaCobradaPesos,0)%>"
                                sdisabled>
                        </td>
                        <td style="display: <%=sDisplay%>">
                            Total Pesos<br>
                            <input style="height: 22px; text-align: right; width: 120px; background-color: DodgerBlue;
                                color: white; font-weight: bold" name="txtTotalPesos" size="15" value="<%=formatnumber(nTotalPesos,0)%>"
                                sdisabled>
                        </td>
                    </tr>
                    <tr>
                        <td width="121">
                            Tipo Cambio<br>
                            <input style="height: 22px; text-align: right; width: 120px; background-color: #4dc087;
                                color: white; font-weight: bold" name="txtTipoCambioOP" size="15" onkeypress="IngresarTexto(1)"
                                value="<%=FormatNumber(nCambioOp, 2)%>">
                        </td>
                        <td width="121">
                            Monto en Dolar<br>
                            <input style="height: 22px; text-align: right; width: 120px; background-color: #4dc087;
                                color: white; font-weight: bold" name="txtMonto" size="15" onkeypress="IngresarTexto(1)"
                                value="<%=FormatNumber(nMonto, 2)%>">
                        </td>
                        <td>
                            <img style="display: <%=sDisplay%>" border="0" height="22" id="imgCalcular2" name="imgCalcular"
                                onmouseover="imgCalcular2.style.cursor ='Hand'" src="../images/BotonFlechaDerecha.jpg"
                                style="left: 0px; position: relative; top: 5px" width="21">
                        </td>
                        <td style="display: <%=sDisplay%>">
                            Tarifa Sugerida Dolar<br>
                            <input style="height: 22px; text-align: right; width: 120px; background-color: #4dc087;
                                color: white; font-weight: bold" name="txtTarifaSugerida" size="15" value="<%=formatnumber(ntarifasugeridaDolar,2)%>"
                                sdisabled>
                        </td>
                        <td style="display: <%=sDisplay%>">
                            Tarifa Cobrada Dolar<br>
                            <input style="height: 22px; text-align: right; width: 120px; background-color: #4dc087;
                                color: white; font-weight: bold" name="txtTarifaCobrada" size="15" onkeypress="IngresarTexto(1)"
                                value="<%=FormatNumber(nTarifaCobradaDolar, 2)%>">
                        </td>
                        <td style="display: <%=sDisplay%>">
                            Total Dolar<br>
                            <input style="height: 22px; text-align: right; width: 120px; background-color: #4dc087;
                                color: white; font-weight: bold" name="txtTotal" size="15" value="<%=formatnumber(nTotal,2)%>"
                                sdisabled>
                        </td>
                    </tr>
                    <% Else %>
                    <% If (sPagador <> Session("CodigoMGEnvio")) then %>
                    <tr height="10">
                        <td>
                            Forma Pago<br>
                            <% If bDeposita = True  Then %>
                            <select name="cbxdeposito" style="width: 130px; background-color: <%=sColorMoneda%>;
                                color: white; font-weight: bold">
                                <% Elseif bDeposita = false then %>
                                <select name="cbxdeposito" style="width: 130px; background-color: <%=sColorMoneda%>;
                                    color: white; font-weight: bold" disabled>
                                    <% End If %>
                                    <% If sPagador <> "" Then
							    CargarFormaPago sTipo, bPagaDomicilio 'INTERNO-8479 MS 11-11-2016
						End If 	%>
                                </select>
                        </td>
                        <td width="121">
                            Monto<br>
                            <input style="height: 22px; text-align: right; width: 120px; background-color: <%=sColorMoneda%>;
                                color: white; font-weight: bold" name="txtMonto" size="15" onkeypress="IngresarTexto(1)"
                                value="<%=FormatNumber(nMonto, nDec)%>">
                        </td>
                        <td>
                            <img style="display: <%=sDisplay%>" border="0" height="22" id="imgCalcular" name="imgCalcular"
                                onmouseover="imgCalcular.style.cursor ='Hand'" src="../images/BotonFlechaDerecha.jpg"
                                style="left: 0px; position: relative; top: 5px" width="21">
                        </td>
                        <td style="display: <%=sDisplay%>">
                            Tarifa Sugerida<br>
                            <input style="height: 22px; text-align: right; width: 120px; background-color: <%=sColorMoneda%>;
                                color: white; font-weight: bold" name="txtTarifaSugerida" size="15" sdisabled>&nbsp;&nbsp;&nbsp;&nbsp;
                        </td>
                        <td style="display: <%=sDisplay%>">
                            Tarifa Cobrada<br>
                            <input style="height: 22px; text-align: right; width: 120px; background-color: <%=sColorMoneda%>;
                                color: white; font-weight: bold" name="txtTarifaCobrada" size="15" onkeypress="IngresarTexto(1)"
                                value="<%=FormatNumber(nTarifaCobrada, nDec)%>">&nbsp;&nbsp;&nbsp;&nbsp;
                        </td>
                        <td style="display: <%=sDisplay%>">
                            Total<br>
                            <input style="height: 22px; text-align: right; width: 120px; background-color: <%=sColorMoneda%>;
                                color: white; font-weight: bold" name="txtTotal" size="15" sdisabled>
                        </td>
                    </tr>
                    <tr height="10">
                        <td></td>
                        <td></td>
                        <td colspan="3" align="right">* Tarifa Cobrada no debe ser superior a la Tarifa Sugerida </br></td>
                    </tr>
                    <% Else %>
                    <% If (sPagador = Session("CodigoMGEnvio")) or session("Perfil") =1 then %>
                    <tr>
                        <td colspan="5">
                            <table>
                                <tr>
                                    <td>
                                        Monto USD
                                        <br>
                                        <input type="Hidden" name="txtMonto" size="15" onkeypress="IngresarTexto(1)" value="<%=FormatNumber(nMonto, nDec)%>">
                                        <input type="Hidden" name="txtTarifaSugerida" value="<%=FormatNumber(nTarifaSugerida, 2)%>">
                                        <input type="Hidden" name="txtTarifaCobrada" size="15" onkeypress="IngresarTexto(1)"
                                            value="<%=FormatNumber(nTarifaCobrada, 2)%>">
                                        <input type="Hidden" name="txtTotal" size="15" sdisabled>
                                        <input style="height: 22px; text-align: right; width: 120px; background-color: <%=sColorMoneda%>;
                                            color: white; font-weight: bold" name="txtMontoDolar" size="15" onkeypress="IngresarTexto(1)"
                                            value="<%=FormatNumber(nMontoDolar, 2)%>">
                                    </td>
                                    <td>
                                        <img style="display: <%=sDisplay%>" border="0" height="22" id="imgCalcular1" name="imgCalcular1"
                                            onmouseover="imgCalcular1.style.cursor ='Hand'" src="../images/BotonFlechaDerecha.jpg"
                                            style="left: 0px; position: relative; top: 5px" width="21">
                                    </td>
                                    <td style="display: <%=sDisplay%>">
                                        Tipo Cambio<br>
                                        <input style="height: 22px; text-align: right; width: 120px; background-color: <%=sColorMoneda%>;
                                            color: white; font-weight: bold" name="txtTipoCambio" size="15" onkeypress="IngresarTexto(1)"
                                            value="<%=FormatNumber(nTipoCambio, 4)%>">
                                    </td>
                                    <% If session("Categoria")<> 3 Then %>
                                    <td style="display: <%=sDisplay%>">
                                        Tipo Cambio MG<br>
                                        <input style="height: 22px; text-align: right; width: 120px; background-color: <%=sColorMoneda%>;
                                            color: white; font-weight: bold" name="txtCambioMG" size="15" onkeypress="IngresarTexto(1)"
                                            value="<%=FormatNumber(nCambioMG, 4)%>" sdisabled>&nbsp;&nbsp;&nbsp;&nbsp;
                                    </td>
                                    <% End If %>
                                </tr>
                                <tr>
                                    <td colspan="2">
                                        Monto en Pesos
                                        <br>
                                        <input style="height: 22px; text-align: right; width: 120px; color: white; font-weight: bold"
                                            name="txtMontopesos" size="15" onkeypress="IngresarTexto(1)" value="<%=nMontopesos%>"
                                            disabled>
                                    </td>
                                    <td style="display: <%=sDisplay%>">
                                        Tarifa Cobrada<br>
                                        <input style="height: 22px; text-align: right; width: 120px; color: white; font-weight: bold"
                                            name="txtTarifaPesos" size="15" value="<%=ntarifaPesos%>" disabled>&nbsp;&nbsp;&nbsp;&nbsp;
                                    </td>
                                    <td style="display: <%=sDisplay%>">
                                        Total<br>
                                        <input style="height: 22px; text-align: right; width: 120px; color: white; font-weight: bold"
                                            name="txtTotalPesos" size="15" value="<%=nTotalPesos%>" disabled>
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                    <% End If %>
                    <% End If %>
                    <% End If %>
                    
                        <%	If sPagador <> "SW" Then %> <!--INTERNO-8479 MS 11-11-2016-->
                    <!-- JFMG 24-05-2012, ahora los agentes podran tener otra moneda de pago -->
                    <tr id="filaOpcionMonedaPago" style="display: none;">
                        <td>
                            <input type="checkbox" id="chkOtraMonedaPago" name="chkOtraMonedaPago">Otra Moneda de Pago</input>
                        </td>
                    </tr>
                    <tr id="filaMonedaPago" style="display: none;">
                        <td colspan="6">
                            <table width="70%">
                                <tr>
                                    <td valign="top">
                                        Moneda de Pago
                                        <br>
                                        <select name="cbxMonedaPago" style="width: 130px" disabled>
                                            <%	If sPagador <> "" Then
                                                If bOtrasMonedasPago Then
							                        CargarOtrasMonedasPago sMonedaPago
							                    Else
							                        CargarMonedaGiro sPagador, sPais, sCiudad, sMoneda
							                    End If							
						                    End If %>
                                        </select>
                                    </td>
                                    <td valign="top">
                                        Tipo de Cambio
                                        <br>
                                        <input style="height: 22px; text-align: right; width: 120px; background-color: white;
                                            color: black; font-weight: bold" name="txtTipoCambioPagador" size="15" disabled
                                            value="<%=FormatNumber(sTipoCambioMonedaPagador,0)%>">
                                    </td>
                                    <td valign="top">
                                        Monto Moneda Pagador
                                        <br>
                                        <input style="height: 22px; text-align: right; width: 120px; background-color: white;
                                            color: black; font-weight: bold" name="txtMontoPagar" size="15" disabled 
                                            value="<%=FormatNumber(nMontoMonedaPagador,0)%>">
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                    <!-- FIN JFMG 24-05-2012-->
                        <% End If %> <!--INTERNO-8479 MS 11-11-2016-->
                </table>
            </td>
        </tr>
        <tr>
        
            <td colspan="5">
                <table cellspacing="0" cellpadding="1">
                    <tr height="15">
                        <td>
                        </td>
                        <% If bExtranjero Then %>
                        <td>
                            Invoice<br>
                            <input style="height: 22px; text-align: right; width: 107px" id="txtInvoiceMG" name="txtInvoiceMG"
                                size="15" value="<%=Request.Form("txtInvoiceMG")%>">
                        </td>
                        <% Else %>
                        <% If  Session("Categoria") = 3 Then %>
                        <td>
                            Invoice<br>
                            <input style="height: 22px; text-align: right; width: 107px" name="txtInvoiceMG"
                                size="15" value="<%=Request.Form("txtInvoiceMG")%>" onkeypress="IngresarTexto(1)">
                        </td>
                        <% Else %>
                        <td>
                            Invoice<br>
                            <input style="height: 22px; text-align: right; width: 107px" name="txtInvoiceMG"
                                size="15" value="<%=Request.Form("txtInvoiceMG")%>" disabled>
                        </td>
                        <% End If %>
                        <% End If %>
                        <td style="display: <%=sDisplay%>">
                            Nº Boleta<br>
                            <input style="height: 22px; text-align: right; width: 100px" name="txtBoleta" size="15"
                                value="<%=Request.Form("txtBoleta")%>">
                        </td>
                        <td>
                        </td>
                        <td>
                            Mensaje al Agente Pagador<br>
                            <input name="txtMsjPagador" style="font-family: verdana; font-size: 9pt; height: 21px;
                                width: 350px" size="255" onblurs="window.frmGiro.txtMsjPagador.value=MayMin(Trim(window.frmGiro.txtMsjPagador.value))"
                                onkeypress="IngresarTexto(3)" value="<%=Request.Form("txtMsjPagador")%>">
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
        <tr height="0">
            <td>
                <img border="0" id="imgAceptar" onclick src="../images/BotonAceptar.jpg" style="left: 461px;
                    position: relative; top: 0px; cursor: hand" width="70" height="20">
            </td>
        </tr>
    </table>
    <input type="hidden" name="txtGiroBrasil" value="">
    <input type="hidden" name="txtDatosDeposito" value="">
    <input type="hidden" name="txtCiudadBeneficiario" value="">
    <input type="hidden" name="txtRutB" value="<%=sRutBeneficiario%>">
    <input type="hidden" name="txtPasapB" value="<%=sPasaporteBeneficiario%>">
    <input type="hidden" name="txtPaisPasB" value="<%=sPaisPasapBeneficiario%>">
    <input type="hidden" name="txtcodigoB" value="<%=scodigobeneficiario%>">
    <input type="hidden" name="txtNombreCompletoB" value="<%=sNombreCompletoB%>">
    <!-- JFMG 29-04-2011 mensajeria cliente -->
    <input type="hidden" name="txtMensajeriaCliente" value="">
    <!-- FIN JFMG 29-04-2011 -->
    <input type="hidden" name="txtAcumuladoCumplimiento" value="<%=nMontoSumatoriaCumplimiento%>">
    <!--CUM-505 MS 02-02-2016-->
    <input type="hidden" name="txtPasapRemitente" value="<%=sPasaporteRemitente%>">
    <input type="hidden" name="txtPaisPasaporteRemitente" value="<%=sPaisPasaporteRemitente%>">
    <input type="hidden" name="txtRutRemitente" value="<%=sRutR%>">
    <input type="hidden" name="txtAutorizadoOperarPasaporte" value="<%=nAutorizacionOperarPasaporte%>">
    <input type="hidden" name="txtAccionCumplimiento" value="">
    <input type="hidden" name="txtPasaporteIngresadoYAprobado" value="<%=nPasaporteIngresadoYAprobado %>">
     <!--CUM-505 MS 02-02-2016-->
    </form>
</body>
<!--#INCLUDE virtual="/Compartido/Rutinas.htm" -->
</html>

<!-- JFMG 24-05-2012-->
<%
    SET rsOtrasMonedasPago = nothing    
%>
<!-- FIN JFMG 24-05-2012-->
