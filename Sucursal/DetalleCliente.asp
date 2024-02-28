<%@ LANGUAGE = VBScript %>
<%
	Response.Expires = 0
	
	' JFMG 004-07-2012 se agrega para el ingreso de usuarios desde otras plataformas
	Session("TipoOrigenLlamada") = 0
	IF request("TipoOrigenLlamada") = "2" THEN
		' se llama desde el menú de una sucursal
		Session("CodigoCliente") = request("SCC")
		Session("CodigoAgente") = request("AGC")
		Session("TipoOrigenLlamada") = 2 ' página envio informacion optima
		Session("NombreUsuarioOperador") = request("NUO")
		Session("VerClienteCorporativo") = True
	END IF
	' FIN JFMG 04-07-2012
	
	
	'If Session("CodigoCliente") = "" Then
	'	response.Redirect "../Compartido/TimeOut.htm"
	'	response.end
	'End If
%>
<!--#INCLUDE virtual="/Compartido/Rutinas.asp" -->
<!--#INCLUDE virtual="/Compartido/Errores.asp" -->
<!--#INCLUDE virtual="/sucursal/Rutinas.asp" -->
<!--#INCLUDE virtual="/Compartido/RutinasEncriptar.asp" -->
<%
	Dim rsHistoria, rsDocumento, rsCliente, rsActividadEconomica
	Dim bCliente, bDocumento, bHistoria, nMenu, sRazonSocial
	Dim sRut, sPasaporte, sNombrePaisPasaporte, sNombreCompleto, sNombres, sApellidoP, sApellidoM
	Dim sDireccion, sPais, sCiudad, sComuna, nTipoCliente,sRepresentante
	Dim sPaisFono, sAreaFono, sNumeroFono, sPaisFax, sAreaFax, sNumeroFax, sCelular, sUsuario, sPassword
	Dim sBanco, sSucursal, sCtaCte, sCtaAhorro, nCredito, nCreditoUsado, nDiasRetencion, sPregunta, sRespuesta
	Dim nAccion		'0 Cargar Detalle; 1 Cargar Pais,ciudad,comuna
	Dim nModo		'0 Detalle; 1 Modificar
	Dim sHabilitado
	Dim sChkInfCom, sChkInfTri, sChkPon, sChkAcrCom, sChkCarCli
	Dim nChkInfCom, nChkInfTri, nChkPon, nChkAcrCom, nChkCarCli
	Dim nEjecutivo, sNacionalidad, sCorreo, nRubro, nCodigoCli
	Dim sFechaCreacion
	Dim sEstado
	dim iRiesgo
	Dim iPerfilPEP
	Dim iPerfilZona
	Dim iPerfilRS
	Dim iPerfilACT
	Dim iPerfilIndustria
	dim sFechaNacimiento, sContacto1, sPorcentageContacto1, sContacto2, sPorcentageContacto2, sFechaActivacionComision
	dim bIDConsultada, sNumeroSerieID
	dim sSQL
	dim rs
	dim sMensajeRegistro, bIDValida
	dim sNumero, sDepto, sProfesion, sSexo ' agregado PSS 28-04-2009 para cumplimiento legal
	Dim nClienteAgencia
    dim iPerfilCliente 
    Dim sOcupacion 'INTERNO-9263 MS 19-01-2017
	
	' JFMG 20-06-2011 datos solicitados por OPTIMA
    Dim sPropositoTransacciones
	Dim sOrigenFondos
	Dim sReferenciaBancaria
	Dim sIngresoAnual
	Dim sCantidadTransaccionesMesCLIENTE
	Dim sCantidadTransaccionesMesAFEX
	Dim sMontoTransaccionesMesCLIENTE
	Dim sMontoTransaccionesMesAFEX
	Dim sActividadEconomica
	Dim sActividadEconomicaInactiva
	Dim sMontoPatrimonio
	Dim sNombreEmpleador
    ' FIN 20-06-2011
    dim sGiroComercial 'JBV17-05-2012
    Dim sMotivo 'MS 20140318
	nAccion = cInt(0 & Request("acc"))
	nMenu = cInt(0 & Request("mnu"))
	nModo = cInt(0 & Request("md"))
	nCodigoCliente = Request("cc")
	nClienteAgencia = request("CA")
	
	' JFMG 04-07-2012 se validara si se muestra o no el menu para permitir ingreso desde otras plataformas
	Dim bMostrarMenu
	bMostrarMenu = True
	If Session("NombreUsuarioOperador") = "" Then
	    bMostrarMenu = False
	End If
	' FIN JFMG 04-07-2012
	
	If nCLienteAgencia ="True" Then
		nCLienteAgencia= "Checked"
	else
		nClienteAgencia= ""
	End If
	'Response.Write  nclienteagencia
	' verifica si se valido la identificación
	if request("Accion") = 6 then
		dim iIDValida
		
		ValidarRegistro
		if bIDConsultada = "true" then
			if bIDValida = "true" then
				iIDValida = 1
				sSQL = "exec validarclienterc " & nCodigoCliente & ", " & evaluarstr(request("NumeroSerie")) & ", " & _
					iIDValida & ", " & evaluarstr(sMensajeRegistro) & ", " & evaluarstr(Session("NombreUsuarioOperador"))
		
				
				Set rs = EjecutarSQLCliente(Session("afxCnxCorporativa"), sSQL)
				if err.number <> 0 then
					response.Redirect "http:../compartido/Error.asp?Titulo=Error en ValidarClienteRC 0&Number=" & Err.Number  & "&Source=" & Err.Source & "&Description=" & sMensajeRegistro & " - " & Err.Description  
				else
					Response.Redirect "DetalleCliente.asp?cc=" & nCodigoCliente
				end if
			else
				iIDValida = 0
			end if
		end if
	end if	
	
	elimina_documento
	habilita_documento
	
	'*******************************************
	'MS 10-03-2014
	autoriza_documento
	rechaza_documento
	'FIN MS 10-03-2014
	'*******************************************
	
	If nModo = 1 Then
		sHabilitado = ""
	Else
		sHabilitado = "disabled"
	End If
	
	Select Case nAccion
	Case 0	'Cargar Cliente
		CargarCliente
		
	Case Else	'Cargar Otros
		bCliente = True
		CargarClienteForm		
	End Select
	
	If Not bCliente Then
		Response.Redirect "http:../Compartido/Error.asp?Titulo=Detalle Cliente&Description=No se encontró el cliente " & Request("cc")
		Response.End 
	End If
    'MS 27-03-2014
    dim nHistoriaCompleta
    if request("hist")="1" then
        nHistoriaCompleta = 1
    else
        nHistoriaCompleta = 0
    end if
    'FIN MS 27-03-2014
    
	Set rsHistoria = ObtenerHistoria(nHistoriaCompleta)
	Set rsDocumento = ObtenerDocumento()
	 
	Set rsActividadEconomica = ObtenerActividadEconomica()
	
	'MS 28-03-2014
	Dim sActividadSi, sActividadNo
	if rsActividadEconomica is nothing then 
    ' Do Until rsActividadEconomica.EOF
    '    If rsActividadEconomica("activa") Then
    '      sActividadSi = " Checked"
    '      sActividadNo = ""
    '     exit Do
    '   End If
    '    rsActividadEconomica.MoveNext							                    
    ' Loop
	else
	    sActividadSi = ""
	    sActividadNo = " Checked"
	end if
	'FIN MS 28-03-2014
	
'************* Funciones y Procedimientos *******************
	sub elimina_documento()
		If request("elim") = 1 then
		    'MS 18-03-2014
            if trim(request("motivo"))<>"" then
			    sMotivo = EvaluarStr(trim(request("motivo")))
		    sSQL = "EXEC ActualizarEstadoDocumentoCliente " & request("cc") & ", " & request("doc") & ", " & request("tip") & ", 4," & sMotivo 'INTERNO-2850
			else
			    sMotivo = "NULL"
		    sSQL = "EXEC ActualizarEstadoDocumentoCliente " & request("cc") & ", " & request("doc") & ", " & request("tip") & ", 4," & sMotivo 'INTERNO-2850
sMotivo = ""
			end if
			'FIN MS 18-03-2014
		    
	'	    sSQL = "EXEC ActualizarEstadoDocumentoCliente " & request("cc") & ", " & request("doc") & ", " & request("tip") & ", 4," & sMotivo 'INTERNO-2850
			Set rs = EjecutarSQLCliente(Session("afxCnxCorporativa"), sSQL)
			Set rs = Nothing
			sqldoc="select nombre from tipo_documento where codigo="& request("tip") &""
			Set rs2 = EjecutarSQLCliente(Session("afxCnxCorporativa"), sqldoc)
			if not rs2.eof then 
				nom_documento="Deshabilitación de " & rs2("nombre") & " desde eAfex. " & sMotivo 
			end if
			Set rs2 = Nothing

			AgregarHistoria nCodigoCliente, nom_documento, 1, 0			
		end if	
	end sub
	
	'MS 21-03-2014
	sub habilita_documento()
		If request("elim") = 2 then
		    
            if trim(request("motivo"))<>"" then
			    sMotivo = trim(request("motivo"))
			else
			    sMotivo = ""
			end if
			
			sSQL = "EXEC ActualizarEstadoDocumentoCliente " & request("cc") & ", " & request("doc") & ", " & request("tip") & ", 3," & sMotivo 'INTERNO-2850
			Set rs = EjecutarSQLCliente(Session("afxCnxCorporativa"), sSQL)
			Set rs = Nothing
			
			sqldoc="select nombre from tipo_documento where codigo="& request("tip") &""
			Set rs2 = EjecutarSQLCliente(Session("afxCnxCorporativa"), sqldoc)
			if not rs2.eof then 
				nom_documento="Habilitación de "& rs2("nombre") & "desde eAfex. " & sMotivo
			end if
			Set rs2 = Nothing
			AgregarHistoria nCodigoCliente, nom_documento, 1, 0			
		end if	
	end sub
	'FIN MS 21-03-2014
	
	'*******************************************
	'MS 10-03-2014
	sub autoriza_documento()
		If request("aut") = 1 then
		    sSQL = "EXEC ActualizarEstadoDocumentoCliente " & request("cc") & ", " & request("doc") & ", " & request("tip") & ", 1,'' " 'INTERNO-2850
			Set rs = EjecutarSQLCliente(Session("afxCnxCorporativa"), sSQL)
			Set rs = Nothing
			
			sqldoc="SELECT nombre FROM tipo_documento WHERE codigo="& request("tip") &""
			Set rs2 = EjecutarSQLCliente(Session("afxCnxCorporativa"), sqldoc)
			if not rs2.eof then 
				nom_documento="Autorización de " & rs2("nombre")  & " desde eAfex."
			end if
			Set rs2 = Nothing
			
			AgregarHistoria nCodigoCliente, nom_documento, 1, 0			
		end if	
	end sub
	
	sub rechaza_documento()
		If request("aut") = 2 then
		     'MS 18-03-2014
		    if trim(request("motivo"))<>"" then
			    sMotivo = trim(request("motivo"))
			else
			    sMotivo = ""
			end if
			'FIN MS 18-03-2014
			sSQL = "EXEC ActualizarEstadoDocumentoCliente " & request("cc") & ", " & request("doc") & ", " & request("tip") & ", 2,'' " 'INTERNO-2850
			Set rs = EjecutarSQLCliente(Session("afxCnxCorporativa"), sSQL)
			Set rs = Nothing
			
			sqldoc="SELECT nombre FROM tipo_documento WHERE codigo=" & request("tip") &""
			Set rs2 = EjecutarSQLCliente(Session("afxCnxCorporativa"), sqldoc)
			
			if not rs2.eof then 
				nom_documento="Documento " & rs2("nombre") & " rechazado desde eAfex. " & sMotivo
			end if
			Set rs2 = Nothing
			'Response.Write nom_documento
			AgregarHistoria nCodigoCliente, nom_documento, 1, 0			
		end if	
	end sub
	'FIN MS 10-03-2014
	'*******************************************
		
	Sub CargarCliente()
		Set rsCliente = ObtenerCliente()
		If rsCliente Is Nothing Then
			bCliente = False
		ElseIf rsCliente.EOF Then
			bCliente = False
		Else
			bCliente = True
		End If
		
		If Not bCliente Then Exit Sub
		
		sRut = EvaluarVar(rsCliente("rut"), "")
		sPasaporte= EvaluarVar(rsCliente("pasaporte"), "")
		sNombrePaisPasaporte = EvaluarVar(rsCliente("nombrepaispasaporte"), "")
		sEstado= EvaluarVar(rsCliente("estado"), "")	
		Session("sEstadoOriginal") = EvaluarVar(rsCliente("estado"), "")		
		sNombres = EvaluarVar(rsCliente("Nombres"), "")
    	sNombreCompleto = EvaluarVar(rsCliente("Nombre"), "")	
		sApellidoP = EvaluarVar(rsCliente("apellido_paterno"), "")
		sApellidoM = EvaluarVar(rsCliente("apellido_materno"), "")
		sRazonSocial = EvaluarVar(rsCliente("razon_social"), "")
		sRepresentante= EvaluarVar(rsCliente("nombre_representante"), "") 
		sPais = EvaluarVar(rsCliente("pais_particular"), "")
		sCiudad = EvaluarVar(rsCliente("ciudad_particular"), "")
		sComuna = EvaluarVar(rsCliente("comuna_particular"), "")
		IF ISNULL(rsCliente("calle_particular")) THEN
			sDireccion = trim(Evaluarvar(rsCliente("direccion_particular"),""))
		ELSE
			sDireccion = trim(EvaluarVar(rsCliente("calle_particular"), ""))
		END IF
		sNumero = EvaluarVar(rsCliente("numero_particular"),"")
		sDepto = Evaluarvar(rsCliente("departamento_particular"),"")
		sProfesion = Evaluarvar(rsCliente("Profesion"),"")
        sOcupacion = Evaluarvar(rsCliente("IdOcupacion"),"0") 'INTERNO-9263	MS 19-01-2017
		sSexo = evaluarvar(rsCliente("Sexo"),"")
		If sSexo = empty then
			sSexo= 0
		else
			sSexo=  rsCliente("Sexo")
		end if
		sPaisFono = EvaluarVar(rsCliente("pais_telefono_particular"), "")
		sAreaFono = EvaluarVar(rsCliente("area_telefono_particular"), "")
		sNumeroFono = EvaluarVar(rsCliente("numero_telefono_particular"), "")
		sPaisFax = EvaluarVar(rsCliente("pais_fax_particular"), "")
		sAreaFax = EvaluarVar(rsCliente("area_fax_particular"), "")
		sNumeroFax = EvaluarVar(rsCliente("numero_fax_particular"), "")
		sCelular = EvaluarVar(rsCliente("numero_celular"), "")
		sBanco = EvaluarVar(rsCliente("banco"), 0)
		sSucursal = EvaluarVar(rsCliente("sucursal_origen"), "")
		sCtaCte = EvaluarVar(rsCliente("cuenta_corriente"), "")
		sCtaAhorro = EvaluarVar(rsCliente("cuenta_ahorro"),"")
		nTipoCliente = EvaluarVar(rsCliente("tipo"), 0)
		nCredito = cCur(0 & EvaluarVar(rsCliente("credito"), 0))
		nDiasRetencion = cInt(0 & EvaluarVar(rsCliente("dias_retencion"), 0))
		sUsuario = EvaluarVar(rsCliente("nombre_usuario"), "")
		sPassword = EvaluarVar(rsCliente("password"), "")
		sPregunta = EvaluarVar(rsCliente("pregunta"), "")
		sRespuesta = EvaluarVar(rsCliente("respuesta"), "")
		iRiesgo = rsCliente("nivelriesgo")
		iPerfilPEP = rsCliente("ppep")
		iPerfilZona = rsCliente("pzona")
		iPerfilRS = rsCliente("presidencia")
		iPerfilACT = rsCliente("pactividad")
		iPerfilIndustria = rsCliente("pindustria")
		sFechaNacimiento = rsCliente("Fecha_Nacimiento")		
		sContacto1 = rsCliente("empleadoContacto")
		sPorcentagecontacto1 = rsCliente("porcentajecontacto")
		sContacto2 = rsCliente("empleadoContacto2")
		sPorcentagecontacto2 = rsCliente("porcentajecontacto2")
		sFechaActivacionComision = rsCliente("fechaactivacioncomision")		
		iPerfilCliente = rsCliente("Perfil_Cliente")
		bIDConsultada = rsCliente("identificacionconsultada")
		if bIDConsultada = 1 then 
			bIDConsultada = true
		else
			bIDConsultada = false
		end if
		sNumeroSerieID = rsCliente("serieidentificacion")

		if trim(iRiesgo) = empty or isnull(iRiesgo) then iRiesgo = 1
		
		nCodigoCli = EvaluarVar(rsCliente("codigo"), 0)
		
		nCreditoUsado = cCur(0 & CreditoUsado(nCodigoCli, nDiasRetencion))

		nChkInfCom = cInt(0 & EvaluarVar(rsCliente("ctf_informacion_comercial"), 0))
		nChkInfTri = cInt(0 & EvaluarVar(rsCliente("ctf_informacion_tributaria"), 0))
		nChkPon = cInt(0 & EvaluarVar(rsCliente("ctf_ponderacion"), 0))
		nChkAcrCom = cInt(0 & EvaluarVar(rsCliente("ctf_acreditacion_comercial"), 0))
		nChkCarCli = cInt(0 & EvaluarVar(rsCliente("ctf_carpeta_cliente"), 0))
		
		nClienteAgencia = cint(0 & rsCliente("Operaagencia"))
		
		If nClienteAgencia = 1 then nClienteAgencia = "Checked"
		If nChkInfCom = 1 Then sChkInfCom = "checked"
		If nChkInfTri = 1 Then sChkInfTri = "checked"
		If nChkPon = 1 Then sChkPon = "checked"
		If nChkAcrCom = 1 Then sChkAcrCom = "checked"
		If nChkCarCli = 1 Then sChkCarCli = "checked"
		
		nEjecutivo = cInt(0 & EvaluarVar(rsCliente("codigo_ejecutivo"), 0))		
		sNacionalidad = EvaluarVar(rsCliente("nacionalidad"), "")
		sCorreo = EvaluarVar(rsCliente("email"), "")
		nRubro = EvaluarVar(rsCliente("codigo_rubro"), 0)
		If IsNull(rsCliente("fecha_creacion")) Then
			sFechaCreacion = Empty
		Else
			sFechaCreacion = FormatDateTime(rsCliente("fecha_creacion"), 2)
		End If
		
		' JFMG 20-06-2011 nuevos datos, solicitados por OPTIMA
		sPropositoTransacciones = EvaluarVar(rsCliente("PropositoTransacciones"), "")
		sOrigenFondos = EvaluarVar(rsCliente("OrigenPatrimonialFondos"), "")
		sReferenciaBancaria = EvaluarVar(rsCliente("ReferenciaComercialBancaria"), "")
        sIngresoAnual = EvaluarVar(rsCliente("MontoIngresoAnualUSD"), "")
		sMontoPatrimonio = EvaluarVar(rsCliente("MontoPatrimonioUSD"), "")
		sCantidadTransaccionesMesCLIENTE = EvaluarVar(rsCliente("CantidadTransaccionesMesCLIENTE"), "")
		sCantidadTransaccionesMesAFEX = EvaluarVar(rsCliente("CantidadTransaccionesMesAFEX"), "")
		sMontoTransaccionesMesCLIENTE = EvaluarVar(rsCliente("MontoTransaccionesUSDMesCLIENTE"), "")
		sMontoTransaccionesMesAFEX = EvaluarVar(rsCliente("MontoTransaccionesUSDMesAFEX"), "")
		sNombreEmpleador = EvaluarVar(rsCliente("NombreEmpleador"), "")
		' FIN JFMG 20-06-2011
		sGiroComercial = EvaluarVar(rsCliente ("Giro"),"") 'JBV17-05-2012
	End Sub

	Sub CargarClienteForm()
		sPais = Request.Form("cbxPaisPersonal")
		sCiudad = Request.Form("cbxCiudadPersonal")
		sComuna = Request.Form("cbxComunaPersonal")
		sBanco = Request.Form("cbxBanco")						
		sRut = Request.Form("txtRut")
		sPasaporte = Request.Form("txtPasaporte")
		sNombrePaisPasaporte = Request.Form("txtNombrePaisPasaporte")
		sNombres = Request.Form("txtNombres")
		sNombreCompleto = Request.Form("txtNombreCompleto")
		sApellidoP = Request.Form("txtApellidoP")
		sApellidoM = Request.Form("txtApellidoM")
		sRazonSocial = Request.Form("txtRazonSocial")
		sDireccion = Request.Form("txtDireccionPersonal")
		sNumero = Request.Form("txtNumero")
		sDepto = Request.Form("txtDepto")
		sProfesion = Request.Form("txtProfesion")
        sOcupacion = Request.Form("cbxOcupacion") 'INTERNO-9263	MS 19-01-2017
		sSexo = Request.Form("cbxSexo")
		sPaisFono = Obtenerddi(1, sPais) 'Request.Form("txtPaisFonoPersonal")
		sAreaFono = Obtenerddi(2, sCiudad) 'Request.Form("txtAreaFonoPersonal")
		sNumeroFono = Request.Form("txtFonoPersonal")
		sPaisFax = Obtenerddi(1, sPais) 'Request.Form("txtPaisFaxPersonal")
		sAreaFax = Obtenerddi(2, sCiudad) 'Request.Form("txtAreaFaxPersonal")
		sNumeroFax = Request.Form("txtFaxPersonal")
		sCelular = Request.Form("txtCelular")
		sSucursal = Request.Form("cbxSucursal")
		sCtaCte = Request.Form("txtCuentaCorriente")
		sCtaAhorro = Request.Form("txtCuentaAhorro")
		nTipoCliente = Request.Form("cbxTipo") 'Request.Form("txtTipoCliente")
		nCredito = Request.Form("txtCredito")
		nCreditoUsado = Request.Form("txtCreditoUsado")
		sUsuario = Request.Form("txtUsuario")
		sPassword = Request.Form("txtPassword")
		sPregunta = Request.Form("txtPregunta")
		sRespuesta = Request.Form("txtRespuesta")
		iRiesgo = Request.Form("cbxRiesgo")
		sFechaNacimiento = Request.Form("txtFechaNacimiento")
		sContacto1 = Request.Form("txtContacto1")
		sPorcentagecontacto1 = Request.Form("txtporcentagecontacto1")
		sContacto2 = Request.Form("txtContacto2")
		sPorcentagecontacto2 = Request.Form("txtporcentagecontacto2")
		sFechaActivacionComision = Request.Form("txtFechaActivacionComision")		
				
		if request.form("opthabilitar") = "on" then 
			sEstado=1
		else 
			sEstado=0
		end if
		
		' JFMG 14-11-2008
		sNumeroSerieID = Request.Form("txtNumeroSerieRut")
		' ************************ FIN **********************************
		
		'response.write request.form("opthabilitar")
		'response.end
		
		If Request.Form("chkInformacionComercial") = "on" Then nChkInfCom = 1 Else nChkInfCom = 0
		If Request.Form("chkInformacionTributaria") = "on" Then nChkInfTri = 1 Else nChkInfTri = 0
		If Request.Form("chkPonderacion") = "on" Then nChkPon = 1 Else nChkPon = 0
		If Request.Form("chkAcreditacionComercial") = "on" Then nChkAcrCom = 1 Else nChkAcrCom = 0
		If Request.Form("chkCarpetaCliente") = "on" Then nChkCarCli = 1 Else nChkCarCli = 0

		If nChkInfCom = 1 Then sChkInfCom = "checked"
		If nChkInfTri = 1 Then sChkInfTri = "checked"
		If nChkPon = 1 Then sChkPon = "checked"
		If nChkAcrCom = 1 Then sChkAcrCom = "checked"
		If nChkCarCli = 1 Then sChkCarCli = "checked"

		nEjecutivo = Request.Form("cbxEjecutivos")
		sNacionalidad = Request.Form("cbxNacionalidad")
		sCorreo = Request.Form("txtCorreo")
		nRubro = Request.Form("cbxRubro")
		sSexo = Request.Form("cbxSexo")
		
		' JFMG 20-06-2011 nuevos datos, solicitados por OPTIMA
		sPropositoTransacciones = Request.Form("txtPropositoTransacciones")
		sOrigenFondos = Request.Form("txtOrigenFondos")
		sReferenciaBancaria = Request.Form("txtReferenciaBancaria")
        sMontoPatrimonio = Request.Form("txtMontoPatrimonio")
		sIngresoAnual = Request.Form("txtIngresoAnual")
		sCantidadTransaccionesMesCLIENTE = Request.Form("txtCantidadTransaccionesMesCLIENTE")
		sCantidadTransaccionesMesAFEX = Request.Form("txtCantidadTransaccionesMesAFEX")
		sMontoTransaccionesMesCLIENTE = Request.Form("txtMontoTransaccionesMesCLIENTE")
		sMontoTransaccionesMesAFEX = Request.Form("txtMontoTransaccionesMesAFEX")
		sActividadEconomica = request.form("txtActividadEconomica")
		sActividadEconomicaInactiva = request.form("txtActividadEconomicaInactiva")
		sNombreEmpleador = request.form("txtNombreEmpleador")
		' FIN JFMG 20-06-2011
		sGiroComercial = request.Form ("txtGiroComercial")'JBV17-05-2012
	End Sub
	
	Function ObtenerHistoria(HistoriaCompleta)
	   Dim rsATC
	   Dim sSQL
	   Dim Completa ' MS 28-03-2014

	   Set ObtenerHistoria = Nothing
       'INTERNO-2850
	   On Error Resume Next
       sSQL = "EXEC ObtenerHistoriaCliente "  & nCodigoCliente & ", " & HistoriaCompleta
	   'FIN INTERNO-2850
	   Set rsATC = EjecutarSQLCliente(Session("afxCnxCorporativa"), sSQL)

	   If Err.Number <> 0 Then 
			Set rsATC = Nothing
			MostrarErrorMS "Obtener Historia 1"
		End If
	   
	   Set rsATC.ActiveConnection = Nothing
	   Set ObtenerHistoria = rsATC

	   Set rsATC = Nothing	
	End Function

	Function ObtenerDocumento()
	   Dim rsATC
	   Dim sSQL

	   Set ObtenerDocumento = Nothing

	   On Error Resume Next
   	   
	   sSQL ="exec BuscarArchivosCliente "  & nCodigoCliente
	   Set rsATC = EjecutarSQLCliente(Session("afxCnxCorporativa"), sSQL)
	   
	   If Err.Number <> 0 Then 
	   		Set rsATC = Nothing
			MostrarErrorMS "Obtener Documento 1"
	   End If
	   
	   
	   Set rsATC.ActiveConnection = Nothing
	   Set ObtenerDocumento = rsATC
	   
	   Set rsATC = Nothing
	End Function

	Function ObtenerCliente()
	   Dim rsATC
	   Dim sSQL

	   Set ObtenerCliente = Nothing

	   On Error Resume Next
        'INTERNO-2850
		' Jonathan Miranda G. 10-04-2007
		' verifica si lo busca por rut, codigo o pasaporte
		if Request("cc") <> empty then
			sSQL = "EXEC ObtenerClienteATC " &  Request("cc") & ", null, null"
		elseif Request("rt") <> empty then
		    sSQL = "EXEC ObtenerClienteATC null, " & evaluarstr(Request("rt")) & ", null "
		elseif Request("ps") <> empty then
			sSQL = "EXEC ObtenerClienteATC null, null, " & evaluarstr(Request("ps")) 
		' ******** FIN JFMG 18-01-2010 *******************
		
		end if		
		
		'FIN INTERNO-2850	
		'----------------------- Fin ------------------------------------
	   Set rsATC = EjecutarSQLCliente(Session("afxCnxCorporativa"), sSQL)

	   If Err.Number <> 0 Then 
			Set rsATC = Nothing
			MostrarErrorMS "Obtener Cliente 1" 
		End If
	
		If rsATC.eof Then 
			Set rsATC = Nothing
			MostrarErrorMS "No se encontró el Cliente"
		End If
		nCodigoCliente = rsATC("codigo")
	   
	   Set rsATC.ActiveConnection = Nothing
	   Set ObtenerCliente = rsATC

	   Set rsATC = Nothing
	End Function

    Function ObtenerActividadEconomica()
	    Dim rs
	    Dim sSQL

	    Set ObtenerActividadEconomica = Nothing

	    On Error Resume Next
        sSQL = " exec MostrarListaActividadEconomicaCliente " & request("cc")
          
	   	Set rs = EjecutarSQLCliente(Session("afxCnxCorporativa"), sSQL)   
		
	    If Err.Number <> 0 Then 
	   	    Set rs = Nothing
			MostrarErrorMS "Obtener Actividad Económica 1"
		End If
    	   
	    Set rs.ActiveConnection = Nothing
	    Set ObtenerActividadEconomica = rs

	    Set rs = Nothing
	End Function

	Sub ValidarRegistro()
		Dim sRut, sDigito, sPasaporte, sRespuestaRegistro, sNumeroRut, sTipo		
				
		sTipo = request("Tipo")
		sNumeroRut = Request("NumeroRut")		
		
		bIDValida = ConsularIdentificacion(sTipo, sNumeroRut, trim(Request("Numeroserie")))
		bConsulta = split(bIDValida,";",3)
		
		bIDConsultada = bConsulta(1)
		bIDValida = bConsulta(0)
		sMensajeRegistro = bConsulta(2)		
		
	End Sub
	
	Function ConsularIdentificacion(byval Tipo, byval NumeroID, byval NumeroSerie)
		dim sXML 
		dim i
		dim sDigitoVerificador, sMensaje
		
		on error resume next		
		
		sXML = "<?xml version='1.0' encoding='utf-8'?>" & _
				"<soapenv:Envelope xmlns:soapenv='http://schemas.xmlsoap.org/soap/envelope/'>" & _
				"<soapenv:Body>" & _
				"<VerificarId xmlns='http://tempuri.org/'>" & _
					"<TipoDocumento>" & Tipo & "</TipoDocumento>" & _
					"<NumeroDocumento>" & NumeroID & "</NumeroDocumento>" & _
					"<NumeroSerie>" & NumeroSerie & "</NumeroSerie>" & _
				"</VerificarId>" & _
			"</soapenv:Body>" & _
			"</soapenv:Envelope>"		
				
		ConsularIdentificacion = "false;false;"
		
		Set WebServices = CreateObject("msxml2.serverxmlhttp")
		Set myXML = CreateObject("MSXML2.DOMDocument")
		Set XMLEnviar = CreateObject("MSXML2.DOMDocument")
		
		XMLEnviar.loadXML(sXML)
			
		myXML.Async = False		
		WebURL = "http://peumo/registrocivil/service.asmx"
		WebServices.Open "POST",WebURL , false
		
		
		WebServices.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
		WebServices.setRequestHeader "Content-Length", "length"
		WebServices.setRequestHeader "SOAPAction", "http://tempuri.org/VerificarId"
		webservices.send(xmlenviar)
				
		if WebServices.readyState <> 4 then
			smensajeregistro = "Transferencia Incompleta ." & webservices.responseText & err.Description 
			
		else		
			if WebServices.status = 200 then ' Respuesta del Servidor OK
				myXML.loadXML(WebServices.responseText)						
				
				Set RSSItems = myXML.getElementsByTagName("string")
				RSSItemsCount = RSSItems.Length-1				
				if (RSSItemsCount > 0) then					
					for i = 0 To RSSItemsCount
						Set RSSItem = RSSItems.Item(i)					
						for each child in RSSItem.childNodes
							select case i
								case 0									
									sDigitoVerificador = child.text
								case 1	' id bueno
									sMensaje = child.text
									'msgbox smensaje, ,"Registro Civil"
																	
									if sDigitoVerificador = ""	then ' id malo
										ConsularIdentificacion = "false"
									else
										ConsularIdentificacion = "true"
									end if
								case 3	' ERROR id malo
									sMensaje = child.text
									'msgbox sMensaje, ,"Registro Civil"
									ConsularIdentificacion = "false"
							end select
						next
					next
				end if
				
				ConsularIdentificacion = trim(ConsularIdentificacion) & ";true;" & sMensaje
			else			
				sMensaje = WebServices.statusText & vbcrlf & err.Description & vbcrl & WebServices.responseText 
				sMensajeRegistro = "Error en la consulta de la Identificación. " & vbcrlf & _
									"Detalle del ERROR: " & vbcrlf & _
									sMensaje
			end if
		end If
		
		if err.number <> 0 then
			sMensajeRegistro = "Error. " & err.Description
		end if
		
		Set WebServices = Nothing 
		Set myXML = Nothing		
		set xmlenviar = nothing
	End Function

	' JFMG 10-11-2008 agregado para que no se caiga la página con clientes que solo tienen identificación
	if nTipoCliente = "" then nTipoCliente = 1
	' ******************************** FIN ******************************


	Response.Expires = 0
%>
<html>
<head>
    <style>
        a:hover
        {
            color: blue;
        }
        INPUT.dINPUT
        {
            border-right: gray 1px solid;
            border-top: silver 1px solid;
            border-left: silver 1px solid;
            border-bottom: gray 1px solid;
        }
    </style>
    <meta name="GENERATOR" content="Microsoft Visual Studio 6.0">
    <meta http-equiv="CACHE-CONTROL" content="NO-CACHE, must-revalidate">
    <meta http-equiv="Pragma" content="no-cache">
    <title>Configuración Consulta de Giros</title>
    <link rel="stylesheet" type="text/css" href="../Estilos/Cliente.css">
    <link rel="stylesheet" type="text/css" href="../Estilos/AFX2004.css">
</head>

<script language="VBScript">
<!--
	On Error Resume Next

	Sub window_onload()
		<%	If sRut = Empty Then %>
				tdRut.style.display = "none"
				frmCliente.cmdValidarRegistro.style.display = "none"
		<%	Else %>
				tdPasaporte.style.display = "none"				
		<%	end If %>

		<%	If nTipoCliente <> 1 Then %>
				trEmpresa.style.display = ""
				trPersona.Style.display = "none"
				tdSexo.style.display ="none"
		<%	End If %>

		<% If nModo = 1 Then %>
				tdModoActualizacion.style.display = ""
		<% Else %>
				frmCliente.chkInformacionComercial.disabled = True
				frmCliente.chkInformacionTributaria.disabled = True
				frmCliente.chkPonderacion.disabled = True	
				frmCliente.chkAcreditacionComercial.disabled = True
				frmCliente.chkCarpetaCliente.disabled = True				
				frmCliente.chkClienteAgencia.disabled = true 
				if (frmCliente.optdeshabilitar.checked=true) then
					frmCliente.opthabilitar.disabled = True
				else
					frmCliente.optDeshabilitar.disabled = True
				end if
		<% End If %>
		
		if "<%=sMensajeRegistro%>" <> "" then
			msgbox "<%=sMensajeRegistro%>",,"Registro Civil"
		end if
		
		' Formatea montos
		frmcliente.txtMontoPatrimonio.value = formatnumber(ccur("0" + frmcliente.txtMontoPatrimonio.value),2)
		frmcliente.txtIngresoAnual.value = formatnumber(ccur("0" + frmcliente.txtIngresoAnual.value),2)
		frmcliente.txtMontoTransaccionesMesCLIENTE.value = formatnumber(ccur("0" + frmcliente.txtMontoTransaccionesMesCLIENTE.value),2)
		frmcliente.txtMontoTransaccionesMesAFEX.value = formatnumber(ccur("0" + frmcliente.txtMontoTransaccionesMesAFEX.value),2)
		
		' JFMG 09-12-2010 si el llamado es desde giros se deshabilita el estado y los datos de perfil operacional
		if cint("0" & <%=Session("TipoOrigenLlamada")%>) = cint(1) then
			frmCliente.opthabilitar.disabled = True
			frmCliente.optDeshabilitar.disabled = True
			
			frmCliente.cbxRiesgo.disabled = True
			frmCliente.cbxPerfilPEP.disabled = True
			frmCliente.cbxPerfilZona.disabled = True
			frmCliente.cbxPerfilRS.disabled = True
			frmCliente.cbxPerfilACT.disabled = True
			frmCliente.cbxPerfilIndustria.disabled = True
			
			frmcliente.cmdAgregarActividad.disabled = True 
			frmcliente.txtPropositoTransacciones.disabled = True
			frmcliente.txtReferenciaBancaria.disabled = True
			frmcliente.txtOrigenFondos.disabled = True
			frmcliente.txtMontoPatrimonio.disabled = True 
			frmcliente.txtIngresoAnual.disabled = True
			frmcliente.txtCantidadTransaccionesMesCLIENTE.disabled = True
			frmcliente.txtCantidadTransaccionesMesAFEX.disabled = True
			frmcliente.txtMontoTransaccionesMesCLIENTE.disabled = True
			frmcliente.txtMontoTransaccionesMesAFEX.disabled = True			
            frmcliente.cbxPerfilCliente.disabled = True		
			
		end if
		' FIN JFMG 09-12-2010
		
		If "<%=request("ActividadEconomica")%>" <> "" Then
		    CargarActividadEconomica
		End If
		
	End Sub
	Sub opthabilitar_onclick()
		frmcliente.optDeshabilitar.Checked=false
	End sub	
	
	Sub optDeshabilitar_onclick()
		frmcliente.opthabilitar.Checked=false				
	End sub	
	
	Sub NuevoCliente_onClick()
		window.navigate "http:NuevoCliente.asp"
	End Sub
	
	Sub NuevoDocumento_onClick()
		'window.navigate "http:NuevoDocumento.asp?cc=<%=Request("cc")%>&nc=<%=Request("nc")%>&rt=<%=Request("rt")%>"
'		Dim sString		
'		sString = Empty
'		sString = window.showModalDialog("http://desarrollo:91/subirarchivos.aspx?idcliente=<%=Request("cc")%>&origen=3&usuario=" & Session("NombreUsuarioOperador") & "nombre=" & replace(Session("NombreOperador")," ","%20"))
	End Sub		

	Sub NuevaHistoria_onClick()
		window.navigate "http:NuevaHistoria.asp?cc=<%=Request("cc")%>&nc=<%=Request("nc")%>&rt=<%=Request("rt")%>"
	End Sub		

	Sub tdHistoria_onClick()
		If tbHistoria.style.display = "none" Then
			tbHistoria.style.display = ""
		Else
			tbHistoria.style.display = "none"
		End If
	End Sub
	
	Sub tdDocumento_onClick()
		If tbDocumento.style.display = "none" Then
			tbDocumento.style.display = ""
		Else
			tbDocumento.style.display = "none"
		End If
	End Sub

	Const sEncabezadoFondo = "Consultas"
	Const sEncabezadoTitulo = "Detalle Cliente"
	Const sClass = "TituloPrincipal"


	Sub cbxPaisPersonal_onblur()
		
		If frmCliente.cbxPaisPersonal.value = "" Then Exit Sub
		If frmCliente.cbxPaisPersonal.value = "<%=sPais%>" Then Exit Sub
		HabilitarControles		
		frmCliente.action = "http:DetalleCliente.asp?acc=1&cc=<%=Request("cc")%>&mnu=<%=nMenu%>&md=1" & "&CA=" & frmCliente.chkClienteAgencia.checked & _
		                    "&ActividadEconomica=" & frmcliente.txtActividadEconomica.value 
		frmCliente.submit 
		frmCliente.action = ""
		
	End Sub

	Sub cbxCiudadPersonal_onblur()
		
		If frmCliente.cbxCiudadPersonal.value = "" Then Exit Sub
		If frmCliente.cbxCiudadPersonal.value = "<%=sCiudad%>" Then Exit Sub
		HabilitarControles
		frmCliente.action = "http:DetalleCliente.asp?acc=1&cc=<%=Request("cc")%>&mnu=<%=nMenu%>&md=1" & "&CA=" &frmCliente.chkClienteAgencia.checked
		frmCliente.submit 
		frmCliente.action = ""
		
	End Sub

	Sub cbxSucursal_onblur()
		
		If frmCliente.cbxSucursal.value = "" Then Exit Sub
		If frmCliente.cbxSucursal.value = "<%=sSucursal%>" Then Exit Sub
		HabilitarControles		
		frmCliente.action = "http:DetalleCliente.asp?acc=1&cc=<%=Request("cc")%>&mnu=<%=nMenu%>&md=1" & "&CA=" &frmCliente.chkClienteAgencia.checked
		frmCliente.submit 
		frmCliente.action = ""
		
	End Sub

	Sub cbxRubro_onblur()
		If frmCliente.cbxRubro.value = "" Then Exit Sub
		If frmCliente.cbxRubro.value = "<%=nRubro%>" Then Exit Sub
	End Sub

'* GuargarCambios está en MenuActualizacionCliente.asp *************************
	Sub GuardarCambios_onclick()
		Dim Shabilita
	
    	If Not ValidarInformacion Then Exit Sub

		' valida que agregue los contactos
		if frmCliente.optdeshabilitar.checked then
			frmCliente.txtContacto1.value = ""
			frmCliente.txtPorcentageContacto1.value = ""
			frmCliente.txtContacto2.value = ""
			frmCliente.txtPorcentageContacto2.value = ""
			frmCliente.txtFechaActivacionComision.value = ""		
		end if
		
		HabilitarControles	
		
		' JFMG 20-06-2011 desde hoy ya no se solicitará la sucursal		      
		'' solicita la ip de la sucursal para ir a la bd y verificar al cliente actualizado
		'frmCliente.txtSucursalSolicitante.value = window.showModalDialog("sucursalsolicitante.asp")		
		' FIN JFMG 20-06-2011

		AsignarActividadEconomica
		
		' MS 28-03-2014
		If frmCliente.optActEconomicaSi.checked and frmCliente.txtActividadEconomica.value = empty then 
		    alert("Debe seleccionar la actividad económica")
		    exit sub
		End if
		' FIN MS 28-03-2014
		    
		frmCliente.action = "http:GuardarActualizacionCliente.asp?cc=<%=Request("cc")%>&tc=" & frmCliente.cbxtipo.value & _
															 "&ca=<%=nCredito%>" & "&CAG=" & frmCliente.chkClienteAgencia.checked 
															 
		frmCliente.submit 
		frmCliente.action = ""
	End Sub
	
    'INTERNO-6830 MS 14-07-2016
    Function ValidarInformacion
		Dim sString
		
		ValidarInformacion = False
		
		sString = Empty

		If frmCliente.cbxTipo.value = 1 or frmCliente.cbxTipo.value = 7 or frmCliente.cbxTipo.value = 8 or frmCliente.cbxTipo.value = 9  then 'Si selecciono Persona
			If frmCliente.txtNombres.value = Empty Then sString = sString & "Nombres, " 
			If frmCliente.txtApellidoP.value = Empty Then sString = sString & "Apellido Paterno, "
			If frmCliente.txtApellidoM.value = Empty Then sString = sString & "Apellido Materno, "
			If frmCliente.txtFechaNacimiento.value  = Empty Then sString = sString & "Fecha Nacimiento, "

		Else 'Si seleccionó empresa
			If frmCliente.txtRazonSocial.value = Empty Then sString = sString & "Razón Social, "
			If frmCliente.txtRepresentante.value = Empty Then sString = sString & "Representante Legal, "
		End If		
		
		If frmCliente.cbxTipo.value = 1 then
			If frmCliente.cbxSexo.value = "0" Then sString = sString & "Sexo, "
            If frmCliente.cbxOcupacion.value = "0" or frmCliente.cbxOcupacion.value = Empty Then sString = sString & "Ocupación, "
            If frmCliente.cbxNacionalidad.value = "0" or frmCliente.cbxNacionalidad.value = Empty Then sString = sString & "Nacionalidad, "
            'If frmCliente.txtcorreo.value = Empty Then sString = sString & "Correo electrónico, "
			If frmCliente.txtDireccionPersonal.value = Empty Then sString = sString & "Dirección, "
			If frmCliente.txtNumero.value = empty then sString = sstring & "Número dirección, "
			If frmCliente.cbxPaisPersonal.value = Empty Then sString = sString & "País, "
			If frmCliente.cbxCiudadPersonal.value = Empty Then sString = sString & "Ciudad, "
			If UCase(frmCliente.cbxPaisPersonal.value) = "CL" Then
				If frmCliente.cbxComunaPersonal.value = Empty Then sString = sString & "Comuna, "
			End If

			If frmCliente.txtCelular.value = Empty and frmCliente.txtFonoPersonal.value = Empty then sString = sString & "Teléfono, "
			If frmCliente.optActEconomicaSi.checked and frmCliente.txtActividadEconomica.value = empty then sString = sString & "Actividad Económica, "
		End if
		If sString <> Empty Then
            MsgBox "Debe ingresar el (los) siguiente(s) campo(s): " & Left(sString, len(sString)-2), vbOKOnly + vbInformation, "Actualizar Cliente"
            exit Function
        end if

		ValidarInformacion = True
	End Function
    'FIN INTERNO-6830 MS 14-07-2016

	Sub HabilitarControles()
		frmCliente.txtRut.Disabled = False
		
		' JFMG 14-11-2008 antes no se consederaba la serie del rut
		frmCliente.txtnumeroSerieRut.disabled = False
		' ************************** FIN ************************
		frmCliente.txtPasaporte.Disabled = False
		frmCliente.txtNombres.Disabled = False
		frmCliente.txtNombreCompleto.Disabled = False
		frmCliente.txtApellidoP.Disabled = False
		frmCliente.txtApellidoM.Disabled = False
		frmCliente.txtRazonSocial.Disabled = False
		frmCliente.txtCodigoCliente.Disabled = False
		frmCliente.txtPaisFonoPersonal.Disabled = False
		frmCliente.txtAreaFonoPersonal.Disabled = False
		frmCliente.txtPaisFaxPersonal.Disabled = False
		frmCliente.txtAreaFaxPersonal.Disabled = False
		frmCliente.opthabilitar.Disabled = False
		frmCliente.optdeshabilitar.Disabled = False
		frmCliente.txtUsuario.disable = False	
		frmCliente.txtPassword.disable = False
		frmCliente.txtPregunta.disable = False	
		frmCliente.txtRespuesta.disable = False
		frmCliente.cbxRiesgo.disable = False
		frmCliente.cbxPerfilPEP.disable = False
		frmCliente.cbxPerfilZona.disable = False
        frmCliente.cbxPerfilCliente.disabled = False
		frmCliente.cbxPerfilRS.disable = False
		frmCliente.cbxPerfilACT.disable = False
		frmCliente.cbxPerfilIndustria.disable = False
		
		frmcliente.txtPropositoTransacciones.disabled = False
		frmcliente.txtReferenciaBancaria.disabled = False
		frmcliente.txtOrigenFondos.disabled = False
		frmcliente.txtMontoPatrimonio.disabled = False
		frmcliente.txtIngresoAnual.disabled = False
		frmcliente.txtCantidadTransaccionesMesCLIENTE.disabled = False
		frmcliente.txtCantidadTransaccionesMesAFEX.disabled = False
		frmcliente.txtMontoTransaccionesMesCLIENTE.disabled = False
		frmcliente.txtMontoTransaccionesMesAFEX.disabled = False
		frmcliente.cbxActividadEconomica.disabled = False
		frmcliente.cmdAgregarActividad.disabled = False
		frmcliente.txtNombreEmpleador.disabled = False 
		frmcliente.txtGiroComercial.disabled = false 'JBV17-05-2012
	End Sub
	
	Function window_onscroll(x)
	End Function

	Sub txtCredito_onBlur()
		frmCliente.txtCredito.value = FormatNumber(frmCliente.txtCredito.value, 2)		
		frmCliente.txtCreditoDisponible.value = FormatNumber(cCur(0 & frmCliente.txtCredito.value) - cCur(0 & frmCliente.txtCreditoUsado.value), 2)
	End Sub

	Sub DetalleCheques()
		Dim sString		
		sString = Empty
		sString = window.showModalDialog("DetalleCheques.asp?cc=<%=nCodigoCli%>&dr=<%=nDiasRetencion%>")
	End Sub
	
	' JFMG 19-10-2007 se agregó combo con mas tipos
	Sub cbxTipo_onChange()
		
		if frmCliente.cbxTipo.value <> 1 then
			trEmpresa.style.display = ""	
			frmCliente.cbxRubro.value = ""
			trContacto.style.display = "none"
			trPersona.style.display = "none"
			tdSexo.style.display = "none"
			frmCliente.txtApellidoM.value = ""
			frmCliente.txtApellidoP.value = ""
			frmCliente.txtNombres.value = ""
			frmcliente.txtNombreEmpleador.value = ""
		else
			trEmpresa.style.display = "none"
			frmCliente.cbxRubro.value = ""
			trContacto.style.display = "none"
			trPersona.style.display = ""
			tdSexo.style.display =""
			frmCliente.txtApellidoMContacto.value = ""
			frmCliente.txtApellidoPContacto.value = ""
			frmCliente.txtNombresContacto.value = ""
			frmCliente.txtRazonSocial.value = ""
			frmCliente.txtRepresentante.value = ""
			
		end if
	end sub
		
	Sub cmdValidarRegistro_onClick()
		Dim sRut, sDigito, sPasaporte, sRespuestaRegistro, sNumeroRut, sTipo
		dim bIDValida, bConsulta
		dim sNumeroSerie
		
		If (frmCliente.txtRut.value = Empty and frmCliente.txtPasaporte.value = Empty) Then Exit Sub		
		
		sNumeroSerie = inputbox("Ingrese el número de serie de la Identificación:","AFEX","<%=sNumeroSerieID%>")
		if sNumeroSerie = "" then exit sub
		
		If frmCliente.txtRut.value <> "" Then
			sTipo = "CEDUL"
			sRut = ValidarRut(frmCliente.txtRut.value)
			
			frmCliente.txtUsuario.value = frmCliente.txtRut.value
			frmCliente.txtRut.value = sRut
		
			sNumeroRut = replace(replace((trim(sRut)),"-",""),".","")
			sNumeroRut = left(sNumeroRut, len(sNumeroRut) - 1)
		
		elseif frmCliente.txtPasaporte.value <> "" Then
			sTipo = "PASAP"
			sNumeroRut = frmCliente.txtPasaporte.value
			
		Else
			msgbox "Debe ingresar una Identificación ", vbOKOnly + vbInformation, "Ingreso de Cliente"			
			Exit Sub
		End If		
		
		window.navigate "DetalleCliente.asp?cc=<%=nCodigoCliente%>&Accion=6&NumeroSerie=" & trim(sNumeroSerie) & "&NumeroRut=" & sNumeroRut & "&Tipo=" & sTipo
	End Sub
		
		
	' JFMG 02-02-2010
	sub EliminarDocumento(byval IdDocumento, byval TipoDocumento)
		'MS 18-3-2014
		if msgbox("¿Está seguro de DESHABILITAR el documento?", 1, "AFEX") <> 1 then exit sub 'MS 10-03-2014
		Dim sString		
		sString = Empty
		sString = window.showModalDialog("Motivo.asp")
		
		window.navigate "DetalleCliente.asp?elim=1&cc=<%=Request("cc")%>&doc=" & IdDocumento & "&tip=" & TipoDocumento & "&motivo=" & sString
		'FIN MS 18-3-2014
	end sub
	' *********** FIN JFMG 02-02-2010
	 
	sub HabilitarDocumento(byval IdDocumento, byval TipoDocumento)
		'MS 18-3-2014
		if msgbox("¿Está seguro de HABILITAR el documento?", 1, "AFEX") <> 1 then exit sub 'MS 10-03-2014
		Dim sString		
		sString = Empty
		sString = window.showModalDialog("Motivo.asp")
		
		window.navigate "DetalleCliente.asp?elim=2&cc=<%=Request("cc")%>&doc=" & IdDocumento & "&tip=" & TipoDocumento & "&motivo=" & sString
		'FIN MS 18-3-2014
	end sub
	
	' JFMG 22-06-2011
	sub cmdAgregarActividad_onClick()
        dim sTabla, sFila, sActividad
            
        if ccur("0" + frmcliente.cbxActividadEconomica.value) = 0 then exit sub
            
        if VerificarActividadEconomicaExiste(frmcliente.cbxActividadEconomica.value) then exit sub
            
        sActividad = frmcliente.cbxActividadEconomica.options(frmcliente.cbxActividadEconomica.selectedIndex).text
        sFila = "<tr style=""cursor: hand; background-color: White;"" onmouseover=""javascript:this.bgColor='#f4f4f4';"" onmouseout=""javascript:this.bgColor='white'"">" 
        sFila = sFila & "<td style=""display: none;"">" & frmcliente.cbxActividadEconomica.value & "</td><td style=""color: Blue;"">" & sActividad & "</td>"
        sFila = sFila & "<td><IMG src=""../images/elimsup.jpg"" border=""0"" onclick=""EliminarActividad(" & frmcliente.cbxActividadEconomica.value & ")"" ALT=""Presione aquí para eliminar"" /></td>"
        sFila = sFila & "</tr></TBODY>"
            
        sTabla = tblactividadeconomica.outerHTML
                      
        sTabla = replace(sTabla, "</TBODY>", sFila)
                       
        tblactividadeconomica.outerHTML = sTabla
            
        frmcliente.txtActividadEconomica.value = frmcliente.cbxActividadEconomica.value & ";" & frmcliente.txtActividadEconomica.value 
    end sub
	
	sub EliminarActividad(ByVal Codigo)
        dim i
            
        for i = 1 to window.tblActividadEconomica.rows.length - 1
            if trim(window.tblActividadEconomica.rows(i).cells(0).innerText) = trim(Codigo) then
                dim bEliminado
                    
                bEliminado = window.showmodaldialog("EliminarActividadEconomica.asp?Actividad=" & Codigo & "&cc=" & "<%=Request("cc")%>")
                                        
                If left(bEliminado, 1) = 1 then
                    window.tblActividadEconomica.deleteRow i
                    frmcliente.txtActividadEconomica.value = replace(trim(frmcliente.txtActividadEconomica.value), trim(Codigo) & ";", "")
                    frmCliente.txtActividadEconomicaInactiva.value = frmcliente.txtActividadEconomica.value & ";" & frmcliente.txtActividadEconomicaInactiva.value
                    msgbox "Actividad Económica Eliminada.",,"AFEX"
                    exit for
                Else
                    msgbox mid(bEliminado,2),,"AFEX"
                    exit for
                End If
            end if
        next
    end sub        
        
    function VerificarActividadEconomicaExiste(ByVal CodigoActividad)
            
        VerificarActividadEconomicaExiste = False
            
        if instr(frmcliente.txtActividadEconomica.value, CodigoActividad) > 0 or _
        instr(frmcliente.txtActividadEconomicaInactiva.value, CodigoActividad) > 0 then
            VerificarActividadEconomicaExiste = True
        end if            
            
    end function
	
	sub AsignarActividadEconomica()
        dim i
        frmcliente.txtActividadEconomica.value = ""      
        for i = 1 to tblactividadeconomica.rows.length - 1
            frmcliente.txtActividadEconomica.value = tblactividadeconomica.rows(i).cells(0).innerTEXT & ";" & frmcliente.txtActividadEconomica.value                
        next 
           
    end sub
        
    sub CargarActividadEconomica()
        if frmcliente.txtActividadEconomica.value = "" then exit sub
            
        dim sTabla, sFila, sActividad
        dim sActividadEconomicai
        sActividadEconomicai = split(frmcliente.txtActividadEconomica.value, ";")
                               
        for i = 0 to UBOUND(sActividadEconomicai) - 1
            if sActividadEconomicai(i) <> "" then
                frmcliente.cbxActividadEconomicaOculta.value = sActividadEconomicai(i)
                sActividad = frmcliente.cbxActividadEconomicaOculta.options(frmcliente.cbxActividadEconomicaOculta.selectedIndex).text
                sFila = "<tr style=""cursor: hand; background-color: White;"" onmouseover=""javascript:this.bgColor='#f4f4f4';"" onmouseout=""javascript:this.bgColor='white'"">" 
                sFila = sFila & "<td style=""display: none;"">" & frmcliente.cbxActividadEconomicaOculta.value & "</td><td style=""color: Blue;"">" & sActividad & "</td>" 
                sFila = sFila & "<td><IMG src=""../images/elimsup.jpg"" border=""0"" onclick=""EliminarActividad(" & frmcliente.cbxActividadEconomicaOculta.value & ")"" ALT=""Presione aquí para eliminar"" /></td>"
                sFila = sFila & "</tr>"
                    
            end if
        next
        sFila = sFila & "</TBODY>"
            
        sTabla = tblactividadeconomica.outerHTML            
        sTabla = replace(sTabla, "</TBODY>", sFila)            
        tblactividadeconomica.outerHTML = sTabla
            
    end sub
    
    Sub txtIngresoAnual_onBlur()
        frmcliente.txtIngresoAnual.value = formatnumber(ccur("0" + frmcliente.txtIngresoAnual.value), 2)
    End Sub
    
    Sub txtMontoPatrimonio_onBlur()
        frmcliente.txtMontoPatrimonio.value = formatnumber(ccur("0" + frmcliente.txtMontoPatrimonio.value), 2)
    End Sub
    
    Sub txtMontoTransaccionesMesAFEX_onBlur()
        frmcliente.txtMontoTransaccionesMesAFEX.value = formatnumber(ccur("0" + frmcliente.txtMontoTransaccionesMesAFEX.value), 2)
    End Sub
    
    Sub txtMontoTransaccionesMesCliente_onBlur()
        frmcliente.txtMontoTransaccionesMesCliente.value = formatnumber(ccur("0" + frmcliente.txtMontoTransaccionesMesCliente.value), 2)
    End Sub
    
	' FIN JFMG 22-06-2011
	
	'MS 10-03-2014
	sub AutorizarDocumento(byval IdDocumento, byval TipoDocumento)
		
		if msgbox("¿Está seguro de AUTORIZAR este documento?", 1, "AFEX") <> 1 then exit sub
		
		window.navigate "DetalleCliente.asp?aut=1&cc=<%=Request("cc")%>&doc=" & IdDocumento & "&tip=" & TipoDocumento
	end sub
	
	sub RechazarDocumento(byval IdDocumento, byval TipoDocumento)
		
		if msgbox("¿Está seguro de RECHAZAR este documento?", 1, "AFEX") <> 1 then exit sub
		Dim sString		
		sString = Empty
		sString = window.showModalDialog("Motivo.asp?")
		window.navigate "DetalleCliente.asp?aut=2&cc=<%=Request("cc")%>&doc=" & IdDocumento & "&tip=" & TipoDocumento & "&motivo=" & sString
	end sub
	'FIN MS 10-03-2014
	
	'MS 26-03-2014
	
	Sub cmdActualizarDocumentos_onClick()
        window.navigate "DetalleCliente.asp?cc=<%=Request("cc")%>&nc=<%=Request("nc")%>&rt=<%=Request("rt")%>"
    End Sub
	'MS FIN 26-03-2014
	
	'MS 28-03-2014
	Sub optActEconomicaSi_onClick()
		frmcliente.optActEconomicaNo.checked = 0
		frmcliente.optActEconomicaSi.checked = 1
	End Sub
	
	Sub optActEconomicaNo_onClick()
		frmcliente.optActEconomicaNo.checked = 1
		frmcliente.optActEconomicaSi.checked = 0
	End Sub
	'FIN 28-03-2014
	
//-->
</script>

<!--INCLUDE virtual="/Compartido/Encabezado.htm" -->
<body id="bb" border="0" style="margin: 2 2 2 2">
    <form id="frmCliente" method="post">
    <input type="hidden" name="txtTipoCliente" value="<%=nTipoCliente%>">
    <input type="hidden" name="txtNombreCompleto" value="<%=sNombreCompleto%>">
    <table class="Borde" id="tabConsulta" border="0" cellpadding="0" cellspacing="0"
        style="height: 150px; width: 100%; background-color: #f4f4f4">
        <tr height="40" style="background-color: #ffeeaa; #ffdd77; #e1e1e1">
            <td colspan="3" style="font-size: 16pt">
                &nbsp;&nbsp;<%=sApellidoP%>&nbsp;<%=sApellidoM%>&nbsp;<%=sNombres%>
            </td>
        </tr>
        <tr height="1" style="background-color: silver">
            <td colspan="3">
            </td>
        </tr>
        <tr height="4">
            <td colspan="3">
            </td>
        </tr>
        <tr>
            <td colspan="3">
                <table width="100%">
                    <tr>
                        <td>
                            <table style="font-size: 8pt;">
                                <tr>
                                    <td id="tdRut" colspan="2">
                                        Rut<br />
                                        <input class="dInput" name="txtRut" value="<%=FormatoRut(sRut)%>" style="width: 100px"
                                            disabled />
                                        <input class="dInput" name="txtNumeroSerieRut" value="<%=sNumeroSerieID%>" style="width: 100px;"
                                            disabled />
                                    </td>
                                    <td id="tdPasaporte">
                                        Pasaporte<br />
                                        <input class="dInput" name="txtPasaporte" value="<%=sPasaporte%>" style="width: 100px"
                                            disabled />
                                        <input class="dInput" name="txtNombrePaisPasaporte" value="<%=sNombrePaisPasaporte%>"
                                            style="width: 100px" disabled />
                                    </td>
                                    <td>
                                        &nbsp;<br />
                                        <input type="button" name="cmdValidarRegistro" value="Validar Rut" />
                                    </td>
                                    <td>
                                        Código<br />
                                        <input class="dInput" name="txtCodigoCliente" value="<%=Request("cc")%>" style="width: 60px"
                                            disabled />
                                    </td>
                                    <td>
                                        Fecha Creación<br />
                                        <input class="dInput" name="txtFechaCreacion" value="<%=sFechaCreacion%>" style="width: 90px"
                                            disabled />
                                    </td>
                                    <td>
                                        <%if (sEstado=1) then %>
                                        <font face="verdana" color="green"><b>Habilitado</b></font><br />
                                        <input type="radio" name="opthabilitar" style="width: 90px" checked />
                                        <%else%>
                                        <font face="verdana" color="black">Habilitado</font><br />
                                        <input type="radio" name="opthabilitar" />
                                        <%end if%>
                                    </td>
                                    <td>
                                        <%if (sEstado=0) then %>
                                        <font face="verdana" color="red"><b>Deshabilitado</b></font><br />
                                        <input type="radio" name="optdeshabilitar" checked />
                                        <%else%>
                                        <font face="verdana" color="black">Deshabilitado</font><br />
                                        <input type="radio" name="optdeshabilitar" />
                                        <%end if%>
                                    </td>
                                </tr>
                            </table>
                        </td>
                        <td align="right">
                            <table>
                                <tr>
                                    <%If bMostrarMenu Then %>
                                    <!-- JFMG 04-07-2012 -->
                                    <%' JFMG 22-05-2010 
									    if Session("TipoOrigenLlamada") = 1 then %>
                                    <td colspan="2">
                                        <% If nMenu = 0 Then %>
                                        <!--#INCLUDE virtual="/Sucursal/MenuDetalleClienteCajeroSucursal.htm" -->
                                        <% Else %>
                                        <!--#INCLUDE virtual="/Sucursal/MenuActualizarCliente.htm" -->
                                        <% End If %>
                                    </td>
                                    <%else%>
                                    <td colspan="2">
                                        <% If nMenu = 0 Then %>
                                        <!--#INCLUDE virtual="/Sucursal/MenuDetalleCliente.htm" -->
                                        <% Else %>
                                        <!--#INCLUDE virtual="/Sucursal/MenuActualizarCliente.htm" -->
                                        <% End If %>
                                    </td>
                                    <%end if%>
                                    <%End If%>
                                    <!-- JFMG 04-07-2012 -->
                                </tr>
                            </table>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
        <!-- ' JFMG 21-09-2010 busca las imagenes del cliente -->
        <tr>
            <td>
            <!-- APPL-6471 MS 20-07-2015 -->
			 <!-- FIN APPL-6471 MS 20-07-2015 -->
            </td>
        </tr>
        <!-- ' FIN JFMG 21-09-2010 -->
        <tr height="10">
            <td>
            </td>
        </tr>
        <tr>
            <td>
                <table cellspacing="1" cellpadding="1" swidth="30%" style="font-family: Verdana;
                    font-size: 10pt; position: relative; top: 0px; sborder: 1; background-color: silver;">
                    <tr height="22">
                        <td id="tdDocumento" style="background-color: #ffffcc; #e1e1e1; cursor: hand">
                            <b>&nbsp;&nbsp;Datos del Cliente&nbsp;&nbsp;</b>
                        </td>
                        <td id="tdModoActualizacion" style="background-color: #ccddee; #e1e1e1; display: none">
                            <b>&nbsp;&nbsp;Modo Actualización&nbsp;&nbsp;</b>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
        <tr height="1" style="background-color: silver">
            <td colspan="3">
            </td>
        </tr>
        <tr height="10">
            <td>
            </td>
        </tr>
        <!-- ************************************ Nuevo ************************************** -->
        <tr height="20">
            <td>
                <table border="0" cellspacing="0" cellpadding="0">
                    <tr>
                        <td>
                            <table>
                                <tr>
                                    <td width="135">
                                        Tipo Cliente<br />
                                        <select name="cbxTipo" <%=sHabilitado%>>
                                            <%CargarTipo "CLIENTE", nTipoCliente%>
                                        </select>
                                    </td>
                                    <td width="150" id="tdSexo">
                                        Sexo<br />
                                        <select name="cbxSexo" <%=sHabilitado%>>
                                            <%CargarTipo "SEXO", sSEXO%>
                                        </select><span>*</span>
                                    </td>
                                    <td width="148">
                                        &nbsp;
                                        <input id="chkCLienteAgencia" type="checkbox" <%=nclienteAgencia%> />
                                        Cliente agencia
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
        <tr>
            <td>
                <table>
                    <tr id="trEmpresa" height="20" style="display: none">
                        <td>
                            <table>
                                <tr>
                                    <td>
                                        Razón Social<br />
                                        <input name="txtRazonSocial" id="txtRazonSocial" style="width: 200px" onkeypress="IngresarTexto(3)"
                                            maxlength="60" value="<%=sRazonSocial%>" <%=sHabilitado%> /> <span>*</span>
                                    </td>
                                    <td>
                                        Representante Legal<br />
                                        <input name="txtRepresentante" id="txtRepresentante" style="width: 200px" onkeypress="IngresarTexto(3)"
                                            maxlength="60" value="<%=sRepresentante%>" <%=sHabilitado%> />
                                    </td>
                                    <td>
                                        Rubro<br />
                                        <select name="cbxRubro" style="width: 255px" <%=sHabilitado%>>
                                            <% 
											CargarRubro nRubro
                                            %>
                                        </select>
                                    </td>
                                </tr>
                                <tr>
                                    <td colspan="3">
                                        Giro Comercial<br />
                                        <input name="txtGiroComercial" id="txtGiroComercial" style="width: 650px" onkeypress="IngresarTexto(3)"
                                            maxlength="60" value="<%=sGiroComercial%>" <%=sHabilitado%> />
                                    </td>
                                </tr>
                            </table>
                            <!--														<table>
								<tr>
									<td>Correo Electrónico<br>
										<input name="txtCorreo" "IngresarTexto(6)" maxlength="42" style="width: 310px" value="<%=sCorreo%>" <%=sHabilitado%>>									</td>
								</tr>
							</table>-->
                        </td>
                    </tr>
                    <tr id="trPersona" height="20">
                        <td>
                            <table>
                                <tr>
                                    <td>
                                        Apellido Paterno<br/>
                                        <input name="txtApellidoP" id="txtApellidoP" onkeypress="IngresarTexto(2)" style="width: 180px"
                                            maxlength="20" value="<%=sApellidoP%>" <%=sHabilitado%> /><span>*</span>
                                    </td>
                                    <td>
                                        Apellido Materno<br/>
                                        <input name="txtApellidoM" id="txtApellidoM" onkeypress="IngresarTexto(2)" style="width: 180px"
                                            maxlength="20" value="<%=sApellidoM%>" <%=sHabilitado%> /><span>*</span>
                                    </td>
                                    <td>
                                        Nombres<br />
                                        <input name="txtNombres" id="txtNombres" style="width: 308px" onkeypress="IngresarTexto(2)"
                                            maxlength="30" value="<%=sNombres%>" <%=sHabilitado%> /><span>*</span>
                                    </td>
                                </tr>
                                <tr>
                                    <td colspan="2">
                                        Fecha nacimiento<br/>
                                        <input name="txtFechaNacimiento" id="txtFechaNacimiento" size="12" maxlength="10"
                                            value="<%=sFechaNacimiento%>" <%=sHabilitado%>/><span>*</span>(dd-mm-aaaa)
                                    </td>
                                    <td colspan="1">
                                        Ocupación<br/>
                                       <select name="cbxOcupacion" style="width: 308px" <%=sHabilitado%>>
                                            <%	CargarOcupacion sOcupacion %>
                                        </select><span>*</span><!--INTERNO-9263	MS 19-01-2017-->
                                    </td>
                                </tr>
                                <tr>
                                    <td colspan="2">
                                        Profesión<br/>
                                        <input name="txtProfesion" id="txtProfesion" onkeypress="IngresarTexto(2)" size="30"
                                            maxlength="100" value="<%=sProfesion%>" <%=sHabilitado%> style="width:300px" />
                                    </td>
                                    <td colspan="1">
                                        Nombre Empleador<br />
                                        <input name="txtNombreEmpleador" id="txtNombreEmpleador" <%=sHabilitado%> onkeypress="IngresarTexto(2)"
                                            size="50" maxlength="100" value="<%=sNombreEmpleador%>" style="width:308px" />
                                    </td>
                                </tr>
                            </table>
                            <table>
                                <tr>
                                    <td>
                                        Nacionalidad<br/>
                                        <select name="cbxNacionalidad" style="width: 233px" <%=sHabilitado%>>
                                            <%	CargarUbicacion 1, "", sNacionalidad 	%>
                                        </select><span>*</span>
                                    </td>
                                    <!--	<td>Correo Electrónico<br>
										<input name="txtCorreo" "IngresarTexto(6)" maxlength="42" style="width: 310px" value="<%=sCorreo%>" <%=sHabilitado%>>									</td>-->
                                </tr>
                            </table>
                        </td>
                    </tr>
                    <tr height="20">
                        <td>
                            <table>
                                <tr>
                                    <td>
                                        Correo Electrónico<br/>
                                        <input name="txtCorreo" "IngresarTexto(6)" maxlength="42" style="width: 310px" value="<%=sCorreo%>"
                                            <%=sHabilitado%> /><span>*</span>
                                    </td>
                                </tr>
                            </table>
                            <table width="77%" border="0" cellspacing="1" cellpadding="0">
                                <tr>
                                    <td width="56%">
                                        Calle
                                    </td>
                                    <td width="16%">
                                        N&uacute;mero
                                    </td>
                                    <td width="28%">
                                        Depto/Oficina
                                    </td>
                                </tr>
                                <tr>
                                    <td>
                                        <input name="txtDireccionPersonal" onkeypress="IngresarTexto(3)" maxlength="40" style="width: 270px"
                                            value="<%=sDireccion%>" <%=sHabilitado%> /> <span>*</span>
                                    </td>
                                    <td>
                                        <input name="txtNumero" onkeypress="IngresarTexto(3)" maxlength="10" style="width: 70px"
                                            value="<%=sNumero%>" <%=sHabilitado%> /><span>*</span>
                                    </td>
                                    <td>
                                        <input name="txtDepto" onkeypress="IngresarTexto(3)" maxlength="10" style="width: 70px"
                                            value="<%=sDepto%>" <%=sHabilitado%> />
                                    </td>
                                </tr>
                                <tr>
                                    <td>
                                        Pa&iacute;s
                                    </td>
                                    <td>
                                        Ciudad
                                    </td>
                                    <td>
                                        &nbsp;
                                    </td>
                                </tr>
                                <tr>
                                    <td>
                                        <select name="cbxPaisPersonal" style="width: 135px" <%=sHabilitado%>>
                                            <%	CargarUbicacion 1, "", sPais 	%>
                                        </select><span>*</span>
                                    </td>
                                    <td colspan="2">
                                        <select name="cbxCiudadPersonal" style="width: 135px" <%=sHabilitado%>>
                                            <% CargarCiudadesPais sPais, sCiudad %>
                                        </select>
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                    <tr height="20">
                        <td>
                            <table>
                                <tr>
                                    <td>
                                        Comuna<br />
                                        <select class="dinput" name="cbxComunaPersonal" style="width: 165px" <%=sHabilitado%>>
                                            <% 
										If sPais = "CL" Then						
											CargarComunaCiudad  sCiudad, sComuna
										End If
                                            %>
                                        </select><span>*</span>
                                    </td>
                                    <td>
                                        Teléfono<br />
                                        <input name="txtPaisFonoPersonal" style="width: 30px" value="<%=sPaisFono%>" disabled />
                                        <input name="txtAreaFonoPersonal" style="width: 30px" value="<%=sAreaFono%>" disabled />
                                        <input name="txtFonoPersonal" style="width: 80px" maxlength="10" value="<%=sNumeroFono%>"
                                            <%=sHabilitado%> /><span>*</span>
                                    </td>
                                    <td>
                                        Fax<br />
                                        <input name="txtPaisFaxPersonal" style="width: 30px" value="<%=sPaisFax%>" disabled />
                                        <input name="txtAreaFaxPersonal" style="width: 30px" value="<%=sAreaFax%>" disabled />
                                        <input name="txtFaxPersonal" style="width: 80px" maxlength="10" value="<%=sNumeroFax%>"
                                            <%=sHabilitado%> />
                                    </td>
                                    <td>
                                        Celular<br />
                                        <input name="txtCelular" style="width: 90px" maxlength="10" value="<%=sCelular%>"
                                            <%=sHabilitado%> />
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <table>
                                <tr>
                                    <td>
                                        Sucursal<br />
                                        <select name="cbxSucursal" style="width: 240px" <%=sHabilitado%>>
                                            <% 
										CargarSucursal sSucursal
                                            %>
                                        </select>
                                    </td>
                                    <td>
                                        Ejecutivo<br />
                                        <select name="cbxEjecutivos" style="width: 305px" <%=sHabilitado%>>
                                            <% 
										CargarEjecutivos sSucursal, nEjecutivo
                                            %>
                                        </select>
                                    </td>
                                </tr>
                            </table>
                            <table>
                                <tr>
                                    <td>
                                        Banco<br />
                                        <select name="cbxBanco" style="width: 240px" <%=sHabilitado%>>
                                            <% 
										CargarBanco sBanco
                                            %>
                                        </select>
                                    </td>
                                    <td>
                                        Cuenta Corriente<br />
                                        <input name="txtCuentaCorriente" style="width: 150px" maxlength="20" value="<%=sCtaCte%>"
                                            <%=sHabilitado%> />
                                    </td>
                                    <td>
                                        Cuenta de Ahorro<br />
                                        <input name="txtCuentaAhorro" style="width: 150px" maxlength="20" value="<%=sCtaAhorro%>"
                                            <%=sHabilitado%> />
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <table width="100%">
                                <tr height="15">
                                    <td>
                                        Crédito<br />
                                        <input name="txtCredito" value="<%=FormatNumber(nCredito, 2)%>" onkeypress="IngresarTexto(1)"
                                            style="text-align: right; height: 22px; width: 155px" <%=sHabilitado%>>
                                    </td>
                                    <td>
                                        Crédito Usado<br />
                                        <input name="txtCreditoUsado" value="<%=FormatNumber(nCreditoUsado, 2)%>" style="text-align: right;
                                            height: 22px; width: 155px" disabled />
                                    </td>
                                    <td>
                                        <img border="0" id="imgTransferencia" src="../images/Transferencia.jpg" style="cursor: hand"
                                            width="19" height="15" onclick="DetalleCheques">
                                    </td>
                                    <td style="cursor: hand" onclick="DetalleCheques">
                                        Detalle
                                        <br />
                                        Cheques
                                    </td>
                                    <td>
                                        Crédito Disponible<br />
                                        <input name="txtCreditoDisponible" value="<%=FormatNumber(nCredito - nCreditoUsado, 2)%>"
                                            style="text-align: right; height: 22px; width: 150px" disabled />
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                    <tr height="20" style="display: none">
                        <td>
                            <table>
                                <tr id="trUsuario">
                                    <td>
                                        Nombre de usuario<br />
                                        <input name="txtUsuario" id="txtUsuario" value="<%=sUsuario%>" onkeypress="IngresarTexto(3)"
                                            maxlength="12" style="width: 190px" />
                                    </td>
                                    <td>
                                        Password<br />
                                        <input type="password" name="txtPassword" id="txtPassword" value="<%=sPassword%>"
                                            onkeypress="IngresarTexto(3)" maxlength="10" style="width: 135px" />
                                    </td>
                                    <td>
                                        Confirmar Password<br />
                                        <input type="password" name="txtConfPassword" id="txtConfPassword" onkeypress="IngresarTexto(3)"
                                            maxlength="10" style="width: 135px" />
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                    <tr id="trContacto" height="20" style="display: none">
                        <td>
                            <table width="100%">
                                <tr height="15">
                                    <td colspan="3" class="titulo">
                                        Antecedentes del Contacto
                                    </td>
                                </tr>
                                <tr>
                                    <td>
                                        Nombres
                                        <br />
                                        <input name="txtNombresContacto" id="txtNombresContacto" size="25" onkeypress="IngresarTexto(2)"
                                            maxlength="20" />
                                    </td>
                                    <td>
                                        Apellido Paterno<br />
                                        <input name="txtApellidoPContacto" id="txtApellidoPContacto" onkeypress="IngresarTexto(2)"
                                            maxlength="15" />
                                    </td>
                                    <td>
                                        Apellido Materno<br />
                                        <input name="txtApellidoMContacto" id="txtApellidoMContacto" onkeypress="IngresarTexto(2)"
                                            maxlength="15" />
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
        <tr>
            <td>
                <table>
                    <tr id="trAntecedentesLaborales" style="display: none">
                        <td class="titulo">
                            Antecedentes Laborales
                        </td>
                    </tr>
                    <tr id="trLaborales1" height="20" style="display: none">
                        <td>
                            <table>
                                <tr>
                                    <td colspan="3">
                                        Empresa<br />
                                        <input name="txtNombreEmpresa" size="50" onkeypress="IngresarTexto(3)" maxlength="40" />
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                    <tr id="trLaborales2" height="20" style="display: none">
                        <td>
                            <table>
                                <tr>
                                    <td>
                                        Dirección<br />
                                        <input name="txtDireccionLaboral" size="30" onkeypress="IngresarTexto(3)" maxlength="40" />
                                    </td>
                                    <td style="display: none">
                                        Pais<br />
                                        <select name="cbxPaisLaboral" style="width: 150px">
                                            <% CargarUbicacion 1, "", sPaisL  %>
                                        </select>
                                    </td>
                                    <td>
                                        Ciudad<br />
                                        <select name="cbxCiudadLaboral" style="width: 150px">
                                            <% CargarCiudadesPais sPais, sCiudadL %>
                                        </select>
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                    <tr id="trLaborales3" height="20" style="display: none">
                        <td>
                            <table>
                                <tr>
                                    <td>
                                        Comuna<br />
                                        <select name="cbxComunaLaboral" style="width: 150px">
                                            <% 
									If sPais = "CL" Then
										CargarComunaCiudad  sCiudadL, sComunaL
									End If
                                            %>
                                        </select>
                                    </td>
                                    <td>
                                        Teléfono<br />
                                        <input name="txtPaisFonoLaboral" style="width: 40px" />
                                        <input name="txtAreaFonoLaboral" style="width: 40px" />
                                        <input name="txtFonoLaboral" style="width: 100px" maxlength="10" />
                                    </td>
                                    <td>
                                        Fax<br />
                                        <input name="txtPaisFaxLaboral" style="width: 40px" />
                                        <input name="txtAreaFaxLaboral" style="width: 40px" />
                                        <input name="txtFaxLaboral" style="width: 100px" maxlength="10" />
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
        <tr height="10">
            <td>
            </td>
        </tr>
        <!-- *********************************** Fin Nuevo ********************************** -->
        <tr>
            <td>
                <table class="Borde" cellspacing="0" cellpadding="0" style="font-family: Verdana;
                    font-size: 10pt; position: relative; top: 0px; sborder: 1;">
                    <tr height="22">
                        <td id="tdCertificacion" style="background-color: #ffffcc; #e1e1e1; cursor: hand">
                            <b>&nbsp;&nbsp;Certificación&nbsp;&nbsp;</b>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
        <tr height="1" style="background-color: silver">
            <td colspan="3">
            </td>
        </tr>
        <tr height="10">
            <td>
            </td>
        </tr>
        <tr>
            <td colspan="3" align="left">
                <table cellpadding="6">
                    <tr>
                        <td>
                            <input type="hidden" name="ChkInfCom" value="<%=nChkInfCom%>" />
                            <input type="hidden" name="ChkInfTri" value="<%=nChkInfTri%>" />
                            <input type="checkbox" name="chkInformacionComercial" <%=sChkInfCom%> />Información
                            Comercial<br />
                            <input type="checkbox" name="chkInformacionTributaria" <%=sChkInfTri%> />Información
                            Tributaria
                        </td>
                        <td>
                            <input type="hidden" name="ChkPon" value="<%=nChkPon%>" />
                            <input type="hidden" name="ChkAcrCom" value="<%=nChkAcrCom%>" />
                            <input type="checkbox" name="chkPonderacion" <%=sChkPon%> />Ponderación<br />
                            <input type="checkbox" name="chkAcreditacionComercial" <%=sChkAcrCom%> />Acreditación
                            Comercial
                        </td>
                        <td>
                            <input type="hidden" name="ChkCarCli" value="<%=nChkCarCli%>" />
                            <input type="checkbox" name="chkCarpetaCliente" <%=sChkCarCli%> />Carpeta Cliente<br />
                            &nbsp;
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
        <tr height="10">
            <td>
            </td>
        </tr>
        <tr height="1" style="background-color: silver">
            <td colspan="3">
            </td>
        </tr>
        <tr height="20">
            <td>
            </td>
        </tr>
        <tr>
            <td colspan="3">
                <table class="Borde" cellspacing="0" cellpadding="0" style="font-family: Verdana;
                    font-size: 10pt; position: relative; top: 0px; background-color: silver;">
                    <tr height="22">
                        <td id="tdDocumento" style="background-color: #ffffcc; cursor: hand">
                            <b>&nbsp;&nbsp;Documentos&nbsp;&nbsp;</b>
                        </td>
                        <td width="690">
                        &nbsp;&nbsp;&nbsp;&nbsp;
                        <input type="button" name="cmdActualizarDocumentos" value="Actualizar" />
                        
                        </td>
                    </tr>
                    
                </table>
                <table cellspacing="1" cellpadding="4" width="800" id="tbDocumento" style="color: #505050;
                    font-family: Verdana; font-size: 10px; position: relative; top: 0px; border: 1;
                    background-color: silver; display: block">
                    <tr style="height: 20px" align="center">
                        <td style="background-color: #e1e1e1" width="150">
                            <b>Nombre</b>
                        </td>
                        <td style="background-color: #e1e1e1" width="50">
                            <b>Número</b>
                        </td>
                        <td style="background-color: #e1e1e1" width="150">
                            <b>Usuario</b>
                        </td>
                        <td style="background-color: #e1e1e1" width="80">
                            <b>Fecha</b>
                        </td>
                        <td style="background-color: #e1e1e1" width="80">
                            <b>Hora</b>
                        </td>
                        <td style="background-color: #e1e1e1" width="400">
                            <b>Nombre Archivo</b>
                        </td>
                        <td style="background-color: #e1e1e1" width="300">
                            <b>Origen</b>
                        </td>
                        <td style="background-color: #e1e1e1" width="100">
                            <b>Estado</b>
                        </td>
                        <td style="background-color: #e1e1e1" width="50">
                            <b><font color="red">Autorizar</font></b>
                        </td>
                        <td style="background-color: #e1e1e1" width="60">
                            <b><font color="red">Deshabilitar</font></b>
                        </td>
                    </tr>
                    <%
						if not rsDocumento is nothing then
						Dim sNombreArchivo
						Dim sNumeroDocumento
						Do Until rsDocumento.EOF
						
					    if cint(rsDocumento("id_tipo_estado_documento")) = 4 then 
                    %>
                    <tr bgcolor="white" style="color: Gray; sbackground-color: white; #f1f1f1; height: 16px;
                        font-style: italic; cursor: hand" onmouseover="javascript:this.bgColor='#f4f4f4'; "
                        onmouseout="javascript:this.bgColor='white'">
                        <%
					    else
                        %>
                        <tr bgcolor="white" style="color: blue; sbackground-color: white; #f1f1f1; height: 16px;
                            cursor: hand" onmouseover="javascript:this.bgColor='#f4f4f4'; " onmouseout="javascript:this.bgColor='white'">
                            <%
					    end if
                            %>
                            <%									
						    
						    if CCur(rsDocumento("id_archivo")) = 0 then 'APPL-23324 MS 10-05-2016
						    ' JFMG 18-04-2012 desde ahora las imagenes se dividen en dos grupos, las diarias y las historicas							    
						        sNombreArchivo = rsDocumento("fecha")						    
						        if cdate(rsDocumento("fecha")) = cdate(date) then
						            sNombreArchivo = session("SitioImagenesClienteDiarias")
						        else
						            sNombreArchivo = session("SitioImagenesClienteHistoricas")
						        end if
    						    
						        if isnull(rsDocumento("nombre_documento")) then								        
						            if trim(rsDocumento("numero")) = "" then
						                sNumeroDocumento = "0"
						            else 
						                sNumeroDocumento = trim(rsDocumento("numero"))
						            end if
						            sNombreArchivo = trim(sNombreArchivo) & trim(rsDocumento("identificacioncliente")) & "_" & trim(rsDocumento("tipo")) & "_" & trim(sNumeroDocumento) & ".jpg"
						        else
						            sNombreArchivo = trim(sNombreArchivo) & trim(rsDocumento("nombre_documento"))
						        end if
						    ' FIN JFMG 18-04-2012
						    else
						        sNombreArchivo = Session("URLVisorArchivo") & "?IdDocumento=" & clng(rsDocumento("id_archivo"))'APPL-7296_MS_11-08-2014
						     
						    end if
						    
								    
                            %>
                            <a href="<%=sNombreArchivo%>" target="_blank">
                                <td>
                                    <%=rsDocumento("descripcion_Documento") %>
                                </td>
                                <td>
                                    <%=rsDocumento("numero") %>
                                </td>
                                <td>
                                    <%=rsDocumento("usuario") %>
                                </td>
                                <td>
                                    <%=rsDocumento("fecha") %>
                                </td>
                                <td>
                                    <%=rsDocumento("hora") %>
                                </td>
                                <td>
                                    <%=rsDocumento("nombre_documento") %>
                                </td>
                                <td>
                                    <%=rsDocumento("origen") %>
                                </td>
                                <td>
                                    <!-- MS 10-03-2014 -->
                                    <%=ucase(rsDocumento("descripcionEstado")) %>
                                    <!-- FIN MS 07-03-2014 -->
                                </td>
                            </a>
                            <td>
                                <div style="width: 55px;">
                                    <!-- JFMG 02-02-2010 se cambia por el llamado a un procedimiento para poder preguntar si está seguro de eliminar-->
                                    <table>
                                    <%
                                 if cint(rsDocumento("id_tipo_estado_documento")) = 1 then 
                                    %>
                                    
                                        <tr>
                                            <td width="30px">
                                                &nbsp;
                                            </td>
                                            <td width="30px">
                                                <img src="../images/button_cancel.png" border="0" onclick="RechazarDocumento <%=rsDocumento("id_documento")%>, <%=rsDocumento("tipo")%>"
                                                    alt="Presione aquí para RECHAZAR este documento" />
                                            </td>
                                        </tr>
                                    <%
                                 elseif cint(rsDocumento("id_tipo_estado_documento")) = 2 then 
                                    %>
                                        <tr>
                                            <td width="30px">
                                                <img src="../images/button_ok.png" border="0" onclick="AutorizarDocumento <%=rsDocumento("id_documento")%>, <%=rsDocumento("tipo")%>"
                                                    alt="Presione aquí para AUTORIZAR este documento" />
                                            </td>
                                            <td width="30px">
                                                &nbsp;
                                            </td>
                                    <%
                                 elseif cint(rsDocumento("id_tipo_estado_documento")) = 3 then 
                                    %>
                                    
                                        <tr>
                                            <td width="30px">
                                                <img src="../images/button_ok.png" border="0" onclick="AutorizarDocumento <%=rsDocumento("id_documento")%>, <%=rsDocumento("tipo")%>"
                                                    alt="Presione aquí para AUTORIZAR este documento" />
                                            </td>
                                            <td width="30px">
                                                <img src="../images/button_cancel.png" border="0" onclick="RechazarDocumento <%=rsDocumento("id_documento")%>, <%=rsDocumento("tipo")%>"
                                                    alt="Presione aquí para RECHAZAR este documento" />
                                            </td>
                                        </tr>
                                   
                                        <%
								    else
								     %>
								     <tr>
								        <td></td>
								     </tr>
								        <%
								    end if
                                        %>
                                   </table>
                                        <!-- **************************** FIN JFMG 02-02-2010 ***************************-->
                                </div>
                            </td>
                            <td>
                                <div style="width: 75px">
                                    <!-- JFMG 02-02-2010 se cambia por el llamado a un procedimiento para poder preguntar si está seguro de eliminar-->
                                    <%
                                 if cint(rsDocumento("id_tipo_estado_documento"))= 1 or cint(rsDocumento("id_tipo_estado_documento"))= 2 or cint(rsDocumento("id_tipo_estado_documento"))= 3   then 
                                    %>
                                    <label id="lblDeshabilitar" onclick="EliminarDocumento <%=rsDocumento("id_documento")%>, <%=rsDocumento("tipo")%>">
                                        Deshabilitar</label>
                                    <%
								 end if
                                    %>
                                    <!-- **************************** FIN JFMG 02-02-2010 ***************************-->
                                </div>
                            </td>
                        </tr>
                        <% 
					    
							rsDocumento.MoveNext
						Loop 
						end if
						Set rsDocumento = Nothing
                        %>
                </table>
            </td>
        </tr>
        <tr height="10">
            <td>
            </td>
        </tr>
        <tr>
            <td>
                <!--Jonathan Miranda G. 14-03-2007-->
                <table class="Borde" cellspacing="0" cellpadding="0" style="font-family: Verdana;
                    font-size: 10pt; position: relative; top: 0px; background-color: silver;">
                    <tr height="22">
                        <td id="td2" style="background-color: #ffffcc; cursor: hand">
                            <b>&nbsp;&nbsp;Perfil Operacional&nbsp;&nbsp;</b>
                        </td>
                    </tr>
                </table>
                <table cellspacing="1" cellpadding="4" width="100%" id="Table2" style="color: #505050;
                    font-family: Verdana; font-size: 10px; position: relative; top: 0px; border: 1;">
                    <tr height="1" style="background-color: silver">
                        <td colspan="3">
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <table width="50%">
                                <tr>
                                    <td>
                                        Nivel de Riesgo<br />
                                        <select name="cbxRiesgo" <%=sHabilitado%>>
                                            <%	CargarEstado "NIVELRIESGO", iRiesgo %>
                                        </select>
                                    </td>
                                    <td>
                                        PEP<br />
                                        <select name="cbxPerfilPEP" <%=sHabilitado%>>
                                            <%	CargarPerfil "1", iPerfilPEP %>
                                        </select>
                                    </td>
                                    <td>
                                        Zona<br />
                                        <select name="cbxPerfilZona" <%=sHabilitado%>>
                                            <%	CargarPerfil "2", iPerfilZona %>
                                        </select>
                                    </td>
                                </tr>
                                <tr>
                                    <td>
                                        Residencia<br />
                                        <select name="cbxPerfilRS" <%=sHabilitado%>>
                                            <%	CargarPerfil "3", iPerfilRS %>
                                        </select>
                                    </td>
                                    <td style="display: none">
                                        Actividad<br />
                                        <select name="cbxPerfilACT" <%=sHabilitado%>>
                                            <%	CargarPerfil "4", iPerfilACT %>
                                        </select>
                                    </td>
                                    <td>
                                        Industria MBS<br />
                                        <select name="cbxPerfilIndustria" <%=sHabilitado%>>
                                            <%	CargarPerfil "5", iPerfilIndustria %>
                                        </select>
                                    </td>
                                    <td>
                                        Perfil del Cliente<br />
                                        <select name="cbxPerfilCliente" <%=sHabilitado%>>
                                            <%	CargarPerfil "6", iPerfilCliente %>
                                        </select>
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                </table>
                <!------------------  Fin  ------------------------------>
            </td>
        </tr>
        <tr height="10">
            <td>
            </td>
        </tr>
        <tr>
            <td>
                <!--Jonathan Miranda G. 13-06-2011-->
                <table class="Borde" cellspacing="0" cellpadding="0" style="font-family: Verdana;
                    font-size: 10pt; position: relative; top: 0px; background-color: silver;">
                    <tr height="22">
                        <td id="td1" style="background-color: #ffffcc; cursor: hand">
                            <b>&nbsp;&nbsp;Perfil Transaccional&nbsp;&nbsp;</b>
                        </td>
                    </tr>
                </table>
                <table cellspacing="1" cellpadding="4" width="100%" id="Table1" style="color: #505050;
                    font-family: Verdana; font-size: 10px; position: relative; top: 0px; border: 1;">
                    <tr height="1" style="background-color: silver">
                        <td colspan="3">
                        </td>
                    </tr>
                    <tr height="80">
                        <td colspan="3">
                            <table>
                                <tr>
                                    <td align="left" colspan="3">
                                        Actividad Econ&oacute;mica<br />
                                        <input type="radio" name="optActEconomicaSi" <%=sHabilitado%> <%=sActividadSi%> />Si
                                        <input type="radio" name="optActEconomicaNo" <%=sHabilitado%> <%=sActividadNo%> />No <br />
                                    </td>
                                </tr>
                                <tr>
                                    <td align="left" colspan="3">
                                        <select name="cbxActividadEconomicaOculta" style="width: 450px; font-size: 10px;
                                            display: none;">
                                            <% CargarActividadEconomica "" %>
                                        </select>
                                        <select name="cbxActividadEconomica" style="width: 450px; font-size: 10px;" <%=sHabilitado%>>
                                            <% CargarActividadEconomica "" %>
                                        </select><span>*</span><input type="button" name="cmdAgregarActividad" value="Agregar" <%=sHabilitado%> />
                                        <table width="100%">
                                            <tr>
                                                <td>
                                                    <table id="tblActividadEconomica" cellspacing="1" cellpadding="4" style="color: #505050;
                                                        font-family: Verdana; font-size: 10px; border: 1;">
                                                        <tr style="height: 20px" align="center">
                                                            <td colspan="3" style="background-color: #e1e1e1; font-size: 12px;" width="90%">
                                                                <b>Descripci&oacute;n</b>
                                                            </td>
                                                        </tr>
                                                        <%
						            if not rsActividadEconomica is nothing and request("ActividadEconomica") = "" then
						                Do Until rsActividadEconomica.EOF
                                                If rsActividadEconomica("activa") Then
                                                        %>
                                                        <tr bgcolor="white" style="color: blue; sbackground-color: white; #f1f1f1; height: 16px;
                                                            cursor: hand" onmouseover="javascript:this.bgColor='#f4f4f4'; " onmouseout="javascript:this.bgColor='white'">
                                                            <td style="display: none;">
                                                                <%=rsActividadEconomica("idActividadEconomica") %>
                                                            </td>
                                                            <td>
                                                                <%=UCASE(rsActividadEconomica("DescripcionActividadEconomica"))%>
                                                            </td>
                                                            <td>
                                                                <img src="../images/elimsup.jpg" border="0" onclick="EliminarActividad <%=rsActividadEconomica("idActividadEconomica") %>"
                                                                    alt="Presione aquí para eliminar" <%=sHabilitado%> />
                                                            </td>
                                                        </tr>
                                                        <%
							                    sActividadEconomica = rsActividadEconomica("idActividadEconomica") & ";" & sActividadEconomica
							                    Else
							                        sActividadEconomicaInactiva = rsActividadEconomica("idActividadEconomica") & ";" & sActividadEconomicaInactiva
							                    End If
							                    rsActividadEconomica.MoveNext							                    
						                Loop
						            end if
						            Set rsActividadEconomica = Nothing
                                                        %>
                                                    </table>
                                                    <input type="hidden" name="txtActividadEconomica" value="<%=sActividadEconomica%>" />
                                                    <input type="hidden" name="txtActividadEconomicaInactiva" value="<%=sActividadEconomicaInactiva%>" />
                                                </td>
                                            </tr>
                                        </table>
                                    </td>
                                </tr>
                                <tr>
                                    <td colspan="2">
                                        Prop&oacute;sito de las Transacciones<br />
                                        <input type="text" name="txtPropositoTransacciones" style="width: 547px; height: 50px;"
                                            value="<%=sPropositoTransacciones%>" <%=sHabilitado%> onkeypress="IngresarTexto(3)"
                                            maxlength="500" />
                                    </td>
                                </tr>
                                <tr>
                                    <td colspan="2">
                                        Ref. Comercial / Bancaria<br />
                                        <input type="text" name="txtReferenciaBancaria" style="width: 547px" value="<%=sReferenciaBancaria%>"
                                            <%=sHabilitado%> onkeypress="IngresarTexto(3)" maxlength="200" />
                                    </td>
                                </tr>
                                <tr>
                                    <td colspan="2">
                                        <b>Transacciones por Mes</b>
                                    </td>
                                </tr>
                                <tr>
                                    <td>
                                        Informaci&oacute;n Cliente<br />
                                        <table>
                                            <tr>
                                                <td>
                                                    Cantidad<br />
                                                    <input type="text" name="txtCantidadTransaccionesMesCLIENTE" style="width: 40px;
                                                        text-align: right;" value="<%=sCantidadTransaccionesMesCLIENTE%>" <%=sHabilitado%>
                                                        onkeypress="IngresarTexto(1)" maxlength="3" />
                                                </td>
                                                <td>
                                                    Monto USD<br />
                                                    <input type="text" name="txtMontoTransaccionesMesCLIENTE" style="width: 80px; text-align: right;"
                                                        value="<%=sMontoTransaccionesMesCLIENTE%>" <%=sHabilitado%> onkeypress="IngresarTexto(1)"
                                                        maxlength="18" />
                                                </td>
                                            </tr>
                                        </table>
                                    </td>
                                    <td>
                                        Informaci&oacute;n AFEX<br />
                                        <table>
                                            <tr>
                                                <td>
                                                    Cantidad<br />
                                                    <input type="text" name="txtCantidadTransaccionesMesAFEX" style="width: 40px; text-align: right;"
                                                        value="<%=sCantidadTransaccionesMesAFEX%>" disabled onkeypress="IngresarTexto(1)"
                                                        maxlength="3" />
                                                </td>
                                                <td>
                                                    Monto USD<br />
                                                    <input type="text" name="txtMontoTransaccionesMesAFEX" style="width: 80px; text-align: right;"
                                                        value="<%=sMontoTransaccionesMesAFEX%>" disabled onkeypress="IngresarTexto(1)"
                                                        maxlength="18" />
                                                </td>
                                            </tr>
                                        </table>
                                    </td>
                                </tr>
                                <tr>
                                    <td>
                                        <b>Valores</b>
                                    </td>
                                </tr>
                                <tr>
                                    <td style="vertical-align: top;">
                                        Ingreso Anual<br />
                                        <input type="text" name="txtIngresoAnual" style="width: 80px; text-align: right;"
                                            value="<%=sIngresoAnual%>" <%=sHabilitado%> onkeypress="IngresarTexto(1)" maxlength="18" />
                                    </td>
                                </tr>
                                <tr>
                                    <td valign="top">
                                        Monto Patrimonio USD<br />
                                        <input type="text" name="txtMontoPatrimonio" style="width: 80px; text-align: right;"
                                            value="<%=sMontoPatrimonio%>" <%=sHabilitado%> onkeypress="IngresarTexto(1)"
                                            maxlength="18" />
                                    </td>
                                    <td>
                                        Origen Patrimonio y/o Fondos<br />
                                        <input type="text" name="txtOrigenFondos" style="width: 397px; height: 50px;" value="<%=sOrigenFondos%>"
                                            <%=sHabilitado%> onkeypress="IngresarTexto(3)" maxlength="200" />
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            &nbsp;
                        </td>
                    </tr>
                    <!------------------  Fin  ------------------------------>
                </table>
            </td>
        </tr>
        <tr>
            <td colspan="3">
                <%
                %>
                <table class="Borde" cellspacing="0" cellpadding="0" style="font-family: Verdana;
                    font-size: 10pt; position: relative; top: 0px; background-color: silver;">
                    <tr height="22">
                        <td id="tdHistoria" style="background-color: #ffffcc; #e1e1e1; cursor: hand">
                            <b>&nbsp;&nbsp;Historia&nbsp;&nbsp;</b>
                        </td>
                    </tr>
                </table>
                <table cellspacing="1" cellpadding="1" width="100%" id="tbHistoria" align="center"
                    style="color: #505050; font-family: Verdana; font-size: 10px; position: relative;
                    top: 0px; border: 1; background-color: silver; display: nones">
                    <tr style="height: 20px" align="center">
                        <td style="background-color: #e1e1e1" width="10%">
                            <b>Fecha</b>
                        </td>
                        <!--<td style="background-color: #e1e1e1" WIDTH="10%">
								<b>Hora</b>
							</td>-->
                        <td style="background-color: #e1e1e1" width="80%">
                            <b>Detalle</b>
                        </td>
                    </tr>
                    <%
							Dim sDetalle, sHora, sBGColor
							Do Until rsHistoria.EOF
								'sHora = Right("000000" & rsHistoria("hora"), 6)
								'sHora = Left(sHora, 2) & ":" & Mid(sHora, 3, 2) & ":" & Right(sHora, 2)
								Select Case rsHistoria("tipo")
									Case 1	'Informacion
										sBGColor = "#ddeeff"
										sColor = "#ddeeff"
									Case 2	'Advertencia
										sBGColor = "#ffffee"
										sColor = "#ffff00"
									Case 3	'Peligro
										sBGColor = "#ffddee"
										sColor = "#ffddee"
									Case Else
										sBGColor = "white"
										sColor = "#ff0000"
									End Select
									sHora= rsHistoria("hora")
                    %>
                    <tr style="background-color: <%=sBGColor%>; color: black; <%=sColor%>; height: 16px">
                        <td>
                            <%=rsHistoria("fecha") %>
                        </td>
                        <!--<td><%=sHora%></td>-->
                        <td>
                            <%=rsHistoria("descripcion") & " (" & rsHistoria("NombreUsuario") & ")"%>
                        </td>
                    </tr>
                    <% 
								rsHistoria.MoveNext
							Loop 
							Set rsHistoria = Nothing
							Set rsCliente = Nothing
                    %>
                </table>
                <%
                 if nHistoriaCompleta = 0 then 
                %>
    
                <a href="http:DetalleCliente.asp?cc=<%=Request("cc")%>&nc=<%=Request("nc")%>&rt=<%=Request("rt")%>&hist=1" >Ver más...</a>
                <%
                 end if
                %>
            </td>
        </tr>
        <tr height="10">
            <td>
            </td>
        </tr>
    </table>
    <input name="txtPregunta" type="hidden" value="<%=sPregunta%>" />
    <input name="txtRespuesta" type="hidden" value="<%=sRespuesta%>" />
    <input name="txtContacto1" type="hidden" value="<%=sContacto1%>" />
    <input name="txtPorcentageContacto1" type="hidden" value="<%=sPorcentageContacto1%>" />
    <input name="txtContacto2" type="hidden" value="<%=sContacto2%>" />
    <input name="txtPorcentageContacto2" type="hidden" value="<%=sPorcentageContacto2%>" />
    <input name="txtFechaActivacionComision" type="hidden" value="<%=sFechaActivacionComision%>" />
    <input name="txtSucursalSolicitante" type="hidden" value="" />
    </form>
    <!--#INCLUDE virtual="/Compartido/Rutinas.htm" -->
</body>
<head>
    <meta http-equiv="CACHE-CONTROL" content="NO-CACHE, must-revalidate" />
    <meta http-equiv="Pragma" content="no-cache" />
</head>
</html>