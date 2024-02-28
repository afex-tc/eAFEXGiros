<script language="JavaScript" src="../script/funciones.js"> </script>
<script language="JavaScript">
    /* ***********************
    Validador de RUT
    29-07-2015
    ************************ */
    function ValidaRut(textoRut) {
        if (textoRut.length > 0) {
            return (Rut(textoRut));
        }
    }

    function revisarDigito(dvr) {
        dv = dvr + ""
        if (dv != '0' && dv != '1' && dv != '2' && dv != '3' && dv != '4' && dv != '5' && dv != '6' && dv != '7' && dv != '8' && dv != '9' && dv != 'k' && dv != 'K') {
            alert("Debe ingresar un digito verificador valido");
            frmCliente.txtRutCautela.focus();
            frmCliente.txtRutCautela.select();
            return false;
        }
        return true;
    }

    function revisarDigito2(crut) {
        largo = crut.length;
        if (largo < 2) {
            alert("Debe ingresar el rut completo")
            frmCliente.txtRutCautela.focus();
            frmCliente.txtRutCautela.select();
            return false;
        }
        if (largo > 2)
            rut = crut.substring(0, largo - 1);
        else
            rut = crut.charAt(0);
        dv = crut.charAt(largo - 1);
        revisarDigito(dv);

        if (rut == null || dv == null)
            return 0

        var dvr = '0'
        suma = 0
        mul = 2

        for (i = rut.length - 1; i >= 0; i--) {
            suma = suma + rut.charAt(i) * mul
            if (mul == 7)
                mul = 2
            else
                mul++
        }
        res = suma % 11
        if (res == 1)
            dvr = 'k'
        else if (res == 0)
            dvr = '0'
        else {
            dvi = 11 - res
            dvr = dvi + ""
        }
        if (dvr != dv.toLowerCase()) {
            alert("EL rut es incorrecto")
            frmCliente.txtRutCautela.focus();
            frmCliente.txtRutCautela.select();
            return false
        }

        return true
    }

    function Rut(texto) {
        var tmpstr = "";
        for (i = 0; i < texto.length; i++)
            if (texto.charAt(i) != ' ' && texto.charAt(i) != '.' && texto.charAt(i) != '-')
            tmpstr = tmpstr + texto.charAt(i);
        texto = tmpstr;
        largo = texto.length;

        if (largo < 2) {
            alert("Debe ingresar el rut completo")
            frmCliente.txtRutCautela.focus();
            frmCliente.txtRutCautela.select();
            return false;
        }

        for (i = 0; i < largo; i++) {
            if (texto.charAt(i) != "0" && texto.charAt(i) != "1" && texto.charAt(i) != "2" && texto.charAt(i) != "3" && texto.charAt(i) != "4" && texto.charAt(i) != "5" && texto.charAt(i) != "6" && texto.charAt(i) != "7" && texto.charAt(i) != "8" && texto.charAt(i) != "9" && texto.charAt(i) != "k" && texto.charAt(i) != "K") {
                alert("El valor ingresado no corresponde a un R.U.T valido");
                frmCliente.txtRutCautela.focus();
                frmCliente.txtRutCautela.select();
                return false;
            }
        }

        var invertido = "";
        for (i = (largo - 1), j = 0; i >= 0; i--, j++)
            invertido = invertido + texto.charAt(i);
        var dtexto = "";
        dtexto = dtexto + invertido.charAt(0);
        dtexto = dtexto + '-';
        cnt = 0;

        for (i = 1, j = 2; i < largo; i++, j++) {
            //alert("i=[" + i + "] j=[" + j +"]" );		
            if (cnt == 3) {
                dtexto = dtexto + '.';
                j++;
                dtexto = dtexto + invertido.charAt(i);
                cnt = 1;
            }
            else {
                dtexto = dtexto + invertido.charAt(i);
                cnt++;
            }
        }

        invertido = "";
        for (i = (dtexto.length - 1), j = 0; i >= 0; i--, j++)
            invertido = invertido + dtexto.charAt(i);
            
        frmCliente.txtRutCautela.value = invertido.toUpperCase()
        //window.document.form1.rut.value = invertido.toUpperCase()

        if (revisarDigito2(texto))
            return true;

        return false;
    }
</script>
<%@  language="VBScript" %>
<!--#INCLUDE virtual="/Compartido/Rutinas.asp" -->
<!--#INCLUDE virtual="/sucursal/Rutinas.asp" -->
<%
	Response.Expires = 0
	' JFMG 29-06-2010 se agrega para el ingreso de clientes desde las sucursales	
	Session("TipoOrigenLlamada") = 0
	IF request("TipoOrigenLlamada") = "1" THEN
		' se llama desde el menú de una sucursal
		Session("CodigoCliente") = request("SCC")
		Session("CodigoAgente") = request("AGC")
		Session("TipoOrigenLlamada") = 1 ' Sucursal página de giros
		Session("NombreUsuarioOperador") = request("NUO")
		Session("VerClienteCorporativo") = True
	END IF
	' FIN JFMG 29-06-2010
	
	If Session("CodigoCliente") = "" Then
		response.Redirect "../Compartido/TimeOut.htm"
		response.end
	End If
%>
<%
	Dim nTipo
	Dim sTitulo
	Dim sEncabezadoFondo
	Dim sEncabezadoTitulo
	Dim sPais, sCiudad, sComuna, sPaisPasaporte, sPaisPasaporteCautela
	Dim sNacionalidad
	Dim sPaisL, sCiudadL, sComunaL
	Dim nTipoCliente, nCodigo
	Dim Sql
	Dim sRut, sPasaporte
	Dim sMensaje
	Dim nEjecutivo
	Dim sSucursal
	Dim nCodigoBanco
	Dim nRubro
	Dim iRiesgo
	Dim iPerilPEP
	Dim iPerilZona
	Dim iPerilRS
	Dim iPerilACT
	Dim iPerilIndustria
	dim sFechaNacimiento, sContacto1, sPorcentageContacto1, sContacto2, sPorcentageContacto2, sFechaActivacionComision
	dim sSerieIdentificacion, bIDConsultada, bIdValida, sMensajeRegistro, nSexo
	dim nClienteAgencia
			
	' JFMG 04-09-2010 datos para buscar el cliente en otro lugar en caso que no se encuentre en corporativa
	dim bClienteOtroLugar	
	dim sNombresCliente
	dim sApellidoPaternoCliente
	dim sApellidoMaternoCliente
	dim iDdiPaisCliente
	dim iDdiCiudadCliente
	dim sDireccionCliente
	dim sRazonSocialCliente
	dim sCorreoElectronicoCliente
	dim iNumeroTelefonoCliente
	' FIN JFMG 04-09-2010
		
	dim sNivelRiesgo
    dim sDescripcionCautela
	dim nvlR
			
    dim sOcupacion 'INTERNO-9263 MS 19-01-2017
							
	nTipo = cInt(0 & Request("Tipo"))
	nTipoCliente = CInt(0 & Request("TipoCliente"))
	
	nSexo = Request("nSexo")
	
	nClienteAgencia = request("ClienteAgencia")
	
	If nClienteAgencia Then
		nClienteAgencia = "Checked"
	else
		nClienteAgencia = " "
	End if
		
	Select Case nTipo
		Case 0
			'sEncabezadoTitulo = "Nuevo Cliente"
	
		Case 1
			'sEncabezadoTitulo = "Actualización de Datos"
	
	End Select
	
	sSerieIdentificacion = Request.form("txtNumeroSerie")
	bIDConsultada = Request.Form("txtIdConsultada")
	if bIDConsultada = "" then bIDConsultada = "false"
	bIDValida = Request.Form("txtIdValida")
	if bIdValida = "" then bIdValida= "false"	
	
	sNacionalidad = Trim(Request("NC"))
	sPaisPasaporte = Trim(Request("PP"))
	sPais = Trim(Request("Pais"))
	sCiudad = Trim(Request("Ciudad"))
	sComuna = Trim(Request("Comuna"))
	sPaisL = Trim(Request("PaisL"))
	sCiudadL = Trim(Request("CiudadL"))
	sComunaL = Trim(Request("ComunaL"))
	sSucursal = Trim(Request("sc"))
	nEjecutivo = cInt(0 & Request("ce"))
	nCodigoBanco = cInt(0 & Request("cb"))
	nRubro = cInt(0 & Request("rb"))	
	
	iRiesgo = request.form("cbxRiesgo")
	iPerfilPEP = request.form("cbxPerfilPEP")
	iPerfilZona = request.form("cbxPerfilZona")
	iPerfilRS = request.form("cbxPerfilRS")
	iPerfilACT = request.form("cbxPerfilACT")
	iPerfilIndustria = request.form("cbxPerfilIndustria")
	sFechaNacimiento = request.form("txtFechaNacimiento")
	sContacto1 = request.form("txtcontacto1")
	sPorcentageContacto1 = request.form("txtporcentagecontacto1")
	sContacto2 = request.form("txtcontacto2")
	sPorcentageContacto2 = request.form("txtporcentagecontacto2")
	sFechaActivacionComision = Request.Form("txtFechaActivacionComision")
	
	If sNacionalidad = "" Then sNacionalidad = "CL"
	If sPaisPasaporte = "" Then sPaisPasaporte = ""
	If sPais = "" Then sPais = "CL"
	If sCiudad = "" Then sCiudad = "SCL"
	If sComuna = "" Then sComuna = "STG"
	If sPaisL = "" Then sPaisL = sPais
	If sCiudadL = "" Then sCiudadL = sCiudad
	If sComunaL = "" Then sComunaL = sComuna	

	if iRiesgo = empty then iRiesgo = 1
	if iPerfilPEP = empty then iPerfilPEP = 6
	if iPerfilZona = empty then iPerfilZona = 3
	if iPerfilRS = empty then iPerfilRS = 4
	if iPerfilACT = empty then iPerfilACT = 0
	if iPerfilIndustria = empty then iPerfilIndustria = 15
	
	sNivelRiesgo =  Trim(request("NivelRiesgo"))
	sRut = Trim(request("Rut"))
	sPasaporte = Trim(request("Pasaporte"))
	
	'AMP 31-07-2015
	On Error Resume Next
    
    If Not IsNull(sRut) And sRut <> "" And sRut <> "null" Then
	    nvlR = BuscaNivelRiesgo (sRut)
	    If nvlR <> 0 Then
	        If nvlR = 1 Then '-----> NORMAL
	            sNivelRiesgo = "NORMAL"
	            nCodigo = BuscarIdentificacion (sRut, sPasaporte, sMensaje)
	            If nCodigo <> 0 Then
	                response.Redirect "http:DetalleCliente.asp?cc=" & nCodigo
	            Else
	                nCodigo = BuscaClienteCautela ("Rut", sRut)
		            If nCodigo <> 0 Then
	                    response.Redirect "http:DetalleClienteCautela.asp?cc=" & nCodigo
	                End If
	            End If
	        Else '-----> CAUTELA
	            sNivelRiesgo = "CAUTELA"
	            nCodigo = BuscaClienteCautela ("Rut", sRut)
	            If nCodigo <> 0 Then
		            response.Redirect "http:DetalleClienteCautela.asp?cc=" & nCodigo
		        else
		            nCodigo = BuscarIdentificacion (sRut, sPasaporte, sMensaje)
	                If nCodigo <> 0 Then
		                response.Redirect "http:DetalleCliente.asp?cc=" & nCodigo
		            End If
		        End If
	        End If
        Else 'Cuando el cliente fue ingresado sin nivel de riesgo mostrara el formulario normal
            nCodigo = BuscarIdentificacion (sRut, sPasaporte, sMensaje)
	        If nCodigo <> 0 Then
	            response.Redirect "http:DetalleCliente.asp?cc=" & nCodigo
	        End If
	    End If
	    	
		' JFMG 04-09-2010 el cliente no se encontró en corporativa, se busca en otro lugar
		if Request.Form("txtNombres") = "" and Request.Form("txtRazonSocial") = "" then
			bClienteOtroLugar = BuscarIdentificacionOtroLugar (sRut, sPasaporte, sMensaje)
		end if
		' FIN JFMG 04-09-2010
	End If
    If Not IsNull(sPasaporte) And sPasaporte <> "" And sPasaporte <> "null" Then
        nCodigo = BuscaClienteCautela ("Pasaporte", sPasaporte)
	    If nCodigo <> 0 Then
		    response.Redirect "http:DetalleClienteCautela.asp?cc=" & nCodigo
		else
		    nCodigo = BuscarIdentificacion (sRut, sPasaporte, sMensaje)
	        If nCodigo <> 0 Then
		        response.Redirect "http:DetalleCliente.asp?cc=" & nCodigo
		    End If
		End If
		
		' JFMG 04-09-2010 el cliente no se encontró en corporativa, se busca en otro lugar
		if Request.Form("txtNombresCautela") = "" then
			bClienteOtroLugar = BuscarIdentificacionOtroLugar (sRut, sPasaporte, sMensaje)
		end if
		' FIN JFMG 04-09-2010		
	End If
	If Err.number <> 0 Then
	    response.Redirect "../Compartido/Error.asp?Titulo=Error en Agregar Cliente&Number=" & Err.number & "&Source=" & Err.Source & "&Description=" & Err.description & nvlR 
    End If
    'FIN AMP 31-07-2015
		
	Function BuscarDDI(ByVal Tipo, ByVal Codigo)
		Dim afxCOM
		Dim ddi
		
		BuscarDDI = Empty
		
		Select Case Tipo
			Case 1
				Set afxCOM = Server.CreateObject("AfexCorporativo.Pais")
			
			Case 2
				Set afxCOM = Server.CreateObject("AfexCorporativo.Ciudad")
		End Select		
		ddi = afxCOM.BuscarDDI(Session("afxCnxCorporativa"), Codigo)
		If afxCOM.ErrNumber <> 0 Then 
			Set afxCOM = Nothing
			Exit Function
		End If
		BuscarDDI = ddi
		Set afxCOM = Nothing		
	End Function
	
	'AMP 31-07-2015
	Function BuscaNivelRiesgo(ByVal rutR)
	    Dim rs, sSQL
		sSQL = "Select ISNULL(nivelriesgo, 0) nivelriesgo From Cliente Where rut = '" & ValorRut(rutR) & "'"
		
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open sSQL, Session("afxCnxCorporativa")
		
		If Err.number <> 0 Then
			Set rs = Nothing
			sMensaje = Err.Description		
			Exit Function
		End If
		
		If rs.EOF Then
			Set rs = Nothing
			Exit Function
		End If
		
		BuscaNivelRiesgo = rs("nivelriesgo")
		Set rs = Nothing
	
	End Function
	
	Function BuscaClienteCautela(ByVal tipo, ByVal valor)

        Dim sSQL, rs
        rutCautela = Replace(rutCautela, ".", "")
        rutCautela = Replace(rutCautela, "-", "")
        rutCautela = Replace(rutCautela, " ", "")
        
        sSQL = "EXECUTE [ObtenerCLienteCautela] "
        select case tipo
	        case "Rut"									
		        sSQL = sSQL & "@Rut = '" & ValorRut(valor) & "'"
	        case "Codigo"
		        sSQL = sSQL & "@Codigo = '" & valor & "'"
	        case "Pasaporte"
		       sSQL = sSQL & "@Pasaporte = '" & valor & "'"
        end select

        'sSQL = sSQL & "@Rut = '" & ValorRut(valor) & "'"
        
	    Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open sSQL, Session("afxCnxCorporativa")
		
		If Err.number <> 0 Then
			Set rs = Nothing
			sMensaje = Err.Description		
			Exit Function
		End If
		
		If rs.EOF Then
			Set rs = Nothing
			Exit Function
		End If
		
		BuscaClienteCautela = rs("codigo")
		Set rs = Nothing
    End Function
    'FIN AMP 31-07-2015
		
	Function BuscarIdentificacion(ByVal sRCl, ByVal sPCl, ByRef sMensaje)
		Dim rsClienteCorp, sSQL, sRt
		Dim sDecripcion
		
		BuscarIdentificacion = 0
		
		If sRCl <> Empty Then
			sRt = ValorRut(sRCl)
			sSQL = "Select codigo From Cliente with(nolock) Where rut = '" & sRt & "'"
			sDescripcion = "Rut ya existe"
		End If
		If sPCl <> Empty Then
			sSQL = "Select codigo From Cliente with(nolock) Where pasaporte = '" & sPCl & "'"
			sDescripcion = "Pasaporte ya existe"
		End If
		
		Set rsClienteCorp = Server.CreateObject("ADODB.Recordset")
		rsClienteCorp.Open sSQL, Session("afxCnxCorporativa")
		
		If Err.number <> 0 Then
			Set rsClienteCorp = Nothing
			sMensaje = Err.Description		
			Exit Function
		End If
		
		If rsClienteCorp.EOF Then
			Set rsClienteCorp = Nothing
			'sMensaje = sDescripcion
			Exit Function
		End If
		
		BuscarIdentificacion = rsClienteCorp("codigo")
		Set rsClienteCorp = Nothing
	End Function
	
	Function BuscarIdentificacionOtroLugar(ByVal sRCl, ByVal sPCl, ByRef sMensaje)
		Dim rsClienteOtroLugar, sSQL, sRt
		Dim sDecripcion
		
		BuscarIdentificacionOtroLugar = False
		
		If sRCl <> Empty Then
			sRt = ValorRut(sRCl)
			sSQL = "exec obtenerclienterut '" & sRt & "'"
			sDescripcion = "Rut ya existe en otro lugar"
		End If
		If sPCl <> Empty Then
			sSQL = "exec obtenerclientepasaporte '" & sPCl & "'"
			sDescripcion = "Pasaporte ya existe en otro lugar"
		End If
		
		Set rsClienteOtroLugar = Server.CreateObject("ADODB.Recordset")
		rsClienteOtroLugar.Open sSQL, Session("afxCnxCorporativa")
		
		If Err.number <> 0 Then
			Set rsClienteOtroLugar = Nothing
			sMensaje = Err.Description		
			Exit Function
		End If
		
		If rsClienteOtroLugar.EOF Then
			Set rsClienteOtroLugar = Nothing
			exit Function
		End If
		
		' carga los datos obtenidos
		nTipoCliente = CInt("0" & rsClienteOtroLugar("Tipo"))
	
		nSexo = rsClienteOtroLugar("sexo")
	
		nClienteAgencia = " "
				
        sNacionalidad = trim(rsClienteOtroLugar("CodigoNacionalidad")) 'miki SMC-55 MM 2016-03-18
		sPaisPasaporte = trim(rsClienteOtroLugar("codigo_paispas"))
		sPais = trim(rsClienteOtroLugar("codigo_pais"))
		sCiudad = trim(rsClienteOtroLugar("codigo_ciudad"))
		sComuna = trim(rsClienteOtroLugar("codigo_comuna"))
								
		sFechaNacimiento = rsClienteOtroLugar("Fecha_Nacimiento")
	
		If sNacionalidad = "" Then sNacionalidad = "CL"
		If sPaisPasaporte = "" Then sPaisPasaporte = ""
		If sPais = "" Then sPais = "CL"
		If sCiudad = "" Then sCiudad = "SCL"
		If sComuna = "" Then sComuna = "STG"
	
		if iRiesgo = empty then iRiesgo = 1
		if iPerfilPEP = empty then iPerfilPEP = 6
		if iPerfilZona = empty then iPerfilZona = 3
		if iPerfilRS = empty then iPerfilRS = 4
		if iPerfilACT = empty then iPerfilACT = 0
		if iPerfilIndustria = empty then iPerfilIndustria = 15
		
		if rsClienteOtroLugar("tipo") = "1" then
			sApellidoMaternoCliente = rsClienteOtroLugar("materno")
			sApellidoPaternoCliente = rsClienteOtroLugar("paterno")
			sNombresCliente = rsClienteOtroLugar("nombre")
					
			sRazonSocialCliente = ""
		
		else
			sApellidoMaternoCliente = ""
			sApellidoPaternoCliente = ""
			sNombresCliente = ""
					
			sRazonSocialCliente = rsClienteOtroLugar("nombre_completo")
		end if
		sCorreoElectronicoCliente = rsClienteOtroLugar("correoelectronico")
		sDireccionCliente = rsClienteOtroLugar("Direccion")
				
		iDdiPaisCliente = rsClienteOtroLugar("ddi_pais")
		iDdiCiudadCliente = rsClienteOtroLugar("ddi_area")
		iNumeroTelefonoCliente = rsClienteOtroLugar("telefono")
				
		BuscarIdentificacionOtroLugar = True
		
		Set rsClienteOtroLugar = Nothing
	End Function
		
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
									msgbox smensaje, ,"Registro Civil"
																	
									if sDigitoVerificador = ""	then ' id malo
										ConsularIdentificacion = "false"
									else
										ConsularIdentificacion = "true"
									end if
								case 3	' ERROR id malo
									sMensaje = child.text
									msgbox sMensaje, ,"Registro Civil"
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
	
	Response.Expires = 0
	
%>

<html>
<head>
    <meta name="GENERATOR" content="Microsoft Visual Studio 6.0">
    <link rel="stylesheet" type="text/css" href="http:../Estilos/Principal.css">
    <style type="text/css">
        #txtNombres
        {
            width: 207px;
        }
    </style>
</head>

<script language="VBScript">
<!--
    

    'AMP 18-06-2015 Seleccion de tipo de riesgo
    Sub cbxRiesgo_onChange()
        
		if frmCliente.cbxRiesgo.options(frmCliente.cbxRiesgo.selectedIndex).text = "NORMAL" then
		    divDatosCautela.style.display = "none"
		    divDatosNormal.style.display = ""
		    btnprint.style.display = ""
		    divBotonesNormal.style.display = ""
		    divBotonesCautela.style.display = "none"
            LimpiaDatosNormal

		elseif frmCliente.cbxRiesgo.options(frmCliente.cbxRiesgo.selectedIndex).text = "CAUTELA" then
		    divDatosCautela.style.display = ""
		    divDatosNormal.style.display = "none"
		    btnprint.style.display = "none"
		    divBotonesNormal.style.display = "none"
		    divBotonesCautela.style.display = ""
            LimpiaDatosCautela
		    else
		        divDatosCautela.style.display = "none"
		        divDatosNormal.style.display = "none"
		        btnprint.style.display = ""
		        divBotonesNormal.style.display = "none"
		        divBotonesCautela.style.display = "none"
		End if
	end sub
	
	Sub btn_ValidarRegistroCautela_onClick()
		Dim sRut, sPasaporte, sNumeroRut, sTipo
		
		If (frmCliente.optRutCautela.checked = true and (frmCliente.txtRutCautela.value = Empty or frmCliente.txtNumSerieCautela.value = Empty) ) Then
		    msgbox "Debe ingresar el Rut y número de serie Correctamente"
		    frmCliente.txtRutCautela.focus
		    Exit Sub
		End if	
		
		If (frmCliente.txtPasCautela.value = Empty and frmCliente.optPasCautela.checked = true) Then
		    msgbox "Debe ingresar el número de Pasaporte Correctamente"
		    frmCliente.txtPasCautela.focus
		    Exit Sub
		End if		
		
		If frmCliente.txtRutCautela.value <> "" Then
			sRut = ValidarRut(frmCliente.txtRutCautela.value)
			
			frmCliente.txtRutCautela.value = sRut
		
			sNumeroRut = replace(replace((trim(sRut)),"-",""),".","")
			sNumeroRut = left(sNumeroRut, len(sNumeroRut) - 1)
		
		elseif frmCliente.txtPasCautela.value <> "" Then
			sNumeroRut = frmCliente.txtPasCautela.value
			
		Else
			msgbox "Debe ingresar una Identificación "
			Exit Sub
		End If
		
		msgbox "Registro Validado!"
		
	End Sub
	'FIN AMP 19-06-2015

	Sub FichaCliente
		Dim sParametros, sNombreCliente, sApellidos
		Dim sComunaPersonal, sCiudadPersonal, sPaisPersonal
		Dim sFonoPersonal, sFaxPersonal, sNombreEmpresa
		
		Dim sComunaLaboral, sCiudadLaboral, sPaisLaboral
		Dim sFonoLaboral, sFaxLaboral, sNombreContacto, sApellidoContacto
		
		If frmCliente.txtNombres.value <> "" Then
			sNombreCliente = frmCliente.txtNombres.value
			sApellidos = Trim(Trim(frmCliente.txtApellidoP.value) & " " & _
						 Trim(frmCliente.txtApellidoM.value))
			sNombreEmpresa = frmCliente.txtNombreEmpresa.value
			sNombreContacto = ""
			sApellidoContacto = ""
		Else
			sNombreCliente = frmCliente.txtRazonSocial.value
			sNombreContacto = Trim(frmCliente.txtNombresContacto.value)
			sApellidoContacto = Trim(Trim(frmCliente.txtApellidoPContacto.value) & " " & _
								Trim(frmCliente.txtApellidoMContacto.value))
			sNombreEmpresa = sNombreContacto
			sApellidos = ""
		End If
		
		If frmCliente.cbxComunaPersonal.selectedIndex = -1 Then
			sComunaPersonal = ""
		Else
			sComunaPersonal = frmCliente.cbxComunaPersonal(frmCliente.cbxComunaPersonal.selectedIndex).Text
		End If
		If frmCliente.cbxCiudadPersonal.selectedIndex = -1 Then
			sCiudadPersonal = ""
		Else
			sCiudadPersonal = frmCliente.cbxCiudadPersonal(frmCliente.cbxCiudadPersonal.selectedIndex).Text
		End If
		If frmCliente.cbxPaisPersonal.selectedIndex = -1 Then
			sPaisPersonal = ""
			sPaisLaboral = ""
		Else
			sPaisPersonal = frmCliente.cbxPaisPersonal(frmCliente.cbxPaisPersonal.selectedIndex).Text
			sPaisLaboral = frmCliente.cbxPaisPersonal(frmCliente.cbxPaisPersonal.selectedIndex).Text
		End If
		If frmCliente.cbxPaisPasaporte.selectedIndex = -1 Then
			sPaisPasaporte = ""
		Else
			sPaisPasaporte = frmCliente.cbxPaisPasaporte(frmCliente.cbxPaisPasaporte.selectedIndex).Text
		End If
		If frmCliente.cbxNacionalidad.selectedIndex = -1 Then
			sNacionalidad = ""
		Else
			sNacionalidad = frmCliente.cbxNacionalidad(frmCliente.cbxNacionalidad.selectedIndex).Text
		End If
		
		If CCur("0" & frmCliente.txtFonoPersonal.value) > 0 Then
			sFonoPersonal = "(" & frmCliente.txtPaisFonoPersonal.value & frmCliente.txtAreaFonoPersonal.value & ")" & " " & _
							frmCliente.txtFonoPersonal.value
		Else
			sFonoPersonal = ""
		End If
		If CCur("0" & frmCliente.txtFaxPersonal.value) > 0 Then
			sFaxPersonal = "(" & frmCliente.txtPaisFaxPersonal.value & frmCliente.txtAreaFaxPersonal.value & ")" & " " & _
							frmCliente.txtFaxPersonal.value
		Else
			sFaxPersonal = ""
		End If
		
		'Antecedentes Laborales
		If frmCliente.cbxComunaLaboral.selectedIndex = -1 Then
			sComunaLaboral = ""
		Else
			sComunaLaboral = frmCliente.cbxComunaLaboral(frmCliente.cbxComunaLaboral.selectedIndex).Text
		End If
		If frmCliente.cbxCiudadLaboral.selectedIndex = -1 Then
			sCiudadLaboral = ""
		Else
			sCiudadLaboral = frmCliente.cbxCiudadLaboral(frmCliente.cbxCiudadLaboral.selectedIndex).Text
		End If

		If CCur("0" & frmCliente.txtFonoLaboral.value) > 0 Then
			sFonoLaboral = "(" & frmCliente.txtPaisFonoLaboral.value & frmCliente.txtAreaFonoLaboral.value & ")" & " " & _
							frmCliente.txtFonoLaboral.value
		Else
			sFonoLaboral = ""
		End If
		If CCur("0" & frmCliente.txtFaxLaboral.value) > 0 Then
			sFaxLaboral = "(" & frmCliente.txtPaisFaxLaboral.value & frmCliente.txtAreaFaxLaboral.value & ")" & " " & _
							frmCliente.txtFaxLaboral.value
		Else
			sFaxLaboral = ""
		End If

		sParamatros = "http:../Reportes/FichaCliente.rpt?init=actx" & _
					  "&prompt0=" & frmCliente.txtRut.value & _
					  "&prompt1=" & MayMin(sNombreCliente) & _ 
					  "&prompt2=" & MayMin(sApellidos) & _
	 				  "&prompt3=" & MayMin(frmCliente.txtDireccionPersonal.value) & _
					  "&prompt4=" & MayMin(sComunaPersonal) & _
					  "&prompt5=" & MayMin(sCiudadPersonal) & _
					  "&prompt6=" & MayMin(sPaisPersonal) & _
					  "&prompt7=" & sFonoPersonal & _
					  "&prompt8=" & sFaxPersonal & _
					  "&prompt9=" & frmCliente.txtCorreo.value & _
					  "&prompt10=" & MayMin(sNombreEmpresa) & _
					  "&prompt11=" & MayMin(sApellidoContacto) & _
					  "&prompt12=" & MayMin(frmCliente.txtDireccionLaboral.value) & _
					  "&prompt13=" & MayMin(sComunaLaboral) & _
					  "&prompt14=" & MayMin(sCiudadLaboral) & _
					  "&prompt15=" & MayMin(sPaisLaboral) & _
					  "&prompt16=" & sFonoLaboral & _
					  "&prompt17=" & sFaxLaboral & _
					  "&prompt18="
							  
		sParamatros = replace(sParamatros, "ñ", "n")
		sParamatros = replace(sParamatros, "Ñ", "N")

		window.open sParamatros, _
					"", "dialogHeight= 800pxl; dialogWidth= 800pxl; " & _
					"dialogTop= 0; dialogLeft= 0; resizable=yes; " & _
					"status=0; scrollbars=1; toolbar=0"
	End Sub

	Sub ContratoProductos

		window.open "http:../Reportes/ContratoProductos.rpt?init=actx" , _
						"", "dialogHeight= 800pxl; dialogWidth= 800pxl; " & _
					    "dialogTop= 0; dialogLeft= 0; resizable=yes; " & _
						"status=0; scrollbars=1; toolbar=0"
		
	End Sub
			
	' JFMG 19-10-2007 se agregó combo con mas tipos
	Sub cbxTipo_onChange()
		
		if frmCliente.cbxTipo.value <> 1 then
			trEmpresa.style.display = ""	
			frmCliente.cbxRubro.value = ""
			trContacto.style.display = "none"
			trPersona.style.display = "none"
			frmCliente.txtApellidoM.value = ""
			frmCliente.txtApellidoP.value = ""
			frmCliente.txtNombres.value = ""
			tdSexo.style.display ="none"
			frmcliente.txtNombreEmpleador.value = ""
		else
			trEmpresa.style.display = "none"
			frmCliente.cbxRubro.value = ""
			trContacto.style.display = "none"
			trPersona.style.display = ""
			frmCliente.txtApellidoMContacto.value = ""
			frmCliente.txtApellidoPContacto.value = ""
			frmCliente.txtNombresContacto.value = ""
			frmCliente.txtRazonSocial.value = ""
			frmCliente.txtRepresentante.value = ""
			tdSexo.style.display =""
			frmcliente.txtGiroComercial.value  = ""'JBV25-05-2012
		end if
	end sub
	
	'************************************** Fin ****************************

	Sub optRut_onClick()
		window.frmcliente.optpasaporte.checked = 0
		window.frmcliente.optRut.checked = 1
		window.frmcliente.txtpasaporte.style.display="none"
		window.frmcliente.txtpasaporte.value = ""
		window.frmCliente.txtRut.style.display = ""
		frmCliente.cbxPaisPasaporte.value = ""
		tdPaisPasaporte.style.display="none"
		frmCliente.txtnumeroserie.style.display = ""
		tdNumeroSerie.style.display = ""
	End Sub

	Sub optPasaporte_onClick()
		window.frmcliente.optRut.checked = 0
		window.frmcliente.optPasaporte.checked = 1
		window.frmCliente.txtRut.style.display = "none"		
		window.frmCliente.txtRut.value = ""
		window.frmcliente.txtpasaporte.style.display=""
		tdPaisPasaporte.style.display=""
		frmCliente.txtnumeroserie.style.display = "none"
		tdNumeroSerie.style.display = "none"
	End Sub
	
	Sub MostrarLaboral()
		Dim sDisplay
		
		If frmCliente.cbxTipo.value = "1" then
			sDisplay = "none"			
		Else
			sDisplay = "none"
		End If
		trAntecedentesLaborales.style.display = sDisplay
	End Sub

	Sub cbxRubro_onblur()
		If frmCliente.cbxRubro.value = "" Then Exit Sub
		If frmCliente.cbxRubro.value = "<%=nRubro%>" Then Exit Sub
	End Sub

	Sub cbxSucursal_onblur()
		Dim nTipoCliente, nSexo, ClienteAgencia

		If frmCliente.cbxSucursal.value = "" Then Exit Sub
		If frmCliente.cbxSucursal.value = "<%=sSucursal%>" Then Exit Sub
		
		nTipoCLiente= frmCliente.cbxTipo.value 
		
		nSexo = frmCliente.cbxSexo.value 
		ClienteAgencia = frmCliente.chkClienteAgencia.checked
		'**************************************** Fin ***********************************

		frmCliente.action = "http:NuevoCliente.asp?Pais=" & frmCliente.cbxPaisPersonal.value & _
							 "&Ciudad=" & frmCliente.cbxCiudadPersonal.value & "&Comuna=" & frmCliente.cbxComunaPersonal.value  & _
							 "&PaisL=" & frmCliente.cbxPaisLaboral.value & _
							 "&CiudadL=" & frmCliente.cbxCiudadLaboral.value & "&ComunaL=" & frmCliente.cbxComunalaboral.value & _
							 "&TipoCliente=" & nTipoCliente & "&Rut=" & frmCliente.txtRut.value & _
							 "&Pasaporte=" & frmCliente.txtPasaporte.value & "&PP=" & frmCliente.cbxPaisPasaporte.value & _
							 "&sc=" & frmCliente.cbxSucursal.value & "&cb=" & frmCliente.cbxBanco.value & _
							 "&nc=" & frmCliente.cbxNacionalidad.value & "&rb=" & frmCliente.cbxRubro.value & _
							 "&SerieIdentificacion=" & frmCliente.txtNumeroserie.value  & "&nSexo=" & nSexo & _
							 "&ClienteAgencia=" & frmCliente.chkClienteAgencia.checked & _
							 "&TipoOrigenLlamada=" & "<%=request("TipoOrigenLlamada")%>" & _
							 "&SCC=" & "<%=request("SCC")%>" & _
		                     "&AGC=" & "<%=request("AGC")%>" & _
		    		         "&NUO=" & "<%=request("NUO")%>" & _
		    		         "&NivelRiesgo=NORMAL"
		frmCliente.submit 
		frmCliente.action = ""
	End Sub

	Sub cbxEjecutivo_onblur()
		If frmCliente.cbxEjecutivo.value = "" Then Exit Sub
		If frmCliente.cbxEjecutivo.value = "<%=nEjecutivo%>" Then Exit Sub
	End Sub

	Sub cbxPaisPasaporte_onblur()
		If frmCliente.cbxPaisPasaporte.value = "" Then Exit Sub
		If frmCliente.cbxPaisPasaporte.value = "<%=sPaisPasaporte%>" Then Exit Sub
	End Sub
	
	Sub cbxNacionalidad_onblur()
		If frmCliente.cbxNacionalidad.value = "" Then Exit Sub
		If frmCliente.cbxNacionalidad.value = "<%=sNacionalidad%>" Then Exit Sub
	End Sub

	Sub cbxPaisPersonal_onblur()
		Dim sCiudad, nTipoCliente, nSexo
		
		If frmCliente.cbxPaisPersonal.value = "" Then Exit Sub
		If frmCliente.cbxPaisPersonal.value = "<%=sPais%>" Then Exit Sub
		
		nTipoCliente = frmCliente.cbxTipo.value
		nSexo = frmCliente.cbxSexo.value 
		ClienteAgencia = frmCliente.chkClienteAgencia.checked
		
		'**************************************** Fin ***********************************

		frmCliente.action = "http:NuevoCliente.asp?Pais=" & frmCliente.cbxPaisPersonal.value & _
							 "&Ciudad=" & sCiudad & "&Comuna=" & frmCliente.cbxComunaPersonal.value  & _
							 "&PaisL=" & frmCliente.cbxPaisLaboral.value & _
							 "&CiudadL=" & frmCliente.cbxCiudadLaboral.value & "&ComunaL=" & frmCliente.cbxComunalaboral.value & _
							 "&TipoCliente=" & nTipoCliente & "&Rut=" & frmCliente.txtRut.value & _
							 "&Pasaporte=" & frmCliente.txtPasaporte.value & "&PP=" & frmCliente.cbxPaisPasaporte.value & _
							 "&sc=" & frmCliente.cbxSucursal.value & "&ce=" & frmCliente.cbxEjecutivo.value & "&cb=" & frmCliente.cbxBanco.value & _
							 "&nc=" & frmCliente.cbxNacionalidad.value & "&rb=" & frmCliente.cbxRubro.value  & _
							 "&SerieIdentificacion=" & frmCliente.txtNumeroserie.value & "&nSexo=" & nSexo & _
							 "&ClienteAgencia=" & frmCliente.chkClienteAgencia.checked & _
							 "&TipoOrigenLlamada=" & "<%=request("TipoOrigenLlamada")%>" & _
							 "&SCC=" & "<%=request("SCC")%>" & _
		                     "&AGC=" & "<%=request("AGC")%>" & _
		    		         "&NUO=" & "<%=request("NUO")%>" & _
		    		         "&NivelRiesgo=NORMAL"
		frmCliente.submit 
		frmCliente.action = ""
	End Sub
	
	Sub cbxCiudadPersonal_onblur()
		Dim sComuna, nTipoCliente, nsexo 
		
		If frmCliente.cbxCiudadPersonal.value = "" Then Exit Sub
		If frmCliente.cbxCiudadPersonal.value = "<%=sCiudad%>" Then Exit Sub
		nTipoCliente = frmCliente.cbxTipo.value
	
		nsexo = frmCliente.cbxSexo.value
		CLienteAgencia = frmCliente.chkClienteAgencia.checked
		'**************************************** Fin ***********************************
		
		frmCliente.action = "http:NuevoCliente.asp?Pais=" & frmCliente.cbxPaisPersonal.value & _
							 "&Ciudad=" & frmCliente.cbxCiudadPersonal.value & "&Comuna=" & sComuna  & _
							 "&PaisL=" & frmCliente.cbxPaisLaboral.value & _
							 "&CiudadL=" & frmCliente.cbxCiudadLaboral.value & "&ComunaL=" & frmCliente.cbxComunalaboral.value & _
							 "&TipoCliente=" & nTipoCliente & "&Rut=" & frmCliente.txtRut.value & _
							 "&Pasaporte=" & frmCliente.txtPasaporte.value & "&PP=" & frmCliente.cbxPaisPasaporte.value & _
							 "&sc=" & frmCliente.cbxSucursal.value & "&ce=" & frmCliente.cbxEjecutivo.value & "&cb=" & frmCliente.cbxBanco.value & _
							 "&nc=" & frmCliente.cbxNacionalidad.value & "&rb=" & frmCliente.cbxRubro.value & _
							 "&SerieIdentificacion=" & frmCliente.txtNumeroserie.value & "&nSexo=" & nSexo & _
							 "&Clienteagencia=" & frmCliente.chkClienteAgencia.checked & _
							 "&TipoOrigenLlamada=" & "<%=request("TipoOrigenLlamada")%>" & _
							 "&SCC=" & "<%=request("SCC")%>" & _
		                     "&AGC=" & "<%=request("AGC")%>" & _
		    		         "&NUO=" & "<%=request("NUO")%>" & _
		    		         "&NivelRiesgo=NORMAL"
		frmCliente.submit 
		frmCliente.action = ""
	End Sub

	Sub cbxPaisLaboral_onblur()
		Dim sCiudad, nSexo
		
		If frmCliente.cbxPaisLaboral.value = "" Then Exit Sub
		If frmCliente.cbxPaisLaboral.value = "<%=sPaisL%>" Then Exit Sub

		nSexo = frmCliente.cbxSexo.value 
		ClienteAgencia = frmCliente.chkClienteAgencia.checked
		frmCliente.action = "http:NuevoCliente.asp?Pais=" & frmCliente.cbxPaisPersonal.value & _
							 "&Ciudad=" & frmCliente.cbxCiudadPersonal.value & "&Comuna=" & frmCliente.cbxComunaPersonal.value  & _
							 "&PaisL=" & frmCliente.cbxPaisLaboral.value & _
							 "&CiudadL=" & sCiudad & "&ComunaL=" & frmCliente.cbxComunalaboral.value & _
							 "&SerieIdentificacion=" & frmCliente.txtNumeroserie.value & "&nSexo=" & nSexo & _
							 "&ClienteAgencia=" & frmCliente.chkClienteAgencia.checked & _
							 "&TipoOrigenLlamada=" & "<%=request("TipoOrigenLlamada")%>" & _
							 "&SCC=" & "<%=request("SCC")%>" & _
		                     "&AGC=" & "<%=request("AGC")%>" & _
		    		         "&NUO=" & "<%=request("NUO")%>" & _
		    		         "&NivelRiesgo=NORMAL"
		frmCliente.submit 
		frmCliente.action = ""
	End Sub
		
	Sub cbxCiudadLaboral_onblur()
		Dim sComuna
		
		If frmCliente.cbxCiudadLaboral.value = "" Then Exit Sub
		If frmCliente.cbxCiudadLaboral.value = "<%=sCiudadL%>" Then Exit Sub
		
		frmCliente.action = "http:NuevoCliente.asp?Pais=" & frmCliente.cbxPaisPersonal.value & _
							 "&Ciudad=" & frmCliente.cbxCiudadPersonal.value & "&Comuna=" & frmCliente.cbxComunaPersonal.value & _
							 "&PaisL=" & frmCliente.cbxPaisLaboral.value & _
							 "&CiudadL=" & frmCliente.cbxCiudadLaboral.value & "&ComunaL=" & sComuna & _
							 "&SerieIdentificacion=" & frmCliente.txtNumeroserie.value & "&Sexo=" & frmCliente.cbxSexo.value  & _
							 "&ClienteAgencia="  & frmCliente.chkCLienteAgencia.checked & _
							 "&TipoOrigenLlamada=" & "<%=request("TipoOrigenLlamada")%>" & _
							 "&SCC=" & "<%=request("SCC")%>" & _
		                     "&AGC=" & "<%=request("AGC")%>" & _
		    		         "&NUO=" & "<%=request("NUO")%>" & _
		    		         "&NivelRiesgo=NORMAL"
		frmCliente.submit 
		frmCliente.action = ""
	End Sub
	
	Sub cmdValidarRegistro_onClick()

	    Dim sRut, sPasaporte, sNumeroRut, sTipo
	    
		If (frmCliente.optRutCautela.checked = true and (frmCliente.txtRut.value = Empty or frmCliente.txtNumeroserie.value = Empty) ) Then
		    msgbox "Debe ingresar el Rut y número de serie Correctamente"
		    frmCliente.txtRut.focus
		    Exit Sub
		End if	
		
		If (frmCliente.txtPasaporte.value = Empty and frmCliente.optPasaporte.checked = true) Then
		    msgbox "Debe ingresar el número de Pasaporte Correctamente"
		    frmCliente.txtPasaporte.focus
		    Exit Sub
		End if		
		
		If frmCliente.txtRut.value <> "" Then
			sRut = ValidarRut(frmCliente.txtRut.value)
			
			frmCliente.txtRut.value = sRut
		
			sNumeroRut = replace(replace((trim(sRut)),"-",""),".","")
			sNumeroRut = left(sNumeroRut, len(sNumeroRut) - 1)
		
		elseif frmCliente.txtPasaporte.value <> "" Then
			sNumeroRut = frmCliente.txtPasaporte.value
			
		Else
			msgbox "Debe ingresar una Identificación "
			Exit Sub
		End If
		
		msgbox "Registro Validado!"
		
	End Sub
		
	'AMP 31-07-2015
	Sub txtRut_onblur()
		Dim sRut, sDigito, sPasaporte, sRespuestaRegistro, sNumeroRut, nSexo
		
		If frmCliente.txtRut.value = Empty Then Exit Sub
		
		sRut = ValidarRut(frmCliente.txtRut.value)
		sPasaporte = Empty
		
		If sRut <> "" Then
			frmCliente.txtUsuario.value = frmCliente.txtRut.value
			frmCliente.txtRut.value = sRut
		Else
			msgbox "Debe ingresar un número de rut válido ", vbOKOnly + vbInformation, "Ingreso de Cliente"
			frmCliente.txtRut.focus
			Exit Sub
		End If
		nSexo = frmCliente.cbxSexo.value  
		    frmCliente.action = "http:NuevoCliente.asp?Rut=" & sRut & "&Pasaporte=" & sPasaporte & _
							 "&SerieIdentificacion=" & frmCliente.txtNumeroserie.value & "&nSexo=" & nSexo & _
							 "&TipoOrigenLlamada=" & "<%=request("TipoOrigenLlamada")%>" & _
							 "&SCC=" & "<%=request("SCC")%>" & _
		                     "&AGC=" & "<%=request("AGC")%>" & _
		    		         "&NUO=" & "<%=request("NUO")%>" & _
		    		         "&NivelRiesgo=NORMAL" 
		    		         
		    frmCliente.submit
		    frmCliente.action = ""
		
		
	End Sub
	'FIN AMP 31-07-2015

	Sub txtPasaporte_onblur()
		Dim sRut, sPasaporte, nSexo
		
		If frmCliente.txtPasaporte.value = Empty Then Exit Sub
		
		sRut = Empty
		sPasaporte = frmCliente.txtPasaporte.value
		
		If sPasaporte <> "" Then
			frmCliente.txtUsuario.value = frmCliente.txtPasaporte.value
			frmCliente.txtPasaporte.value = sPasaporte
		Else
			msgbox "Debe ingresar pasaporte", vbOKOnly + vbInformation, "Ingreso de Cliente"
			frmCliente.txtPasaporte.focus
			Exit Sub
		End If
		nSexo = frmCliente.cbxSexo.value 
		
		frmCliente.action = "NuevoCliente.asp?Rut=" & sRut & "&Pasaporte=" & sPasaporte & _
							 "&SerieIdentificacion=" & frmCliente.txtNumeroserie.value & "nSexo=" & nSexo & _
							 "&TipoOrigenLlamada=" & "<%=request("TipoOrigenLlamada")%>" & _
							 "&SCC=" & "<%=request("SCC")%>" & _
		                     "&AGC=" & "<%=request("AGC")%>" & _
		    		         "&NUO=" & "<%=request("NUO")%>" & _
		    		         "&NivelRiesgo=NORMAL" 
		frmCliente.submit
		frmCliente.action = ""
	End Sub

	'MS 28-03-2014
	Sub optActEconomicaSi_onClick()
		window.frmcliente.optActEconomicaNo.checked = 0
		window.frmcliente.optActEconomicaSi.checked = 1
	End Sub
	
	Sub optActEconomicaNo_onClick()
		window.frmcliente.optActEconomicaNo.checked = 1
		window.frmcliente.optActEconomicaSi.checked = 0
	End Sub
	'FIN 28-03-2014
	
	'AMP 19-06-2015
	Sub optRutCautela_onClick()
		window.frmcliente.optPasCautela.checked = 0
		window.frmcliente.optRutCautela.checked = 1
		window.frmcliente.txtPasCautela.style.display="none"
		window.frmCliente.txtRutCautela.style.display = ""
		frmCliente.txtNumSerieCautela.style.display = ""
		tdNSerieCautela.style.display = ""
        tdPaisPasaporteCautela.style.display="none"
		window.frmcliente.txtRutCautela.focus
        LimpiaDatosCautela
	End Sub

	Sub optPasCautela_onClick()
		window.frmcliente.optRutCautela.checked = 0
		window.frmcliente.optPasCautela.checked = 1
		window.frmCliente.txtRutCautela.style.display = "none"		
		window.frmcliente.txtPasCautela.style.display=""
		window.frmcliente.txtNumSerieCautela.style.display = "none"
		window.frmcliente.txtNumSerieCautela.value = ""
		tdNSerieCautela.style.display = "none"
        tdPaisPasaporteCautela.style.display=""
		window.frmcliente.txtPasCautela.focus
		LimpiaDatosCautela
				
	End Sub
	'FIN AMP 19-06-2015
	
-->
</script>

<!--#INCLUDE virtual="/Compartido/Rutinas.htm" -->
<!--#INCLUDE virtual="/Compartido/Encabezado.htm" -->
<body>
    <object id="objTab" style="height: 0px; width: 0px" type="text/x-scriptlet" viewastext>
        <param name="Scrollbar" value="0">
        <param name="URL" value="http:../Scriptlets/Tab.htm">
    </object>
    <form id="frmCliente" method="post">
    <table border="0" cellpadding="1" cellspacing="0" style="height: 10px; width: 570;
        background-color: #f4f4f4">
        <tr>
            <td>
                <table id="tabConsulta" border="0" cellpadding="0" cellspacing="0" style="height: 150px;
                    width: 570; background-color: #f4f4f4">
                    <tr height="40" style="background-color: #ffeeaa; #ffdd77; #e1e1e1">
                        <td colspan="3" style="font-size: 16pt">
                            &nbsp;&nbsp;Nuevo Cliente
                        </td>
                    </tr>
                    <tr height="1" style="background-color: silver">
                        <td colspan="3">
                        </td>
                    </tr>
                    <tr height="20">
                        <td colspan="3">
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
        <tr>
            <td>
                <table cellspacing="1" cellpadding="1" swidth="30%" style="color: #505050; font-family: Verdana;
                    font-size: 10pt; position: relative; top: 0px; sborder: 1; background-color: silver;">
                    <tr height="1">
                        <td id="tdDocumento" style="background-color: #ffffcc; #e1e1e1">
                            <b>&nbsp;&nbsp;Antecedentes del Usuario&nbsp;&nbsp;</b>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
        <tr height="1" style="background-color: silver">
            <td colspan="3">
            </td>
        </tr>
        <tr>
            <td>
                <table>
                    <tr>
                        <td>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            Nivel de Riesgo<br>
                            <select name="cbxRiesgo">
                                <%	CargarEstado "NIVELRIESGO", iRiesgo %>
                            </select>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
    <div id="divDatosCautela" style="display: none">
        <table border="0" cellpadding="0" cellspacing="0" style="height: 10px; width: 570px;
            background-color: #f4f4f4">
            <tr>
                <td>
                    <input type="radio" name="optRutCautela" checked>
                    Rut
                    <input type="radio" name="optPasCautela">
                    Pasaporte
                </td>
                <td width="150" id="tdPaisPasaporteCautela2" style="display: none">
                    Pa&iacute;s
                </td>
                <td id="tdNSerieCautela">
                    N&uacute;mero de Serie
                </td>
            </tr>
            <tr>
                <td>
                    <input name="txtRutCautela" maxlength="10">
                    <input name="txtPasCautela" style="display: none">
                </td>
                <td id="tdPaisPasaporteCautela" style="display: none">
                    <select id="cbxPaisPasaporteCautela" name="cbxPaisPasaporteCautela" style="width: 150px">
                        <%	CargarPaisPasaporte sPaisPasaporte %>
                    </select>
                </td>
                <td>
                    <input type="text" name="txtNumSerieCautela" id="txtNumSerieCautela" value="<%=sSerieIdentificacion%>" maxlength="15">&nbsp;&nbsp;
                </td>
            </tr>
            <tr>
                <td>
                    &nbsp;
                </td>
                 <td id="tdPaisPasaporteCautela3" style="display: none">
                    &nbsp;
                </td>
                <td>
                    <input type="button" name="btn_ValidarRegistroCautela" id = "btn_ValidarRegistroCautela" value="Validar Registro">
                </td>
            </tr>
            <tr>
                <td>
                    Apellido Paterno<br />
                    <input name="txtApellidoPCautela" id="txtApellidoPCautela" style="width: 160px" onkeypress="IngresarTexto(3)"
                        maxlength="20" value="<%=sApellidoPaternoCliente%>">
                </td>
                <td>
                    Apellido Materno<br />
                    <input name="txtApellidoMCautela" id="txtApellidoMCautela" style="width: 160px" onkeypress="IngresarTexto(3)"
                        maxlength="20" value="<%=sApellidoMaternoCliente%>" />
                </td>
                <td>
                    Nombres<br />
                    <input name="txtNombresCautela" id="txtNombresCautela" style="width: 180px" onkeypress="IngresarTexto(3)"
                        maxlength="30" value="<%=sNombresCliente%>" />
                </td>
            </tr>
            <tr>
                <td colspan="4">
                     Descripci&oacute;n<br />
                    <input name="txtDescripcionCautela" id="txtDescripcionCautela" style="width: 547px" onkeypress="IngresarTexto(3)"
                        maxlength="40" value="<%sDescripcionCautela %>" />
                </td>
            </tr>
        </table>
    </div>
    <div id="divDatosNormal" style="display: none">
        <table border="0" cellpadding="1" cellspacing="0" style="height: 10px; width: 570px;
            background-color: #f4f4f4">
            <tr height="80">
                <td>
                    <table>
                        <tr>
                            <td width="148">
                                <input type="radio" name="optRut" checked>
                                Rut
                                <input type="radio" name="optPasaporte">
                                Pasaporte
                            </td>
                            <td width="150" id="tdPaisPasaporte2" style="display: none">
                                Pa&iacute;s
                            </td>
                            <td width="178" id="tdNumeroSerie">
                                N&uacute;mero de Serie
                            </td>
                        </tr>
                        <tr>
                            <td>
                                <input name="txtRut">
                                <input name="txtPasaporte" style="display: none">
                            </td>
                            <td id="tdPaisPasaporte" style="display: none">
                                <select id="cbxPaisPasaporte" name="cbxPaisPasaporte" style="width: 150px">
                                    <%	CargarPaisPasaporte sPaisPasaporte ' INTERNO-1831 - JFMG 28-07-2014 	%>
                                </select>
                            </td>
                            <td>
                                <input type="text" name="txtNumeroserie" value="<%=sSerieIdentificacion%>">&nbsp;&nbsp;
                            </td>
                        </tr>
                        <tr>
                            <td>
                                &nbsp;
                            </td>
                            <td id="tdPaisPasaporte3" style="display: none">
                                &nbsp;
                            </td>
                            <td>
                                <input type="button" name="cmdValidarRegistro" value="Validar Registro">
                            </td>
                        </tr>
                    </table>
                    <table style="width: 558px">
                        <td width="113">
                            Tipo Cliente<br />
                            <select name="cbxTipo">
                                <%CargarTipo "CLIENTE", nTipoCliente%>
                            </select>
                        </td>
                        <td width="125" id="tdSexo">
                            Sexo<br />
                            <select name="cbxSexo">
                                <option value="0"></option>
                                <%CargarTipo "SEXO", nSEXO%>
                            </select>
                        </td>
                        <td width="148">
                            &nbsp;
                            <input id="chkCLienteAgencia" type="checkbox" <%=nclienteAgencia%> />
                            Cliente agencia
                        </td>
                    </table>
                </td>
            </tr>
            <tr id="trEmpresa" height="20" style="display: none">
                <td>
                    <table>
                        <tr>
                            <td>
                                Razón Social<br>
                                <input name="txtRazonSocial" id="txtRazonSocial" size="40" style="width: 180px" onkeypress="IngresarTexto(3)"
                                    value="<%=sRazonSocialCliente%>">
                            </td>
                            <td>
                                Representante Legal<br>
                                <input name="txtRepresentante" id="txtRepresentante" size="40" style="width: 180px"
                                    onkeypress="IngresarTexto(3)">
                            </td>
                            <td>
                                Rubro<br>
                                <select id="cbxRubro" name="cbxRubro" style="width: 190px">
                                    <%	CargarRubro nRubro 	%>
                                </select>
                            </td>
                        </tr>
                        <tr>
                            <td colspan="3">
                                Giro Comercial<br>
                                <input name="txtGiroComercial" id="txtGiroComercial" size="40" style="width: 550px"
                                    onkeypress="IngresarTexto(3)">
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr id="trPersona" height="20">
                <td>
                    <table>
                        <tr>
                            <td colspan="4">
                                <table>
                                    <tr>
                                        <td>
                                            Apellido Paterno<br />
                                            <input name="txtApellidoP" id="txtApellidoP" style="width: 180px" onkeypress="IngresarTexto(3)"
                                                maxlength="20" value="<%=sApellidoPaternoCliente%>">
                                        </td>
                                        <td>
                                            Apellido Materno<br />
                                            <input name="txtApellidoM" id="txtApellidoM" style="width: 160px" onkeypress="IngresarTexto(3)"
                                                maxlength="20" value="<%=sApellidoMaternoCliente%>" />
                                        </td>
                                        <td>
                                            Nombres<br />
                                            <input name="txtNombres" id="txtNombres" size="30" onkeypress="IngresarTexto(3)"
                                                maxlength="30" value="<%=sNombresCliente%>" />
                                        </td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                        <tr>
                            <td>
                                Fecha nacimiento<br>
                                <input name="txtFechaNacimiento" id="txtFechaNacimiento" size="12" maxlength="100"
                                    value="<%=sFechaNacimiento%>" />(dd-mm-aaaa)
                            </td>
                             <td>
                                Ocupación<br/>
                                <select name="cbxOcupacion" style="width: 233px" <%=sHabilitado%>>
                                    <%	CargarOcupacion sOcupacion %>
                                </select><!--INTERNO-9263 MS 19-01-2017-->
                            </td>
                        </tr>
                        <tr>
                            <td>
                                Profesión<br />
                                <input name="txtProfesion" id="txtProfesion" size="30" onkeypress="IngresarTexto(2)"
                                    maxlength="30" />
                            </td>
                            <td colspan="3">
                                Nombre Empleador<br />
                                <input name="txtNombreEmpleador" id="txtNombreEmpleador" style="width: 325px" onkeypress="IngresarTexto(2)"
                                    size="50" maxlength="100" value="<%=sNombreEmpleador%>" />
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
                                Nacionalidad<br>
                                <select id="cbxNacionalidad" name="cbxNacionalidad" style="width: 223px">
                                    <%	CargarUbicacion 1, "", sNacionalidad 	%>
                                </select>
                            </td>
                            <td>
                                Correo electrónico<br>
                                <input name="txtCorreo" id="txtCorreo" style="width: 325px" onkeypress="IngresarTexto(6)"
                                    maxlength="42" value="<%=sCorreoElectronicoCliente%>">
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr height="20">
                <td>
                    <table>
                        <tr id="trUsuario">
                            <td>
                                Nombre de usuario<br>
                                <input name="txtUsuario" id="txtUsuario" maxlength="12" readonly>
                            </td>
                            <td>
                                Password<br>
                                <input type="password" name="txtPassword" id="txtPassword" maxlength="10">
                            </td>
                            <td>
                                Confirmar Password<br>
                                <input type="password" name="txtConfPassword" id="txtConfPassword" maxlength="10">
                            </td>
                        </tr>
                        <tr id="preg_res">
                            <td>
                                Pregunta<br>
                                <input name="txtPregunta" id="txtPregunta" maxlength="100">
                            </td>
                            <td>
                                Respuesta<br>
                                <input name="txtRespuesta" id="txtRespuesta" maxlength="100">
                            </td>
                            <td>
                                <br>
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
                                <br>
                                <input name="txtNombresContacto" id="txtNombresContacto" size="25" onkeypress="IngresarTexto(2)"
                                    maxlength="30">
                            </td>
                            <td>
                                Apellido Paterno<br>
                                <input name="txtApellidoPContacto" id="txtApellidoPContacto" onkeypress="IngresarTexto(2)"
                                    maxlength="20">
                            </td>
                            <td>
                                Apellido Materno<br>
                                <input name="txtApellidoMContacto" id="txtApellidoMContacto" onkeypress="IngresarTexto(2)"
                                    maxlength="20">
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr height="100%">
                <td>
                    &nbsp;
                </td>
            </tr>
        </table>
        <!-- Paso 3 -->
        <table border="0" cellpadding="0" cellspacing="0" height="180" width="570" style="background-color: #f4f4f4">
            <tr>
                <td>
                    <table cellspacing="1" cellpadding="1" swidth="30%" style="color: #505050; font-family: Verdana;
                        font-size: 10pt; position: relative; top: 0px; sborder: 1; background-color: silver;">
                        <tr height="1">
                            <td id="tdDocumento" style="background-color: #ffffcc; #e1e1e1">
                                <b>&nbsp;&nbsp;Antecedentes Personales&nbsp;&nbsp;</b>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr height="1" style="background-color: silver">
                <td colspan="3">
                </td>
            </tr>
            <tr height="80">
                <td>
                    <table width="73%" border="0" cellspacing="0" cellpadding="1">
                        <tr>
                            <td width="50%">
                                Direcci&oacute;n
                            </td>
                            <td width="15%">
                                N&uacute;mero
                            </td>
                            <td width="35%">
                                Depto/Oficina
                            </td>
                        </tr>
                        <tr>
                            <td>
                                <input name="txtDireccionPersonal" size="30" onkeypress="IngresarTexto(3)" maxlength="40"
                                    value="<%=sDireccionCliente%>">
                            </td>
                            <td>
                                <input name="txtNumero" size="10" onkeypress="IngresarTexto(3)" maxlength="5">
                            </td>
                            <td>
                                <input name="txtdepto" size="10" onkeypress="IngresarTexto(3)" maxlength="5">
                            </td>
                        </tr>
                        <tr>
                            <td>
                                Pa&iacute;s
                            </td>
                            <td colspan="2">
                                Ciudad
                            </td>
                        </tr>
                        <tr>
                            <td>
                                <select name="cbxPaisPersonal" style="width: 160px">
                                    <%	CargarUbicacion 1, "", sPais 	%>
                                </select>
                            </td>
                            <td colspan="2">
                                <select name="cbxCiudadPersonal" style="width: 161px">
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
                                Comuna<br>
                                <select name="cbxComunaPersonal" style="width: 150px">
                                    <% 
								If sPais = "CL" Then
                                    CargarComunaCiudad  sCiudad, sComuna
								End If
                                    %>
                                </select>
                            </td>
                            <td>
                                Teléfono<br>
                                <input name="txtPaisFonoPersonal" style="width: 30px" disabled value="<%=iDdiPaisCliente%>">
                                <input name="txtAreaFonoPersonal" style="width: 30px" disabled value="<%=iDdiCiudadCliente%>">
                                <input name="txtFonoPersonal" style="width: 70px" maxlength="10" value="<%=iNumeroTelefonoCliente%>">
                            </td>
                            <td>
                                Fax<br>
                                <input name="txtPaisFaxPersonal" style="width: 30px" disabled>
                                <input name="txtAreaFaxPersonal" style="width: 30px" disabled>
                                <input name="txtFaxPersonal" style="width: 70px" maxlength="10">
                            </td>
                            <td>
                                Celular<br>
                                <input name="txtCelular" style="width: 90px" maxlength="10">
                            </td>
                        </tr>
                    </table>
                    <table>
                        <tr>
                            <td>
                                Sucursal<br>
                                <select name="cbxSucursal" style="width: 240px">
                                    <%	CargarSucursal sSucursal	%>
                                </select>
                            </td>
                            <td>
                                Ejecutivo<br>
                                <select name="cbxEjecutivo" style="width: 308px">
                                    <%	CargarEjecutivos sSucursal, nEjecutivo %>
                                </select>
                            </td>
                        </tr>
                    </table>
                    <table>
                        <tr>
                            <td>
                                Banco<br>
                                <select name="cbxBanco" style="width: 240px">
                                    <%	CargarBanco nCodigoBanco%>
                                </select>
                            </td>
                            <td>
                                Cuenta Corriente<br>
                                <input name="txtCuentaCorriente" style="width: 153px" maxlength="20">
                            </td>
                            <td>
                                Cuenta de Ahorro<br>
                                <input name="txtCuentaAhorro" style="width: 152px" maxlength="20">
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
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
                                Empresa<br>
                                <input name="txtNombreEmpresa" size="50" onkeypress="IngresarTexto(3)" maxlength="40">
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
                                Dirección<br>
                                <input name="txtDireccionLaboral" size="30" onkeypress="IngresarTexto(3)" maxlength="40">
                            </td>
                            <td style="display: none">
                                Pais<br>
                                <select name="cbxPaisLaboral" style="width: 150px">
                                    <% CargarUbicacion 1, "", sPaisL  %>
                                </select>
                            </td>
                            <td>
                                Ciudad<br>
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
                                Comuna<br>
                                <select name="cbxComunaLaboral" style="width: 150px">
                                    <% 
								If sPais = "CL" Then
                                    CargarComunaCiudad  sCiudadL, sComunaL
								End If
                                    %>
                                </select>
                            </td>
                            <td>
                                Teléfono<br>
                                <input name="txtPaisFonoLaboral" style="width: 40px" disabled>
                                <input name="txtAreaFonoLaboral" style="width: 40px">
                                <input name="txtFonoLaboral" style="width: 100px" maxlength="10">
                            </td>
                            <td>
                                Fax<br>
                                <input name="txtPaisFaxLaboral" style="width: 40px" disabled>
                                <input name="txtAreaFaxLaboral" style="width: 40px">
                                <input name="txtFaxLaboral" style="width: 100px" maxlength="10">
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <!-- height="70"-->
                <td>
                    &nbsp;
                </td>
            </tr>
        </table>
        <!--Jonathan Miranda G. 14-03-2007-->
        <table cellspacing="0" cellpadding="0" width="570" style="font-family: Verdana; font-size: 10pt;
            position: relative; top: 0px; sborder: 1; background-color: #F4F4F4">
            <tr>
                <td>
                    <table cellspacing="1" cellpadding="1" swidth="30%" style="color: #505050; font-family: Verdana;
                        font-size: 10pt; position: relative; top: 0px; sborder: 1; background-color: silver;">
                        <tr height="1">
                            <td style="background-color: #ffffcc; #e1e1e1">
                                <b>&nbsp;&nbsp;Perfil Operacional&nbsp;&nbsp;</b>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr height="1" style="background-color: silver">
                <td colspan="4">
                </td>
            </tr>
            <tr height="80">
                <td>
                    <table>
                        <tr>
                            <td>
                            </td>
                        </tr>
                        <tr>
                            <!--<td>
                                Nivel de Riesgo<br>
                                <select name="cbxRiesgo">
                                <%	CargarEstado "NIVELRIESGO", iRiesgo %>
                            </select>
                            </td>-->
                            <td>
                                PEP<br>
                                <select name="cbxPerfilPEP">
                                    <%	CargarPerfil "1", iPerfilPEP %>
                                </select>
                            </td>
                        </tr>
                        <tr>
                            <td>
                                Zona<br>
                                <select name="cbxPerfilZona">
                                    <%	CargarPerfil "2", iPerfilZona %>
                                </select>
                            </td>
                        </tr>
                        <tr>
                            <td>
                                Residencia<br>
                                <select name="cbxPerfilRS">
                                    <%	CargarPerfil "3", iPerfilRS %>
                                </select>
                            </td>
                            <td style="display: none">
                                Actividad<br>
                                <select name="cbxPerfilACT">
                                    <%	CargarPerfil "4", iPerfilACT %>
                                </select>
                            </td>
                        </tr>
                        <tr>
                            <td>
                                Industria MBS<br>
                                <select name="cbxPerfilIndustria">
                                    <%	CargarPerfil "5", iPerfilIndustria %>
                                </select>
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
        </table>
        <!------------------  Fin  ------------------------------>
        <!--Jonathan Miranda G. 13-06-2011-->
        <table cellspacing="0" cellpadding="0" width="570" style="font-family: Verdana; font-size: 10pt;
            position: relative; top: 0px; sborder: 1; background-color: #F4F4F4">
            <tr>
                <td colspan="4">
                    <table cellspacing="1" cellpadding="1" swidth="30%" style="color: #505050; font-family: Verdana;
                        font-size: 10pt; position: relative; top: 0px; sborder: 1; background-color: silver;">
                        <tr height="1">
                            <td id="td1" style="background-color: #ffffcc; #e1e1e1">
                                <b>&nbsp;&nbsp;Perfil Transaccional&nbsp;&nbsp;</b>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr height="1" style="background-color: silver">
                <td colspan="4">
                </td>
            </tr>
            <tr height="80">
                <td colspan="4">
                    <table>
                        <tr>
                            <td align="center" colspan="2">
                                <table width="100%">
                                    <!--MS 28-03-2014-->
                                    <tr>
                                        <td valign="top" colspan="3">
                                            Actividad Econ&oacute;mica<br />
                                            <input type="radio" name="optActEconomicaSi" checked>
                                            Si
                                            <input type="radio" name="optActEconomicaNo">No
                                        </td>
                                    </tr>
                                    <!--MS 28-03-2014-->
                                    <tr>
                                        <td valign="top" colspan="3">
                                            <select name="cbxActividadEconomicaOculta" style="width: 450px; font-size: 10px;
                                                display: none;">
                                                <% CargarActividadEconomica "" %>
                                            </select>
                                            <select name="cbxActividadEconomica" style="width: 450px; font-size: 10px;">
                                                <% CargarActividadEconomica "" %>
                                            </select><input type="button" name="cmdAgregarActividad" value="Agregar" />
                                        </td>
                                    </tr>
                                    <tr>
                                        <td>
                                            <table id="tblActividadEconomica" cellspacing="1" cellpadding="4" style="color: #505050;
                                                font-family: Verdana; font-size: 10px; border: 1;">
                                                <tr style="height: 20px" align="center">
                                                    <td colspan="3" style="background-color: #e1e1e1; font-size: 12px;" width="40%">
                                                        <b>Descripci&oacute;n</b>
                                                    </td>
                                                </tr>
                                            </table>
                                            <input type="hidden" name="txtActividadEconomica" value="" />
                                        </td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                        <tr>
                            <td colspan="2">
                                Prop&oacute;sito de las Transacciones<br />
                                <input type="text" name="txtPropositoTransacciones" style="width: 547px; height: 50px;"
                                    onkeypress="IngresarTexto(3)" maxlength="500" />
                            </td>
                        </tr>
                        <tr>
                            <td colspan="2">
                                Ref. Comercial / Bancaria<br />
                                <input type="text" name="txtReferenciaBancaria" style="width: 547px" onkeypress="IngresarTexto(3)"
                                    maxlength="200" />
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
                                                text-align: right;" onkeypress="IngresarTexto(1)" maxlength="3" />
                                        </td>
                                        <td>
                                            Monto USD<br />
                                            <input type="text" name="txtMontoTransaccionesMesCLIENTE" style="width: 80px; text-align: right;"
                                                onkeypress="IngresarTexto(1)" maxlength="18" />
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
                                                onkeypress="IngresarTexto(1)" maxlength="3" />
                                        </td>
                                        <td>
                                            Monto USD<br />
                                            <input type="text" name="txtMontoTransaccionesMesAFEX" style="width: 80px; text-align: right;"
                                                onkeypress="IngresarTexto(1)" maxlength="18" />
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
                                    onkeypress="IngresarTexto(1)" maxlength="18" />
                            </td>
                        </tr>
                        <tr>
                            <td valign="top">
                                Monto Patrimonio USD<br />
                                <input type="text" name="txtMontoPatrimonio" style="width: 80px; text-align: right;"
                                    onkeypress="IngresarTexto(1)" maxlength="18" />
                            </td>
                            <td>
                                Origen Patrimonio y/o Fondos<br />
                                <input type="text" name="txtOrigenFondos" style="width: 359px; height: 50px;" onkeypress="IngresarTexto(3)"
                                    maxlength="200" />
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
        </table>
    </div>
    <div id="divBotonesCautela" style="display: none">
        <table border="0" cellpadding="0" cellspacing="0" style="height: 10px; width: 570;
            background-color: #f4f4f4">
            <tr align="right">
                <td >
                    &nbsp;
                </td>
                <td id="btnAceptarCautela" onclick>
                    <img align="absMiddle" src="../images/BotonAceptar.jpg" style="cursor: hand" width="70"
                        height="20">
                </td>
            </tr>
        </table>
    </div>
    <div id="divBotonesNormal" style="display: none">
        <table border="0" cellpadding="0" cellspacing="0" style="height: 10px; width: 570;
            background-color: #f4f4f4">
            <tr align="right">
                <td onclick="FichaCliente" id="btnprint">
                    Ficha de Cliente
                    <img align="absMiddle" src="../images/BotonImprimir.jpg" style="cursor: hand" width="70"
                        height="20">
                </td>
                <td id="cmdAceptar" onclick>
                    <img align="absMiddle" src="../images/BotonAceptar.jpg" style="cursor: hand" width="70"
                        height="20">
                </td>
            </tr>
        </table>
    </div>
    <input name="txtContacto1" type="hidden" value="">
    <input name="txtPorcentageContacto1" type="hidden" value="">
    <input name="txtContacto2" type="hidden" value="">
    <input name="txtPorcentageContacto2" type="hidden" value="">
    <input name="txtFechaActivacionComision" type="hidden" value="">
    <input name="txtSucursalSolicitante" type="hidden" value="">
    <input name="txtIDConsultada" type="hidden" value="<%=bIDConsultada%>">
    <input name="txtIDValida" type="hidden" value="<%=bIDValida%>">
    <input name="txtMensajeRegistro" type="hidden" value="<%=sMensajeRegistro%>">
    </form>

    <script language="vbscript">
	Dim afxConexion 
	afxConexion = "DSN=AfexCorporativa;UID=corporativa;PWD=afxsqlcor;"

	Sub window_onload()
		Dim i, nTipoCliente	
		
		'AMP 19-06-2015
		CargaSelRiesgo
		
		If Not IsNull("<%=sNivelRiesgo%>") And "<%=sNivelRiesgo%>" <> "" then
		    If("<%=sNivelRiesgo%>" = "CAUTELA") then
                divDatosCautela.style.display = ""
	            divBotonesCautela.style.display = ""
	            'frmCliente.txtRutCautela.value = "<%=sRut%>"
	            frmcliente.cbxRiesgo.selectedIndex=0
	            'frmcliente.txtNumSerieCautela.focus

                <%if not bClienteOtroLugar then%>
			        CargarDatosCautela
		        <%end if%>

                <%if bClienteOtroLugar then%>
			        msgbox "El Cliente no fue encontrado en Atención Clientes, pero sus datos se exportarán desde Giros para que lo pueda agregar.", , "AFEX"
				
		        <%end if%>

                If Not IsNull("<%=sRut%>") And "<%=sRut%>" <> "" And "<%=sRut%>" <> "null" Then
		            frmcliente.optPasCautela.checked = 0
		            frmcliente.optRutCautela.checked = 1
		            frmcliente.txtPasCautela.style.display="none"
		            frmcliente.txtPasCautela.value = ""
		            frmCliente.txtRutCautela.style.display = ""
		            frmCliente.txtNumSerieCautela.style.display = ""
		            tdNSerieCautela.style.display = ""
                    tdPaisPasaporteCautela.style.display="none"
                    frmCliente.txtRutCautela.value = "<%= request.form("txtRutCautela") %>"
		        End If
                If Not IsNull("<%=sPasaporte%>") And "<%=sPasaporte%>" <> "" And "<%=sPasaporte%>" <> "null" Then
			        frmcliente.optRutCautela.checked = 0
		            frmcliente.optPasCautela.checked = 1
		            frmCliente.txtRutCautela.style.display = "none"		
		            frmCliente.txtRutCautela.value = ""
		            frmcliente.txtPasCautela.style.display=""
		            frmcliente.txtNumSerieCautela.style.display = "none"
		            frmcliente.txtNumSerieCautela.value = ""
		            tdNSerieCautela.style.display = "none"
                    tdPaisPasaporteCautela.style.display=""
                    frmcliente.txtPasCautela.value = "<%= request.form("txtPasCautela") %>"
		        End If
            
	        Else
	            divDatosNormal.style.display = ""
	            divBotonesNormal.style.display = ""
	            frmcliente.cbxRiesgo.selectedIndex=1
                ' JFMG 04-09-2010 verifica si el cliente fue encontrado en otro lugar		
		        <%if not bClienteOtroLugar then%>
		        ' FIN JFMG 
			        CargarDatos
		        ' JFMG
		        <%end if%>
		        ' FIN JFMG

		        If CInt(0 & "<%=nTipoCliente%>") = 1 Then
			        trEmpresa.style.display = "none"
			        trContacto.style.display = "none"
			        trPersona.style.display = ""
			
		        ElseIf CInt(0 & "<%=nTipoCliente%>") > 1 Then
			        trEmpresa.style.display = ""
			        trContacto.style.display = "none"
			        trPersona.style.display = "none"	 
		        End If
		        if "<%=nTipoCliente%>" = "" or "<%=nTipoCliente%>" = "0" then
			        frmCliente.cbxTipo.value = 1
		        else
			        frmCliente.cbxTipo.value = "<%=nTipoCliente%>"
		        end if
		
		        If "<%=sPasaporte%>" = "" Then
			        optRut_onclick
			        frmCliente.txtRut.value = "<%=sRut%>"
		        Else
			        optPasaporte_onclick
			        frmCliente.txtPasaporte.value = "<%=sPasaporte%>"
		        End If
		
		        MostrarLaboral	
		
		        <%if trim(sMensajeRegistro) <> "" then%>
			        msgbox "<%=sMensajeRegistro%>",,"Consulta Registro Civil"
		        <%end if%>
		
		        ' JFMG 04-09-2010 verifica si el cliente fue encontrado en otro lugar		
		        <%if bClienteOtroLugar then%>
			        msgbox "El Cliente no fue encontrado en Atención Clientes, pero sus datos se exportarán desde Giros para que lo pueda agregar.", , "AFEX"
				
		        <%end if%>
		        ' FIN JFMG 04-09-2010
		
		        IF cint("0" & <%=Session("TipoOrigenLlamada")%>) = cint(1) Then
		            frmcliente.txtPropositoTransacciones.disabled = True
		            frmCliente.txtOrigenFondos.disabled = True
		            frmcliente.txtReferenciaBancaria.disabled = True
		            frmcliente.txtIngresoAnual.disabled = True
		            frmcliente.txtCantidadTransaccionesMesCLIENTE.disabled = True
		            frmcliente.txtCantidadTransaccionesMesAFEX.disabled = True
		            frmcliente.txtMontoTransaccionesMesCLIENTE.disabled = True
		            frmcliente.txtMontoTransaccionesMesAFEX.disabled = True
		        End If
	        End If
	    End If
		'FIN AMP 19-06-2015
	End Sub
	
	'AMP 19-06-2015
	Sub CargaSelRiesgo
        Dim MyOption 
        Set MyOption = Document.createElement("OPTION") 
        frmcliente.cbxRiesgo.Add MyOption 
        MyOption.innerText = "<-- Seleccione -->" 
        MyOption.Value = "0" 
        MyOption.selected = true
    End Sub
    

    Sub CargarDatosCautela
        
        '**************************************** CAUTELA ***********************************	
        frmCliente.txtRutCautela.value = "<%= request.form("txtRutCautela") %>"
        frmCliente.txtNumSerieCautela.value = "<%= request.form("txtNumSerieCautela") %>"
        frmCliente.txtPasCautela.value =  "<%= request.form("txtPasCautela") %>"
        frmCliente.txtApellidoPCautela.value =  "<%= request.form("txtApellidoPCautela") %>"
        frmCliente.txtApellidoMCautela.value =  "<%= request.form("txtApellidoMCautela") %>"
        frmCliente.txtNombresCautela.value =  "<%= request.form("txtNombresCautela") %>"
        '**************************************** Fin CAUTELA***********************************
    End Sub

    Sub LimpiaDatosCautela

        frmCliente.txtRutCautela.value = ""
        frmCliente.txtPasCautela.value = ""
        frmCliente.txtApellidoPCautela.value = ""
        frmCliente.txtApellidoMCautela.value = ""
        frmCliente.txtNombresCautela.value =  ""
        frmCliente.txtDescripcionCautela.value = ""

    End Sub

    Sub LimpiaDatosNormal

        frmCliente.txtRut.value = ""
        frmCliente.txtPasaporte.value = ""
        frmCliente.txtApellidoM.value = ""
		frmCliente.txtApellidoP.value = ""
		frmCliente.txtNombres.value = ""
		frmCliente.txtRazonSocial.value = ""
		frmCliente.txtGiroComercial.value = ""
		frmCliente.txtRepresentante.value =""
		frmCliente.txtCorreo.value = ""
		frmCliente.txtUsuario.value = ""
		frmCliente.txtPassword.value = ""
		frmCliente.txtConfPassword.value =""
		frmCliente.cbxTipo.value = 7
		frmCliente.txtApellidoMContacto.value = ""
		frmCliente.txtApellidoPContacto.value = ""
		frmCliente.txtNombresContacto.value = ""
		
		frmCliente.txtDireccionPersonal.value = ""
		frmCliente.txtNumero.value = ""
		frmCliente.txtdepto.value = ""
		frmCliente.txtPaisFonoPersonal.value = ""
		frmCliente.txtAreaFonoPersonal.value = ""
				
		frmCliente.txtFonoPersonal.value = ""
		frmCliente.txtCelular.value = ""
		frmCliente.txtPaisFaxPersonal.value = ""
		frmCliente.txtAreaFaxPersonal.value = ""
		frmCliente.txtFaxPersonal.value = ""
		frmCliente.txtNombreEmpresa.value = ""
		frmCliente.txtDireccionLaboral.value = ""
		frmCliente.txtPaisFonoLaboral.value = ""
		frmCliente.txtAreaFonolaboral.value = ""
		frmCliente.txtFonoLaboral.value = ""
				
		frmCliente.txtPaisFaxLaboral.value = ""
		frmCliente.txtAreaFaxlaboral.value = ""
		frmCliente.txtFaxLaboral.value = ""
		frmCliente.txtPregunta.value = ""
		frmCliente.txtRespuesta.value = ""
		frmCliente.txtFechaNacimiento.value = ""
		frmCliente.txtContacto1.value = ""
		frmCliente.txtPorcentageContacto1.value = ""
		frmCliente.txtContacto2.value = ""
		frmCliente.txtPorcentageContacto2.value =  ""
		frmCliente.txtFechaActivacionComision.value = ""
				
		frmCliente.txtnumeroserie.value = ""
		frmCliente.txtIDConsultada.value = ""
		frmCliente.txtIDValida.value = ""
		frmCliente.txtMensajeRegistro.value = ""
		
		frmcliente.txtPropositoTransacciones.value = ""
		frmCliente.txtOrigenFondos.value = ""
		frmcliente.txtReferenciaBancaria.value = ""
        frmcliente.txtMontoPatrimonio.value = ""
		frmcliente.txtIngresoAnual.value = ""
		frmcliente.txtCantidadTransaccionesMesCLIENTE.value = ""
		frmcliente.txtCantidadTransaccionesMesAFEX.value = ""
		frmcliente.txtMontoTransaccionesMesCLIENTE.value = ""
		frmcliente.txtMontoTransaccionesMesAFEX.value = ""
		frmcliente.txtActividadEconomica.value = ""
		frmcliente.txtNombreEmpleador.value = ""

        frmcliente.cbxSexo.value = 0
        frmcliente.cbxPaisPersonal.value = ""
        frmcliente.cbxCiudadPersonal.value = ""
        frmcliente.cbxComunaPersonal.value = ""
        frmcliente.cbxOcupacion.value = "" 'INTERNO-9263 MS 19-01-2017

    End Sub
    'FIN AMP 19-06-2015
	Sub CargarDatos
		frmCliente.txtApellidoM.value = "<%= request.form("txtApellidoM") %>"
		frmCliente.txtApellidoP.value = "<%= request.form("txtApellidoP") %>"
		frmCliente.txtNombres.value = "<%= request.form("txtNombres") %>"
		frmCliente.txtRazonSocial.value = "<%= request.form("txtRazonSocial") %>"
		frmCliente.txtGiroComercial.value = "<%= request.form("txtGiroComercial") %>"'JBV25-05-2012
		frmCliente.txtRepresentante.value = "<%= request.form("txtRepresentante") %>"
		frmCliente.txtCorreo.value = "<%= request.form("txtCorreo") %>"
		frmCliente.txtUsuario.value = "<%= request.form("txtUsuario") %>"
		frmCliente.txtPassword.value = "<%= request.form("txtPassword") %>"
		frmCliente.txtConfPassword.value = "<%= request.form("txtConfPassword") %>"
		' JFMG 19-10-2007 se agregó combo con mas tipos		
		frmCliente.cbxTipo.value = "<%= request.form("cbxTipo") %>"
		'**************************************** Fin ***********************************				
		frmCliente.optRut.value = "<%= request.form("optRut") %>"
		frmCliente.optPasaporte.value = "<%= request.form("optPasaporte") %>"
		frmCliente.txtApellidoMContacto.value = "<%= request.form("txtApellidoMContacto") %>"
		frmCliente.txtApellidoPContacto.value = "<%= request.form("txtApellidoPContacto") %>"
		frmCliente.txtNombresContacto.value = "<%= request.form("txtNombresContacto") %>"
		
		frmCliente.txtDireccionPersonal.value = "<%= request.form("txtDireccionPersonal") %>"
		frmCliente.txtNumero.value = "<%= request.form("txtnumero") %>"
		frmCliente.txtdepto.value = "<%= request.form("txtDepto") %>"
		
		'frmCliente.txtProfesion.value = "<%= request.form("txtProfesion") %>"
		
		frmCliente.txtPaisFonoPersonal.value = "<%=Buscarddi(1, sPais)%>" 
				
		frmCliente.txtAreaFonoPersonal.value = "<%=Buscarddi(2, sCiudad)%>"
				
		frmCliente.txtFonoPersonal.value = "<%= request.form("txtFonoPersonal") %>"
		frmCliente.txtCelular.value = "<%= request.form("txtCelular") %>"
		frmCliente.txtPaisFaxPersonal.value = frmCliente.txtPaisFonoPersonal.value
		frmCliente.txtAreaFaxPersonal.value = frmCliente.txtAreaFonoPersonal.value
		frmCliente.txtFaxPersonal.value = "<%= request.form("txtFaxPersonal") %>"
		frmCliente.txtNombreEmpresa.value = "<%= request.form("txtNombreEmpresa") %>"
		frmCliente.txtDireccionLaboral.value = "<%= request.form("txtDireccionLaboral") %>"
		frmCliente.txtPaisFonoLaboral.value = "<%=Buscarddi(1, sPais)%>" 
		frmCliente.txtAreaFonolaboral.value = "<%=Buscarddi(2, sCiudadL)%>"
		frmCliente.txtFonoLaboral.value = "<%= request.form("txtFonoLaboral") %>"
				
		frmCliente.txtPaisFaxLaboral.value = frmCliente.txtPaisFonoLaboral.value
		frmCliente.txtAreaFaxlaboral.value = frmCliente.txtAreaFonolaboral.value
		frmCliente.txtFaxLaboral.value = "<%= request.form("txtFaxLaboral") %>"
		frmCliente.txtPregunta.value = "<%= request.form("txtPregunta") %>"
		frmCliente.txtRespuesta.value = "<%= request.form("txtRespuesta") %>"
		frmCliente.txtFechaNacimiento.value = "<%= request.form("txtFechaNacimiento") %>"
		frmCliente.txtContacto1.value = "<%= request.form("txtContacto1") %>"
		frmCliente.txtPorcentageContacto1.value = "<%= request.form("txtPorcentageContacto1") %>"
		frmCliente.txtContacto2.value = "<%= request.form("txtContacto2") %>"
		frmCliente.txtPorcentageContacto2.value =  "<%= request.form("txtPorcentageContacto2") %>"
		frmCliente.txtFechaActivacionComision.value = "<%=request.form("txtFechaActivacionComision")%>"
				
		frmCliente.txtnumeroserie.value = "<%=request.form("txtNumeroSerie")%>"
		frmCliente.txtIDConsultada.value = "<%=bIDConsultada%>"
		frmCliente.txtIDValida.value = "<%=bIDValida%>"
		frmCliente.txtMensajeRegistro.value = "<%=sMensajeRegistro%>"
		
		frmcliente.txtPropositoTransacciones.value = "<%=request.form("txtPropositoTransacciones")%>"
		frmCliente.txtOrigenFondos.value = "<%=request.form("txtOrigenFondos")%>"
		frmcliente.txtReferenciaBancaria.value = "<%=request.form("txtReferenciaBancaria")%>"
        frmcliente.txtMontoPatrimonio.value = "<%=request.form("txtMontoPatrimonio")%>"
		frmcliente.txtIngresoAnual.value = "<%=request.form("txtIngresoAnual")%>"
		frmcliente.txtCantidadTransaccionesMesCLIENTE.value = "<%=request.form("txtCantidadTransaccionesMesCLIENTE")%>"
		frmcliente.txtCantidadTransaccionesMesAFEX.value = "<%=request.form("txtCantidadTransaccionesMesAFEX")%>"
		frmcliente.txtMontoTransaccionesMesCLIENTE.value = "<%=request.form("txtMontoTransaccionesMesCLIENTE")%>"
		frmcliente.txtMontoTransaccionesMesAFEX.value = "<%=request.form("txtMontoTransaccionesMesAFEX")%>"
		frmcliente.txtActividadEconomica.value = "<%=request.form("txtActividadEconomica")%>"
		frmcliente.txtNombreEmpleador.value = "<%=request.form("txtNombreEmpleador")%>"
		frmCliente.cbxOcupacion.value = "<%= request.form("cbxOcupacion") %>" 'INTERNO-9263 MS 19-01-2017
		CargarActividadEconomica
	End Sub
		
	Sub objTab_OnScriptletEvent(strEventName, varEventData)
	   Select Case strEventName
	   
			Case "linkClick"
				
				If varEventData <> sOldPaso Then

'	Nota: Estal lineas estan comentadas para no demorar la revision.
'	NO SE DEBEN ELIMINAR, por el contrario se deben quitar los comentarios
'	al momento de su revisión final.
						
					Select Case varEventData
						Case "tabPaso2"
							
							'If Not bPasoUno Then ValidarPasoUno
							'If Not bPasoUno Then MostrarPaso("tabPaso1"): Exit Sub
							'ValidarPasoUno
							'If Not bPasoUno Then MostrarPaso("tabPaso1"): Exit Sub
							
						Case "tabPaso3"						
							'If Not bPasoUno Then MostrarPaso("tabPaso1"): Exit Sub
							'If Not bPasoDos Then MostrarPaso(sOldPaso): Exit Sub
							'ValidarPasoUno
							'If Not bPasoUno Then MostrarPaso("tabPaso1"): Exit Sub
							'ValidarPasoDos
							'If Not bPasoDos Then MostrarPaso(sOldPaso): Exit Sub
							'If Not bPasoCuatro Then MostrarPaso(sOldPaso): Exit Sub
							
						'Case "tabPaso4"
							'If Not bPasoUno Then MostrarPaso("tabPaso1"): Exit Sub
							'If Not bPasoDos Then MostrarPaso("tabPaso2"): Exit Sub
							'If Not bPasoTres Then ValidarPasoTres
							'If Not bPasoTres Then MostrarPaso(sOldPaso): Exit Sub
							'ValidarPasoUno
							'If Not bPasoUno Then MostrarPaso("tabPaso1"): Exit Sub
							'If Not bPasoDos Then MostrarPaso("tabPaso2"): Exit Sub
							'ValidarPasoTres
							'If Not bPasoTres Then MostrarPaso(sOldPaso): Exit Sub
							
					End Select
				
					'document.all.item(sOldPaso).style.display = "none"
					'document.all.item(varEventData ).style.display = ""
					'sOldPaso = varEventData				 
					'MostrarPaso(varEventData)

				End If				
		End Select
		
	End Sub
	
	'AMP 22-06-2015
	Sub btnAceptarCautela_onClick()
	
	    sString = Empty
	    sRutC = Empty
	    sPasC = Empty
	    
	    If frmCliente.optRutCautela.checked Then
		    If frmCliente.txtRutCautela.value = Empty Then 
		        sString = "Rut, "
		        sRutC = "Null"
		    Else
		        sRutC = frmCliente.txtRutCautela.value
		    End If
		Else
			If frmCliente.txtPasCautela.value = Empty Then 
			    sString = "Pasaporte, "
			    sPasC = "Null"
		    Else
		        sPasC = frmCliente.txtPasCautela.value
                If frmCliente.cbxPaisPasaporteCautela.selectedIndex = -1 Then
			        sPaisPasaporteCautela = ""
                    MsgBox "Debe seleccionar el Pais del Pasaporte ", vbOKOnly + vbInformation, "Agregar Nuevo Cliente"
			        Exit Sub
		        Else
			        sPaisPasaporteCautela = frmCliente.cbxPaisPasaporteCautela(frmCliente.cbxPaisPasaporteCautela.selectedIndex).Text
		        End If
		    End If
		End If
		
		If frmCliente.txtNombresCautela.value = Empty Then sString = sString & "Nombres, "
		If frmCliente.txtApellidoPCautela.value = Empty Then sString = sString & "Apellido Paterno, "
		If frmCliente.txtApellidoMCautela.value = Empty Then sString = sString & "Apellido Materno, "
        If frmCliente.txtDescripcionCautela.value = Empty Then sString = sString & "Descripción, "
		
		If sString <> Empty Then
			If Right(Trim(sString), 1) = "," Then
				sString = Mid(Trim(sString), 1, Len(Trim(sString)) - 1)
			End If
			MsgBox "Debe ingresar los siguientes campos: " & sString, vbOKOnly + vbInformation, "Agregar Nuevo Cliente"
			Exit Sub
		else
		    nTipoCliente = CInt(0 & "<%=nTipoCliente%>")
		    If sPasC = Empty Then
		        sPasC = "null"
		    End If
		    
		    sNumeroRut = replace(replace((trim(sRutC)),"-",""),".","")

            frmCliente.action = "http:GuardarNuevoClienteCautela.asp?TC=" & nTipoCliente & "&RutCautela=" & sNumeroRut & "&NserieCautela=" & frmCliente.txtNumSerieCautela.value & "&PasCautela=" & sPasC & "&NomCautela=" & frmCliente.txtNombresCautela.value & "&PatCautela=" & frmCliente.txtApellidoPCautela.value & "&MatCautela=" & frmCliente.txtApellidoMCautela.value & "&DescCautela=" & frmCliente.txtDescripcionCautela.value & "&PaisPas=" &sPaisPasaporteCautela
		    frmCliente.submit
		    frmCliente.action = ""
		End If
	End Sub
	'FIN AMP 22-06-2015
	
	Sub cmdAceptar_onClick()
		Dim nTipoCliente, nClienteAgencia
		
		AsignarActividadEconomica
				
		If trLaborales1.style.display = "none" Then
			frmCliente.txtNombreEmpresa.value = ""
			frmCliente.txtDireccionLaboral.value = ""
			frmCliente.cbxCiudadLaboral.value = ""
			frmCliente.cbxComunaLaboral.value = ""
			frmCliente.txtPaisFaxLaboral.value = ""
			frmCliente.txtAreaFaxLaboral.value = ""
			frmCliente.txtFaxLaboral.value = ""
			frmCliente.txtPaisFonoLaboral.value = ""
			frmCliente.txtAreaFonoLaboral.value = ""
			frmCliente.txtFonoLaboral.value = ""
			
		End If
		
		If Not ValidarInformacion Then Exit Sub
		
		' JFMG 19-10-2007 se agregó un combo con mas tipos
		nTipoCliente = frmCliente.cbxTipo.value
		nClienteAgencia = frmCliente.chkCLienteAgencia.checked
		'**************************** Fin ****************************
		
		HabilitarCampos
		
		' JFMG 15-06-2011 desde hoy ya no se consultará más la sucursal solicitante
		' solicita la ip de la sucursal para ir a la bd y verificar al nuevo cliente
		'frmCliente.txtSucursalSolicitante.value  = window.showModalDialog("sucursalsolicitante.asp")		
		' FIN JFMG 15-06-2011
		
		frmCliente.action = "GuardarNuevoCliente.asp?TCl=" & nTipoCliente & "&ConsultaID=" & "<%=bIdCorrecto%>" & "&CA=" & nClienteAgencia 
		frmCliente.submit
		frmCliente.action = ""
	End Sub

	Sub HabilitarCampos()
		frmCliente.txtPaisFonoPersonal.disabled = False
		frmCliente.txtAreaFonoPersonal.disabled = False
		frmCliente.txtPaisFaxPersonal.disabled = False
		frmCliente.txtAreaFaxPersonal.disabled = False
	End Sub
	
	Function ValidarInformacion
		Dim sString
		
		ValidarInformacion = False
		
		sString = Empty

		If frmCliente.optRut.checked Then
			If frmCliente.txtRut.value = Empty Then sString = "Rut, "
		Else
			If frmCliente.txtPasaporte.value = Empty Then sString = "Pasaporte, "
			If frmCliente.cbxPaisPasaporte.value = Empty Then sString = "País del Pasaporte, "
		End If

		' JFMG 30-09-2010 si el llamado es desde giros no se validan estos datos
		if cint("0" & <%=Session("TipoOrigenLlamada")%>) <> cint(1) then
		' FIN JFMG 30-09-2010

			If frmCliente.cbxTipo.value = 1 then 'frmCliente.optPersona.checked Then		'Si selecciono Persona
				If frmCliente.txtNombres.value = Empty Then sString = sString & "Nombres, "
				If frmCliente.txtApellidoP.value = Empty Then sString = sString & "Apellido Paterno, "
				If frmCliente.txtApellidoM.value = Empty Then sString = sString & "Apellido Materno, "
				If frmCliente.txtFechaNacimiento.value = Empty Then sString = sString & "Fecha Nacimiento, "
				If frmCliente.cbxSexo.value = "0" Then sString = sString & "Sexo, "
                If frmCliente.cbxOcupacion.value = "0" or frmCliente.cbxOcupacion.value = Empty Then sString = sString & "Ocupación, " 'INTERNO-9263 MS 19-01-2017
                If frmCliente.cbxNacionalidad.value = "0" or frmCliente.cbxNacionalidad.value = Empty Then sString = sString & "Nacionalidad, "
			Else		'Si seleccionó empresa
				If frmCliente.txtRazonSocial.value = Empty Then sString = sString & "Razón Social, "
				If frmCliente.txtRepresentante.value = Empty Then sString = sString & "Representante Legal, "
			End If		
		
		    If frmCliente.txtDireccionPersonal.value = Empty Then sString = sString & "Dirección, "
		    If frmCliente.txtNumero.value = empty then sString = sstring & "Número dirección, "
		    If frmCliente.cbxPaisPersonal.value = Empty Then sString = sString & "País, "
		    If frmCliente.cbxCiudadPersonal.value = Empty Then sString = sString & "Ciudad, "
		    If UCase(frmCliente.cbxPaisPersonal.value) = "CL" Then
			    If frmCliente.cbxComunaPersonal.value = Empty Then sString = sString & "Comuna, "
		    End If
		    If frmCliente.txtCelular.value = Empty and frmCliente.txtFonoPersonal.value = Empty then sString = sString & "Teléfono, "
		    
		    ' MS 28-03-2014
		    If frmCliente.optActEconomicaSi.checked and frmCliente.txtActividadEconomica.value = empty then sString = sString & "Actividad Económica, "
		    ' FIN MS 28-03-2014
		
		' JFMG 30-09-2010
		end if
		' FIN JFMG 30-09-2010
		
		' valida que agregue los contactos

		If sString <> Empty Then
			If Right(Trim(sString), 1) = "," Then
				sString = Mid(Trim(sString), 1, Len(Trim(sString)) - 1)
			End If
			
			if frmCliente.txtPassword.value <> frmCliente.txtConfPassword.value then
				sString=""
				MsgBox "Campos de Password no coinciden " & sString, vbOKOnly + vbInformation, "Agregar Nuevo Cliente"
			else 
				MsgBox "Debe ingresar los siguientes campos: " & sString, vbOKOnly + vbInformation, "Agregar Nuevo Cliente"
			end if	
			Exit Function
		else
			if frmCliente.txtPassword.value <> frmCliente.txtConfPassword.value then
				sString=""
				MsgBox "Campos de Password no coinciden " & sString, vbOKOnly + vbInformation, "Agregar Nuevo Cliente"
				Exit Function
			end if 
			
		End If

		' valida si desea consultar la identificacion en el registro civil		
		if frmCliente.txtIDConsultada.value = false then
			if msgbox("Recuerde consultar la Identificación en el Registro Civil, ¿Desea realizar la consulta?",1,"AFEX") = 1 then
				frmCliente.txtnumeroserie.focus
				frmCliente.txtnumeroserie.select
				exit function
			end if
		elseif frmCliente.txtIDValida.value = false then
			msgbox "La identificación presenta problemas en el Registro Civil, por lo tanto el Cliente no será Grabado."
			exit function
		end if
		
        'APPL-5076 MS 25-04-2014: Si falla la validación del evento onBlur del rut, se realiza una doble validación.
        If frmCliente.optRut.checked Then
			sRut = ValidarRut(frmCliente.txtRut.value)
			If sRut = "" Then
			    msgbox "Debe ingresar un número de rut válido ", vbOKOnly + vbInformation, "Ingreso de Cliente"
			    frmCliente.txtRut.focus
			    Exit function
		    End If
		end if
        'FIN APPL-5076 MS 25-04-2014
       
		ValidarInformacion = True
	End Function
	
	sub cmdContacto_onClick()
		dim sContacto
		dim i
		
		sContacto = window.showmodaldialog("EmpleadoContacto.asp?contacto1=" & frmCliente.txtContacto1.value & _
													"&porcentagecontacto1=" & frmCliente.txtPorcentageContacto1.value & _
													"&contacto2=" & frmCliente.txtContacto2.value & _
													"&porcentagecontacto2=" & frmCliente.txtPorcentageContacto2.value & _
													"&fechaactivacion=" & frmCliente.txtFechaActivacionComision.value)
		if  trim(sContacto) = empty then
			exit sub
		elseif len(trim(sContacto)) = 4 then 
			frmCliente.txtContacto1.value = ""
			frmCliente.txtporcentageContacto1.value = ""
			frmCliente.txtContacto2.value = ""
			frmCliente.txtporcentageContacto2.value = ""
			frmCliente.txtfechaactivacioncomision.value = ""		
		else		
			i = instr(sContacto, ";")
			frmCliente.txtContacto1.value = left(sContacto, i - 1)
			sContacto = mid(scontacto, i + 1)
			i = instr(sContacto, ";")
			frmCliente.txtporcentageContacto1.value = left(sContacto, i - 1)
			sContacto = mid(scontacto, i + 1)
			i = instr(sContacto, ";")
			frmCliente.txtContacto2.value = left(sContacto, i - 1)
			sContacto = mid(scontacto, i + 1)
			i = instr(sContacto, ";")
			frmCliente.txtporcentageContacto2.value = left(sContacto, i - 1)
			sContacto = mid(scontacto, i + 1)			
			frmCliente.txtfechaactivacioncomision.value = sContacto
		end if
	end sub
	
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
                    window.tblActividadEconomica.deleteRow i
                    frmcliente.txtActividadEconomica.value = replace(trim(frmcliente.txtActividadEconomica.value), trim(Codigo) & ";", "")
                    exit for
                end if
            next            
        end sub        
        
        function VerificarActividadEconomicaExiste(ByVal CodigoActividad)
            
            VerificarActividadEconomicaExiste = False
            
            if instr(frmcliente.txtActividadEconomica.value, CodigoActividad) > 0 then
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
            dim sActividadEconomica
            sActividadEconomica = split(frmcliente.txtActividadEconomica.value, ";")
                               
            for i = 0 to UBOUND(sActividadEconomica) - 1
                if sActividadEconomica(i) <> "" then
                    frmcliente.cbxActividadEconomicaOculta.value = sActividadEconomica(i)
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
        
        Sub txtRutCautela_onBlur()
            If frmCliente.txtRutCautela.value = Empty Then Exit Sub
            If (ValidaRut(frmcliente.txtRutCautela.value)) then
                frmCliente.action = "http:NuevoCliente.asp?Rut=" + frmCliente.txtRutCautela.value + "&NivelRiesgo=CAUTELA"
                frmCliente.submit()
                frmCliente.action = ""
            End If
        End Sub

        Sub txtPasCautela_onBlur()
            frmCliente.action = "http:NuevoCliente.asp?Pasaporte=" + frmCliente.txtPasCautela.value + "&NivelRiesgo=CAUTELA"
            frmCliente.submit()
            frmCliente.action = ""
        End Sub
        
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
        
    </script>

</body>
</html>
