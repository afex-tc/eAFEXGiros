<%@ LANGUAGE = VBScript %>
<!--#INCLUDE virtual="/Compartido/Errores.asp" -->
<!--#INCLUDE virtual="/Compartido/Rutinas.asp" -->
<%	
Dim CodigoClienteXc  
Dim CodigoClienteXp  
Dim sCodigo
Dim nTipoCliente

nTipoCliente = CInt(0 & Request("TCl"))

AgregarCliente
EnviarCorreo

Private Sub AgregarCliente
	Dim afxCliente
	
	On Error	Resume Next
	Set afxCliente = Server.CreateObject("AfexCorporativo.Cliente")
	If Err.number <> 0 Then
		'response.Redirect "http://200.72.160.51/afexmoneyweb/Compartido/Error.asp?Titulo=Error en HágaseCliente&Number=" & Err.number   & "&Source=" & Err.Source & "&Description=" & Err.description
		response.Redirect "http://www.afex.cl/Compartido/Error.asp?Titulo=Error en HágaseCliente&Number=" & Err.number   & "&Source=" & Err.Source & "&Description=" & Err.description
	End If
	On Error	Resume Next
	Dim TipoCliente
	Dim nEnvioGiro, nRecepcionGiro, nInformeGiro
	Dim nEnvioTransfer, nInformeTransfer
	Dim nCompraVenta, nAlarmas, nNoticias
	Dim bContacto

	
	If Request.Form("chkEnvioGiro") = "on" Then
		nEnvioGiro = 1
	Else
		nEnvioGiro = 0
	End If
	If Request.Form("chkRecepcionGiro") = "on" Then
		nRecepcionGiro = 1
	Else
		nRecepcionGiro = 0
	End If
	If Request.Form("chkInformeGiro") = "on" Then
		nInformeGiro = 1
	Else
		nInformeGiro = 0
	End If
	If Request.Form("chkEnvioTransfer") = "on" Then
		nEnvioTransfer = 1
	Else
		nEnvioTransfer = 0
	End If
	If Request.Form("chkInformeTransfer") = "on" Then
		nInformeTransfer = 1
	Else
		nInformeTransfer = 0
	End If
	If Request.Form("chkAlarmas") = "on" Then
		nAlarmas = 1
	Else
		nAlarmas = 0
	End If
	If Request.Form("chkNoticias") = "on" Then
		nNoticias = 1
	Else
		nNoticias = 0
	End If
	If Request.Form("chkCompraVenta") = "on" Then
		nCompraVenta = 1
	Else
		nCompraVenta = 0
	End If
	
	'If Request.Form("optPersona") = "1" Then
	If nTipoCliente = 1 Then
		TipoCliente = 1
		bContacto = False
	Else
		TipoCliente = 2
		bContacto = True
	End If
	'response.Redirect "Compartido/Error.asp?description=" & Session("afxCnxCorporativa") & ", " & "WB" & ", " &  _
	'						 Request.Form("txtRut") & ", " &  Request.Form("txtPasaporte") & ", " &  TipoCliente
 	'Session("afxCnxCorporativa") = "DSN=wAfexCorporativa;UID=corporativa;PWD=afxsqlcor;"
	If Err.number <> 0 Then
		'response.Redirect "http://200.72.160.51/afexmoneyweb/compartido/Error.asp?Titulo=Error en HágaseCliente 0&Number=" & Err.Number  & "&Source=" & Err.Source & "&Description=" & Err.Description
		response.Redirect "http://www.afex.cl/compartido/Error.asp?Titulo=Error en HágaseCliente 0&Number=" & Err.Number  & "&Source=" & Err.Source & "&Description=" & Err.Description
	End If
						 
	sCodigo = Agregar(Session("afxCnxCorporativa"), "WB", _
							 Request.Form("txtRut"), Request.Form("txtPasaporte"), TipoCliente, _
							 Request.Form("txtApellidoP"), Request.Form("txtApellidoM"), _
							 Request.Form("txtNombres"), "Null", _ 
							 request.Form("txtDireccionPersonal"), _							 
							 request.Form("cbxPaisPersonal"), _
							 request.Form("cbxCiudadPersonal"), request.Form("cbxComunaPersonal"), _
							 CInt(0 & request.Form("txtPaisFonoPersonal")), CInt(0 & request.Form("txtAreaFonoPersonal")), CCur(0 & request.Form("txtFonoPersonal")), _
							 CInt(0 & request.Form("txtPaisFaxPersonal")), CInt(0 & request.Form("txtAreaFaxPersonal")), CCur(0 & request.Form("txtFaxPersonal")), "", _
							 Request.Form("txtCorreo"), Request.Form("txtUsuario"), Request.Form("txtPassword"), _
							 Trim(Request.Form("txtRazonSocial") & Request.Form("txtNombreEmpresa")), _
							 request.Form("txtDireccionLaboral"), Request.Form("cbxCiudadLaboral"), _
							 Request.Form("cbxComunaLaboral"), _
							 CInt(0 & request.Form("txtPaisFonoLaboral")), cInt(0 & request.Form("txtAreaFonoLaboral")), CCur(0 & request.Form("txtFonoLaboral")), _
							 CInt(0 & request.Form("txtPaisFaxLaboral")), cInt(0 & request.Form("txtAreaFaxLaboral")), CCUr(0 & request.Form("txtFaxLaboral")), _
							 "", "", bContacto, "", "",  _
							 Request.Form("txtApellidoPContacto"), Request.Form("txtApellidoMContacto"), _
							 Request.Form("txtNombresContacto"), "", nEnvioGiro, nRecepcionGiro, nInformeGiro, _
							 nEnvioTransfer, nInformeTransfer, nCompraVenta, nAlarmas, nNoticias, 0, 0)
 	'Session("afxCnxCorporativa") = "DSN=AfexCorporativa;UID=corporativa;PWD=afxsqlcor;"

	If Err.number <> 0 Then
		'response.Redirect "http://200.72.160.51/afexmoneyweb/compartido/Error.asp?Titulo=Error en HágaseCliente 1&Number=" & Err.Number  & "&Source=" & Err.Source & "&Description=" & Err.Description
		response.Redirect "http://www.afex.cl/compartido/Error.asp?Titulo=Error en HágaseCliente 1&Number=" & Err.Number  & "&Source=" & Err.Source & "&Description=" & Err.Description
	End If
	If afxCliente.ErrNumber <> 0 Then
		'response.Redirect "http://200.72.160.51/afexmoneyweb/compartido/Error.asp?Titulo=Error en HágaseCliente 2&Number=" & afxCliente.ErrNumber  & "&Source=" & afxCliente.ErrSource & "&Description=" & replace(afxCliente.ErrDescription, vbCrLf , "^")
		response.Redirect "http://www.afex.cl/compartido/Error.asp?Titulo=Error en HágaseCliente 2&Number=" & afxCliente.ErrNumber  & "&Source=" & afxCliente.ErrSource & "&Description=" & replace(afxCliente.ErrDescription, vbCrLf , "^")
	End If			
	Set afxCliente = Nothing
End Sub	

Private Sub EnviarCorreo
	Dim objEMail, sNombre, sRazonSocial, sContacto, sDireccion
	Dim hora, fecha
	hora= replace(time,":","")
	fecha= replace(date,"/","")
	fecha= replace(fecha,"-","")
	'On Error Resume Next
	On Error Goto 0
	Set objEMail = Server.CreateObject("CDONTS.NewMail")

	sNombre = Trim(Trim(Request.Form("txtNombres")) & " " & Trim(Request.Form("txtApellidoP")) & " " & Trim(Request.Form("txtApellidoM")))
	sRazonSocial = Trim(Request.Form("txtRazonSocial"))
	sContacto = Trim(Trim(Request.Form("txtNombresContacto")) & " " & Trim(Request.Form("txtApellidoPContacto")) & " " & Trim(Request.Form("txtApellidoMContacto")))
	
							 
	objEMail.From = MayusculaMinuscula(sNombre) & " <" & Request.Form("txtCorreo") & ">"
	'objEmail.To ="roxana.mercado@afex.cl"
	objEMail.To = "silvia.venegas@afex.cl;laura.greene@afex.cl"	
	objEMail.cc = "julio.greene@afex.cl"			
	objEMail.Subject = "AFEX Ltda.   " &"(HC"& fecha & hora &")" 
	objEMail.Body = "Se ha creado un nuevo cliente desde la página web AFEXweb y sus datos son: " & Chr(13) & Chr(13) & _
						"Nombre		: " & sNombre & Chr(13) & _
						"Razon Social: " & sRazonSocial & Chr(13) & _
						"Identificacion:" & Request.Form("txtRut") & Request.Form("txtPasaporte") & Chr(13) & _
						"Teléfono	: (" & Request.Form("txtPaisFonoPersonal") & Request.Form("txtAreaFonoPersonal") & ") " & Request.Form("txtFonoPersonal") & Chr(13) & _
						"Contacto	: " & sContacto & Chr(13) & _
						"email		: " & Request.Form("txtCorreo") & chr(13) & _
						"Direccion	: " & Request.Form("txtDireccionPersonal") & ", " & Request.Form("cbxPaisPersonal") & ", " & Request.Form("cbxCiudadPersonal") & ", " & Request.Form("cbxComunaPersonal") & Chr(13) &  _
						"Empresa	: " & Trim(Request.Form("txtNombreEmpresa")) & chr(13) & _
						"Teléfono	: (" & Request.Form("txtPaisFonoLaboral") & Request.Form("txtAreaFonoLaboral") & ") " & Request.Form("txtFonoLaboral") & Chr(13) & _
						"Corporativo: " & sCodigo & Chr(13) & _
						"AFEXchange: " & CodigoClienteXc & Chr(13) & _
						"AFEXpress: " & CodigoClienteXp

	objEMail.Send 
	Set objEMail = Nothing

	If Err.number = 0 then
		'If Session("CodigoAgente") <> "" Then
		'	Response.Redirect "http:Agente\AtencionClientes.asp"
		'	Response.End 
		'Else
			Response.Redirect "http:Principal.asp"
		'End If
	End If

	'response.Redirect "Principal.asp"
End Sub


Function Agregar(ByVal Conexion, ByVal Sucursal,  ByVal Rut,  ByVal Pasaporte, _
                ByVal Tipo, ByVal ApellidoPaterno,  ByVal ApellidoMaterno, _
                ByVal Nombres,  ByVal FechaNacimiento, _
                ByVal DireccionParticular,  ByVal PaisParticular, _
                ByVal CiudadParticular,  ByVal ComunaParticular, _
                ByVal PaisFonoParticular,  ByVal AreaFonoParticular,  ByVal NumeroFonoParticular, _
                ByVal PaisFaxParticular,  ByVal AreaFaxParticular,  ByVal NumeroFaxParticular, _
                ByVal Celular,  ByVal CorreoElectronicoCliente, _
                ByVal NombreUsuario,  ByVal Password, _
                ByVal RazonSocial, _
                ByVal DireccionComercial,  ByVal CiudadComercial, _
                ByVal ComunaComercial, _
                ByVal PaisFonoComercial,  ByVal AreaFonoComercial,  ByVal NumeroFonoComercial, _
                ByVal PaisFaxComercial,  ByVal AreaFaxComercial,  ByVal NumeroFaxComercial, _
                ByVal CargoCliente,  ByVal ProfesionCliente, _
                ByVal Contacto,  ByVal CargoContacto,  ByVal RutContacto, _
                ByVal ApellidoPaternoContacto,  ByVal ApellidoMaternoContacto,  ByVal NombresContacto, _
                ByVal CorreoElectronicoContacto, _
                ByVal EnvioGiro,  ByVal RecepcionGiro, _
                ByVal InformeGiro,  ByVal EnvioTransferencia,  ByVal InformeTransferencia, _
                ByVal CompraVenta,  ByVal Alarma,  ByVal Noticias, _
                ByVal ContratoProducto,  ByVal Auditoria)

   Dim sSQL
   Dim BD
   Dim rsCodigoCliente
   Dim Mensaje
   Dim NombreCompleto 
   Dim objContext
   Dim RutCliente
   Dim Raya
   Dim sApellidos
   
   'Variables para ClienteXc
   Dim afxClienteXc     
   Dim ConexionXc       
   '*******************************************
   'Variables para ClienteXp
   Dim afxClienteXp     
   Dim ConexionXp       
   '*******************************************
   
   Agregar = 0
   
   ' controla los errores
   On Error Resume Next
   
   
   'Setear variables para AFEXchange   
   'If MTS Then Set objContext = GetObjectContext()
   'If MTS Then Set afxClienteXc = objContext.CreateInstance("AFEXClienteXc.Cliente")
   Set afxClienteXc = server.CreateObject("AFEXClienteXc.Cliente")
   
   'Setear variables para AFEXpress
   'If MTS Then Set objContext = GetObjectContext()
   'If MTS Then Set afxClienteXp = objContext.CreateInstance("AFEXClienteXp.Cliente")
   'If Not MTS Then Set afxClienteXp = New AFEXClienteXp.Cliente
   Set afxClienteXp = server.CreateObject("AFEXClienteXp.Cliente")
   
   'Leer vista de la base de datos corporativa
   sSQL = "SELECT exchange, express FROM vatencionclienteprueba " & _
          "WHERE 1 = 1 "
   
   If Rut <> Empty Then
		Rut = ValorRut(Rut)
      sSQL = sSQL & " And Rut = '" & Rut & "' "
      'sSQL = sSQL & "AND rut = '" & ValorRut(Rut) & "' "
   End If
   If Pasaporte <> Empty Then
      sSQL = sSQL & "AND pasaporte = '" & Pasaporte & "' "
   End If
   
   Set rsCodigoCliente = EjecutarSQLCliente(Conexion, sSQL)
	
   If Err.Number <> 0 Then 
		Set BD = Nothing
		Set rsCodigoCliente = Nothing
		Set afxClienteXc = Nothing
		Set afxClienteXp = Nothing
		'Response.Redirect "http://200.72.160.51/afexmoneyweb/compartido/error.asp?description=" &  "1"
		Response.Redirect "http://www.afex.cl/compartido/error.asp?description=" &  "1"
	End If

   If Rut = "" Then
      Rut = "Null"
   'Else
   '   Rut = EvaluarStr(Rut)
   End If
   If Err.Number <> 0 Then 
		Set BD = Nothing
		Set rsCodigoCliente = Nothing
		Set afxClienteXc = Nothing
		Set afxClienteXp = Nothing
		'Response.Redirect "http://200.72.160.51/afexmoneyweb/compartido/error.asp?description=" &  "2"
		Response.Redirect "http://www.afex.cl/compartido/error.asp?description=" &  "2"
	End If
   
   If rsCodigoCliente.EOF Then
      CodigoClienteXc = Empty
      CodigoClienteXp = Empty
   Else
		If IsNull(rsCodigoCliente("exchange")) Then
			CodigoClienteXc = ""
		Else
			CodigoClienteXc = rsCodigoCliente("exchange")
		End If
      If IsNull(rsCodigoCliente("express")) Then
			CodigoClienteXp = ""
		Else
			CodigoClienteXp = rsCodigoCliente("express")
		End If
   End If
   If Err.Number <> 0 Then 
		Set BD = Nothing
		Set rsCodigoCliente = Nothing
		Set afxClienteXc = Nothing
		Set afxClienteXp = Nothing
	  'Response.Redirect "http://200.72.160.51/afexmoneyweb/compartido/error.asp?description=" &  "3"
	  Response.Redirect "http://www.afex.cl/compartido/error.asp?description=" &  "3"
   End If
  
   NombreCompleto = Trim(Nombres) & " " & Trim(ApellidoPaterno) & " " & Trim(ApellidoMaterno)
   
   sSQL = "InsertarCliente " & EvaluarStr(Rut) & ", " & EvaluarStr(Pasaporte) & ", " & Tipo & ", " & _
                               EvaluarStr(NombreUsuario) & "," & EvaluarStr(Password) & "," & EvaluarStr(NombreCompleto) & ", " & _
                               EnvioGiro & ", " & RecepcionGiro & ", " & InformeGiro & ", " & _
                               EnvioTransferencia & ", " & InformeTransferencia & ", " & CompraVenta & ", " & _
                               Alarma & ", " & Noticias & ", " & ContratoProducto & ", " & _
                               Auditoria & ", " & EvaluarStr(CodigoClienteXp) & ", " & EvaluarStr(CodigoClienteXc)
	
   'Conexion
   Set BD = Server.CreateObject("ADODB.Connection")
   BD.Open Conexion                          'Abre la conexion
   If Err.Number <> 0 Then 
		Set BD = Nothing
		Set rsCodigoCliente = Nothing
		Set afxClienteXc = Nothing
		Set afxClienteXp = Nothing
	  'Response.Redirect "http://200.72.160.51/afexmoneyweb/compartido/error.asp?description=" &  "4"
	  Response.Redirect "http://www.afex.cl/compartido/error.asp?description=" &  "4"
   End If
   
   'Consulta
   BD.Execute sSQL                           'Ejecuta la consulta
    If Err.Number <> 0 Then 
		Set BD = Nothing
		Set rsCodigoCliente = Nothing
		Set afxClienteXc = Nothing
		Set afxClienteXp = Nothing
		'Response.Redirect "http://200.72.160.51/afexmoneyweb/compartido/error.asp?description=" &  Err.number & "<br>" & sSQL
		Response.Redirect "http://www.afex.cl/compartido/error.asp?description=" &  Err.number & "<br>" & sSQL
    End If
	
   If CodigoClienteXc = Empty Then
      'Obtener codigo de la bd corporativa
      sSQL = "Select codigo, codigo_afexchange " & _
             "from cliente with(nolock) " & _
             "where codigo in (select max(codigo) as codigo from cliente)"
   
      Set rsCodigoCliente = EjecutarSQLCliente(Conexion, sSQL)
   
		If Err.Number <> 0 Then 
			Set BD = Nothing
			Set rsCodigoCliente = Nothing
			Set afxClienteXc = Nothing
			Set afxClienteXp = Nothing
			'Response.Redirect "http://200.72.160.51/afexmoneyweb/compartido/error.asp?description=" &  "6"
			Response.Redirect "http://www.afex.cl/compartido/error.asp?description=" &  "6"
		End If
   
      If rsCodigoCliente.EOF Then
			Set BD = Nothing
			Set rsCodigoCliente = Nothing
			Set afxClienteXc = Nothing
			Set afxClienteXp = Nothing
         'Response.Redirect "http://200.72.160.51/afexmoneyweb/compartido/error.asp?description=" &  "7"
         Response.Redirect "http://www.afex.cl/compartido/error.asp?description=" &  "7"
      Else
         CodigoClienteXc = rsCodigoCliente("codigo_afexchange")
      End If
      
      If Err.Number <> 0 Then 
			Set BD = Nothing
			Set rsCodigoCliente = Nothing
			Set afxClienteXc = Nothing
			Set afxClienteXp = Nothing
			'Response.Redirect "http://200.72.160.51/afexmoneyweb/compartido/error.asp?titulo=Agregar Cliente 8&description=" &  Err.description
			Response.Redirect "http://www.afex.cl/compartido/error.asp?titulo=Agregar Cliente 8&description=" &  Err.description
		End If
      'Ingresar cliente en el afexchange
      '**********************************************************
      'ConexionXc = "DSN=wafexchange;UID=cambios;PWD=cambios;"
      ConexionXc = Session("afxCnxAFEXchange")
		RutContacto = ValorRut(RutContacto)
      If Err.Number <> 0 Then 
			Set BD = Nothing
			Set rsCodigoCliente = Nothing
			Set afxClienteXc = Nothing
			Set afxClienteXp = Nothing
			'Response.Redirect "http://200.72.160.51/afexmoneyweb/compartido/error.asp?titulo=Agregar Cliente 9&description=" &  Err.description
			Response.Redirect "http://www.afex.cl/compartido/error.asp?titulo=Agregar Cliente 9&description=" &  Err.description
		End If
      afxClienteXc.Agregar ConexionXc, CodigoClienteXc, Sucursal, Rut, _
                           Pasaporte, PaisParticular, cInt(Tipo), _
                           ApellidoPaterno, _
                           ApellidoMaterno, Nombres, NombreCompleto, _
                           RazonSocial, _
                           DireccionParticular, ComunaParticular, _
                           CiudadParticular, cInt(PaisFonoParticular), _
                           cInt(AreaFonoParticular), cCur(NumeroFonoParticular), _
                           cInt(PaisFaxParticular), cInt(AreaFaxParticular), _
                           cCur(NumeroFaxParticular), DireccionComercial, _
                           cInt(PaisFonoComercial), cInt(AreaFonoComercial), _
                           cCur(NumeroFonoComercial), cInt(PaisFaxComercial), _
				               cInt(AreaFaxComercial), cCur(NumeroFaxComercial), _
                           ComunaComercial, CiudadComercial, _
                           Celular, CorreoElectronicoCliente, CargoCliente, _
                           ProfesionCliente, Contacto, CargoContacto, RutContacto, _
                           ApellidoPaternoContacto, ApellidoMaternoContacto, _
                           NombresContacto, CorreoElectronicoContacto
   
      If Err.Number <> 0 Then 
			Set BD = Nothing
			Set rsCodigoCliente = Nothing
			Set afxClienteXc = Nothing
			Set afxClienteXp = Nothing
			'Response.Redirect "http://200.72.160.51/afexmoneyweb/compartido/error.asp?titulo=Agregar Cliente 10&description=" &  Err.description
			Response.Redirect "http://www.afex.cl/compartido/error.asp?titulo=Agregar Cliente 10&description=" &  Err.description  
		End If
		'Response.Redirect "http:compartido/error.asp?titulo=Agregar Cliente 100&description=" &  Err.description
      If afxClienteXc.ErrNumber <> 0 Then 
			Set BD = Nothing
			Set rsCodigoCliente = Nothing
			Set afxClienteXc = Nothing
			Set afxClienteXp = Nothing
			'Response.Redirect "http://200.72.160.51/afexmoneyweb/compartido/error.asp?description=" & "11"
			Response.Redirect "http://www.afex.cl/compartido/error.asp?description=" & "11"
		End If

      '************************************************************
   End If

   If CodigoClienteXp = Empty Then
      'Obtener codigo de la bd corporativa
      sSQL = "Select codigo, codigo_afexpress " & _
             "from cliente with(nolock) " & _
             "where codigo in (select max(codigo) as codigo from cliente)"
   
      Set rsCodigoCliente = EjecutarSQLCliente(Conexion, sSQL)
      If Err.Number <> 0 Then 
			Set BD = Nothing
			Set rsCodigoCliente = Nothing
			Set afxClienteXc = Nothing
			Set afxClienteXp = Nothing
			'Response.Redirect "http://200.72.160.51/afexmoneyweb/compartido/error.asp?titulo=Agregar Cliente 12&description=" &  Err.description
			Response.Redirect "http://www.afex.cl/compartido/error.asp?titulo=Agregar Cliente 12&description=" &  Err.description
		End If
   
      If Err.Number <> 0 Then 
			Set BD = Nothing
			Set rsCodigoCliente = Nothing
			Set afxClienteXc = Nothing
			Set afxClienteXp = Nothing
			'Response.Redirect "http://200.72.160.51/afexmoneyweb/compartido/error.asp?description=" &  "9"
			Response.Redirect "http://www.afex.cl/compartido/error.asp?description=" &  "9"
		End If
   
      If rsCodigoCliente.EOF Then
			Set BD = Nothing
			Set rsCodigoCliente = Nothing
			Set afxClienteXc = Nothing
			Set afxClienteXp = Nothing
         'Response.Redirect "http://200.72.160.51/afexmoneyweb/compartido/error.asp?description=" &  "10"
         Response.Redirect "http://www.afex.cl/compartido/error.asp?description=" &  "10"
      Else
         CodigoClienteXp = rsCodigoCliente("codigo_afexpress")
      End If
      If Nombres = "" Then	Nombres = RazonSocial
      sApellidos = Trim(ApellidoPaterno & " " & ApellidoMaterno)
		If sApellidos = "" Then sApellidos = "EMPRESA"

      'Ingresar cliente en el afexpress
      '**********************************************************
      'ConexionXp = "DSN=wafex_giros;UID=giros;PWD=giros;"
		ConexionXp = Session("afxCnxAFEXpress")
      afxClienteXp.Agregar ConexionXp, CodigoClienteXp, Sucursal, Rut, _
                           Pasaporte, PaisParticular, Nombres, sApellidos, _
                           FechaNacimiento, DireccionParticular, ComunaParticular, _
                           CiudadParticular, PaisParticular, PaisFonoParticular, AreaFonoParticular, _
                           NumeroFonoParticular
                           
      If Err.Number <> 0 Then 
			Set BD = Nothing
			Set rsCodigoCliente = Nothing
			Set afxClienteXc = Nothing
			Set afxClienteXp = Nothing
			'Response.Redirect "http://200.72.160.51/afexmoneyweb/compartido/error.asp?description=" &  "12"
			Response.Redirect "http://www.afex.cl/compartido/error.asp?description=" &  "12"
	  End If		   
      
      If afxClienteXp.ErrNumber <> 0 Then 
			Set BD = Nothing
			Set rsCodigoCliente = Nothing
			Set afxClienteXc = Nothing
			Set afxClienteXp = Nothing
			'Response.Redirect "http://200.72.160.51/afexmoneyweb/compartido/error.asp?description=" &  "13"
			Response.Redirect "http://www.afex.cl/compartido/error.asp?description=" &  "13"
	  End If
      '************************************************************
   End If
   
	sSQL = "Select max(codigo) as Codigo From Cliente"
	Set rsCodigoCliente = EjecutarSQLCliente(Conexion, sSQL)

	If Err.number <> 0 Then
		Set BD = Nothing
		Set rsCodigoCliente = Nothing
		Set afxClienteXc = Nothing
		Set afxClienteXp = Nothing

		Response.Redirect "http:../compartido/error.asp?description=" &  Err.number & "<br>" & sSQL
		Exit Function
	End If

   'If MTS Then GetObjectContext.SetComplete
   
   Agregar = rsCodigoCliente("Codigo")
   
   Set BD = Nothing
   Set rsCodigoCliente = Nothing
   Set afxClienteXc = Nothing
   Set afxClienteXp = Nothing
   
End Function

Function EvaluarStr(ByVal Valor)
	Dim Devuelve
	
	If Valor="" Then 
		EvaluarStr = "Null"	
	Else
		EvaluarStr = "'" & Valor & "'"
	End If

End Function

%>
<!--#INCLUDE virtual="/Compartido/Rutinas.asp" -->
<html>
<head>
<META name=VI60_defaultClientScript content=VBScript>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link rel="stylesheet" type="text/css" href="Estilos/Cliente.css">
<title>AFEX Ltda.</title>
</head>
<script language="VBScript">
<!--
	Sub imgVolver_onClick()
		window.navigate "Principal.asp"
	End Sub
-->
</script>
<body>
<table ID="tabPaso1" CELLSPACING="0" CELLPADDING="1" BORDER="0" style="font-family: Verdana; color: black" style="position: relative; left: 10px">
<tr>
	<td style="font-family: Verdana; font-size: 20; color: silver">Solicitud Exitosa</td>
</tr>
<tr>
	<td><br>Su solicitud fue procesada exitosamente</td>
</tr>
<tr align="right">
	<td><br><img id="imgVolver" src="images/BotonVolver.jpg" border="0" alt="Volver" WIDTH="70" HEIGHT="20"></td>
</tr>
</table>
</body>
</html>
