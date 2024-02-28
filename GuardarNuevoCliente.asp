<%@ LANGUAGE = VBScript %>
<% 
	response.Buffer = True
	response.Clear 

	'Response.expires = 0
	'Response.expiresabsolute = Now() - 1
	'Response.addHeader "pragma", "no-cache"
	'Response.addHeader "cache-control", "private"
	'Response.CacheControl = "no-cache"
%>
<!--#INCLUDE virtual="../Compartido/Errores.asp" -->
<!--#INCLUDE virtual="../Compartido/Rutinas.asp" -->
<!--#INCLUDE virtual="../Sucursal/Rutinas.asp" -->
<%	
	
Dim CodigoClienteXc  
Dim CodigoClienteXp  
Dim sCodigo, sNombre
Dim nTipoCliente

nTipoCliente = CInt(0& Request("TCl"))

AgregarCliente
AgregaHistoria

'Response.Redirect "http:DetalleCliente.asp"
sNombre = TRIM(Request.Form("txtNombres")) & " " & TRIM(Request.Form("txtApellidoM")) & " " & TRIM(Request.Form("txtApellidoP"))

Response.Redirect "http:DetalleCliente.asp?cc=" & sCodigo & "&nc=" & sNombre & _
						 "&rt=" & Request.Form("txtRut")

Private Sub AgregaHistoria
	AgregarHistoria sCodigo, "Agrega Cliente", 1
End Sub

Private Sub AgregarCliente
	Dim afxCliente
	
	On Error	Resume Next
	Set afxCliente = Server.CreateObject("AfexCorporativo.Cliente")
	If Err.number <> 0 Then
		response.Redirect "http:../Compartido/Error.asp?Titulo=Error en HágaseCliente&Number=" & Err.number   & "&Source=" & Err.Source & "&Description=" & Err.description
	End If
	On Error	Resume Next
	Dim TipoCliente
	Dim nEnvioGiro, nRecepcionGiro, nInformeGiro
	Dim nEnvioTransfer, nInformeTransfer
	Dim nCompraVenta, nAlarmas, nNoticias
	Dim bContacto

	'If Request.Form("chkEnvioGiro") = "on" Then
	'	nEnvioGiro = 1
	'Else
		nEnvioGiro = 0
	'End If
	'If Request.Form("chkRecepcionGiro") = "on" Then
	'	nRecepcionGiro = 1
	'Else
		nRecepcionGiro = 0
	'End If
	'If Request.Form("chkInformeGiro") = "on" Then
	'	nInformeGiro = 1
	'Else
		nInformeGiro = 0
	'End If
	'If Request.Form("chkEnvioTransfer") = "on" Then
	'	nEnvioTransfer = 1
	'Else
		nEnvioTransfer = 0
	'End If
	'If Request.Form("chkInformeTransfer") = "on" Then
	'	nInformeTransfer = 1
	'Else
		nInformeTransfer = 0
	'End If
	'If Request.Form("chkAlarmas") = "on" Then
	'	nAlarmas = 1
	'Else
		nAlarmas = 0
	'End If
	'If Request.Form("chkNoticias") = "on" Then
	'	nNoticias = 1
	'Else
		nNoticias = 0
	'End If
	'If Request.Form("chkCompraVenta") = "on" Then
	'	nCompraVenta = 1
	'Else
		nCompraVenta = 0
	'End If
	
	'If Request.Form("optPersona") = "1" Then
	'	TipoCliente = 1
	'	bContacto = False
	'Else
	'	TipoCliente = 2
	'	bContacto = True
	'End If
	'response.Redirect "Compartido/Error.asp?description=" & Session("afxCnxCorporativa") & ", " & "WB" & ", " &  _
	'						 Request.Form("txtRut") & ", " &  Request.Form("txtPasaporte") & ", " &  TipoCliente
 	'Session("afxCnxCorporativa") = "DSN=wAfexCorporativa;UID=corporativa;PWD=afxsqlcor;"
	If Err.number <> 0 Then
		response.Redirect "http:../compartido/Error.asp?Titulo=Error en HágaseCliente 0&Number=" & Err.Number  & "&Source=" & Err.Source & "&Description=" & Err.Description
	End If
	
	sCodigo = Agregar(Session("afxCnxCorporativa"), "WB", _
					  Request.Form("txtRut"), Request.Form("txtPasaporte"), nTipoCliente, _
					  Request.Form("txtApellidoP"), Request.Form("txtApellidoM"), _
					  Request.Form("txtNombres"), "Null", _
					  Request.Form("txtDireccionPersonal"), _							 
					  Request.Form("cbxPaisPersonal"), _
					  Request.Form("cbxCiudadPersonal"), request.Form("cbxComunaPersonal"), _
					  CInt(0 & Request.Form("txtPaisFonoPersonal")), CInt(0 & request.Form("txtAreaFonoPersonal")), CCur(0 & request.Form("txtFonoPersonal")), _
					  CInt(0 & Request.Form("txtPaisFaxPersonal")), CInt(0 & request.Form("txtAreaFaxPersonal")), CCur(0 & request.Form("txtFaxPersonal")), CCur(0 & request.Form("txtCelular")), _
					  Request.Form("txtCorreo"), Request.Form("txtUsuario"), Request.Form("txtPassword"), _
					  Trim(Request.Form("txtRazonSocial") & Request.Form("txtNombreEmpresa")), _
					  Request.Form("txtDireccionLaboral"), Request.Form("cbxCiudadLaboral"), _
					  Request.Form("cbxComunaLaboral"), _
					  CInt(0 & request.Form("txtPaisFonoLaboral")), cInt(0 & request.Form("txtAreaFonoLaboral")), CCur(0 & request.Form("txtFonoLaboral")), _
					  CInt(0 & request.Form("txtPaisFaxLaboral")), cInt(0 & request.Form("txtAreaFaxLaboral")), CCUr(0 & request.Form("txtFaxLaboral")),  _
					  "", "",  bContacto, "", "",  _
					  Request.Form("txtApellidoPContacto"), Request.Form("txtApellidoMContacto"), _
					  Request.Form("txtNombresContacto"), "", nEnvioGiro, nRecepcionGiro, nInformeGiro, _
					  nEnvioTransfer, nInformeTransfer, nCompraVenta, nAlarmas, nNoticias, 0, 0, _
					  Request.Form("cbxSucursal"), 1, CInt(Request.Form("cbxBanco")), Request.Form("txtCuentaCorriente"), Request.Form("txtCuentaAhorro"), _
					  Request.Form("cbxPaisPasaporte"), CInt(0 & Request.Form("cbxEjecutivo")), _
					  Request.Form("cbxNacionalidad"), CInt(0 & Request.Form("cbxRubro")))
 	'Session("afxCnxCorporativa") = "DSN=AfexCorporativa;UID=corporativa;PWD=afxsqlcor;"

	If Err.number <> 0 Then
		response.Redirect "http:../compartido/Error.asp?Titulo=Error en HágaseCliente 1&Number=" & Err.Number  & "&Source=" & Err.Source & "&Description=" & Err.Description
	End If
	If afxCliente.ErrNumber <> 0 Then
		response.Redirect "http:../compartido/Error.asp?Titulo=Error en HágaseCliente 2&Number=" & afxCliente.ErrNumber  & "&Source=" & afxCliente.ErrSource & "&Description=" & replace(afxCliente.ErrDescription, vbCrLf , "^")
	End If
			
	Set afxCliente = Nothing
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
                 ByVal ContratoProducto,  ByVal Auditoria, ByVal SucursalOrigen, _
                 ByVal IngresoWeb, ByVal Banco, ByVal CuentaCorriente, _
                 ByVal CuentaAhorro, ByVal PaisPasaporte, ByVal CodigoEjecutivo, _
                 ByVal Nacionalidad, ByVal Rubro)

	Dim sSQL
	Dim BD
	Dim rsCodigoCliente
	Dim Mensaje
	Dim NombreCompleto 
	Dim objContext
	Dim RutCliente
	Dim Raya
	Dim sApellidos
      
	Agregar = 0
   
	' controla los errores
	On Error Resume Next
   
	'Leer vista de la base de datos corporativa
	'sSQL = "SELECT exchange, express FROM vatencionclienteprueba " & _
	'       "WHERE 1 = 1 "
   
	If Rut <> Empty Then
		Rut = ValorRut(Rut)
		'sSQL = sSQL & " And Rut = '" & Rut & "' "
		'sSQL = sSQL & "AND rut = '" & ValorRut(Rut) & "' "
	End If
	'If Pasaporte <> Empty Then
	'	sSQL = sSQL & "AND pasaporte = '" & Pasaporte & "' "
	'End If
	
   
	'Set rsCodigoCliente = EjecutarSQLCliente(Conexion, sSQL)
	
	'If Err.Number <> 0 Then 
	'	Set BD = Nothing
	'	Set rsCodigoCliente = Nothing
	'	Set afxClienteXc = Nothing
	'	Set afxClienteXp = Nothing
	'	Response.Redirect "http:/compartido/error.asp?description=" &  "1"
	'End If

	'If Rut = "" Then
	'	Rut = "Null"
	'Else
	'   Rut = EvaluarStr(Rut)
	'End If
	'If Err.Number <> 0 Then 
	'	Set BD = Nothing
	'	Set rsCodigoCliente = Nothing
	'	Response.Redirect "http:/compartido/error.asp?description=" &  "2"
	'End If
   
	'If rsCodigoCliente.EOF Then
	   CodigoClienteXc = Empty
	   CodigoClienteXp = Empty
	'End If
	'If Err.Number <> 0 Then 
	'	Set BD = Nothing
	'	Set rsCodigoCliente = Nothing
	'  Response.Redirect "http:/compartido/error.asp?description=" &  "3"
	'End If
	
	NombreCompleto = Trim(Nombres) & " " & Trim(ApellidoPaterno) & " " & Trim(ApellidoMaterno)
	If RazonSocial <> Empty Then
		NombreCompleto = RazonSocial
	End If
	
	If Banco = 0 Then
		Banco = "Null"
	End If

	sSQL = "InsertarClienteCorporativo " & _
								EvaluarStr(Rut) & ", " & EvaluarStr(Pasaporte) & ", " & Tipo & ", " & _
								EvaluarStr(NombreUsuario) & "," & EvaluarStr(Password) & "," & EvaluarStr(NombreCompleto) & ", " & _
								EnvioGiro & ", " & RecepcionGiro & ", " & InformeGiro & ", " & _
								EnvioTransferencia & ", " & InformeTransferencia & ", " & CompraVenta & ", " & _
								Alarma & ", " & Noticias & ", " & ContratoProducto & ", " & _
								Auditoria & ", " & EvaluarStr(CodigoClienteXp) & ", " & EvaluarStr(CodigoClienteXc) & ", " & _
								EvaluarStr(SucursalOrigen) & ", " & IngresoWeb & ", " & Banco & ", " & _
								EvaluarStr(CuentaCorriente) & ", " & EvaluarStr(CuentaAhorro) & ", " & _
								EvaluarStr(CorreoElectronicoCliente) & ", " & EvaluarStr(RazonSocial) & ", " & _
								PaisFonoParticular & ", " & AreaFonoParticular & ", " & NumeroFonoParticular  & ", " & _
								PaisFaxParticular & ", " & AreaFaxParticular & ", " & NumeroFaxParticular & ", " & _
								Celular & ", " & EvaluarStr(DireccionParticular) & ", " & EvaluarStr(PaisParticular) & ", " & EvaluarStr(CiudadParticular) & ", " & EvaluarStr(ComunaParticular) & ", " & _
								EvaluarStr(Nombres) & ", " & EvaluarStr(ApellidoPaterno) & ", " & EvaluarStr(ApellidoMaterno) & ", " & _
								EvaluarStr(PaisPasaporte) & ", " & CodigoEjecutivo & ", " & EvaluarStr(Nacionalidad) & ", " & _
								Rubro
								
	
	'Conexion
	Set BD = Server.CreateObject("ADODB.Connection")
	BD.Open Conexion                          'Abre la conexion
	If Err.Number <> 0 Then 
		Set BD = Nothing
		'Set rsCodigoCliente = Nothing
		'Set afxClienteXc = Nothing
		'Set afxClienteXp = Nothing
		Response.Redirect "http:../compartido/error.asp?description=" &  "4"
	End If
   
	BD.BeginTrans
    
	'Consulta
	BD.Execute sSQL                           'Ejecuta la consulta
    If Err.Number <> 0 Then 
		BD.RollbackTrans
		Set BD = Nothing
		'Set rsCodigoCliente = Nothing
		'Set afxClienteXc = Nothing
		'Set afxClienteXp = Nothing
		Response.Redirect "http:../compartido/error.asp?description=" &  Err.number & "<br>" & Err.Description
		Exit Function	
    End If
   
	BD.CommitTrans
	
	sSQL = "Select max(codigo) as Codigo From Cliente"
	Set rsCodigoCliente = EjecutarSQLCliente(Conexion, sSQL)
	
	If Err.number <> 0 Then
		Set rsCodigoCliente = Nothing
		Set BD = Nothing
		Response.Redirect "http:../compartido/error.asp?description=" &  Err.number & "<br>" & sSQL
		Exit Function
	End If
	
	Agregar = rsCodigoCliente("Codigo")
   
	Set rsCodigoCliente = Nothing
	Set BD = Nothing
	'Set rsCodigoCliente = Nothing
   
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
<!--#INCLUDE virtual="../Compartido/Rutinas.asp" -->
