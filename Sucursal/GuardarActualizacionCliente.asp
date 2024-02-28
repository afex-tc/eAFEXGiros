<%@ LANGUAGE = VBScript %>
<% 
	response.Buffer = True
	response.Clear 
%>
<!--#INCLUDE Virtual="/Compartido/Errores.asp" -->
<!--#INCLUDE Virtual="/Compartido/Rutinas.asp" -->
<!--#INCLUDE Virtual="/sucursal/Rutinas.asp" -->
<%	
	
Dim CodigoClienteXc  
Dim CodigoClienteXp  
Dim nCodigo, sNombre, sRut
Dim nTipoCliente
Dim nChkInfCom, nChkInfTri, nChkPon, nChkAcrCom, nChkCarCli, sStr
Dim nCreditoAnterior
Dim shabilita
Dim nClienteAgencia
Dim sEstadoOriginal

nCodigo = Request("cc")
nTipoCliente = Request("tc")
nCreditoAnterior = Request("ca")
nClienteAgencia = request("CAG")
sEstadoOriginal = Session("sEstadoOriginal")

If nClienteAgencia Then
	nClienteAgencia = 1
Else
	nClienteAgencia = 0 	
End If

if request.form("opthabilitar")="on" then shabilita=1 else shabilita=0
If Request.Form("chkInformacionComercial") = "on" Then nChkInfCom = 1 Else nChkInfCom = 0
If Request.Form("chkInformacionTributaria") = "on" Then nChkInfTri = 1 Else nChkInfTri = 0
If Request.Form("chkPonderacion") = "on" Then nChkPon = 1 Else nChkPon = 0
If Request.Form("chkAcreditacionComercial") = "on" Then nChkAcrCom = 1 Else nChkAcrCom = 0
If Request.Form("chkCarpetaCliente") = "on" Then nChkCarCli = 1 Else nChkCarCli = 0

If Not ActualizarCliente Then 
End If

sNombre = TRIM(Request.Form("txtNombres")) & " " & TRIM(Request.Form("txtApellidoP")) & " " & TRIM(Request.Form("txtApellidoM")) 

If cCur(0 & Request.Form("txtCredito")) <> cCur(0 & nCreditoAnterior) Then
	AgregarHistoria nCodigo, "Actualización de Crédito: de " & cCur(0 & nCreditoAnterior) & " a " & cCur(0 & Request.Form("txtCredito")), 1, 0
	AgregarHistoriaCredito Session("afxCnxCorporativa")
	'Se envía mail MS 10-07-2013
	Dim Cuerpo, Asunto
	Asunto = "Modificación de Crédito del Cliente " & trim(sNombres)
			 
	Cuerpo = "Estimado,<br /><br />Se informa que el usuario " & Session("NombreUsuarioOperador") & _
			 " ha realizado la siguiente acción con el cliente&nbsp;" &  trim(sNombre) & ", RUT " &  Request.Form("txtRut") & ":" & _
			 "<br /><br /><b>Actualización de Crédito: de " & cCur(0 & nCreditoAnterior) & " a " & cCur(0 & Request.Form("txtCredito")) & "</b>" & _
			 "<br /><br /> Atte,<br /><br />Servicio de Mensajería Afex."
	
	EnviarEMailBD 8, 29,Session("AmbienteServidorCorreo"),Asunto, Cuerpo
End If

If cInt(0 & nChkInfCom) <> cInt(0 & Request.Form("ChkInfCom")) Then
	If nChkInfCom = 1 Then sStr = "activó" Else sStr = "desactivó"
	AgregarHistoria nCodigo, "CERTIFICACION: Se " & sStr & " Información Comercial", 1, 0
End If
If cInt(0 & nChkInfTri) <> cInt(0 & Request.Form("ChkInfTri")) Then
	If nChkInfTri = 1 Then sStr = "activó" Else sStr = "desactivó"
	AgregarHistoria nCodigo, "CERTIFICACION: Se " & sStr & " Información Tributaria", 1, 0
End If
If cInt(0 & nChkPon) <> cInt(0 & Request.Form("ChkPon")) Then
	If nChkPon = 1 Then sStr = "activó" Else sStr = "desactivó"
	AgregarHistoria nCodigo, "CERTIFICACION: Se " & sStr & " Ponderación", 1 , 0
End If
If cInt(0 & nChkAcrCom) <> cInt(0 & Request.Form("ChkAcrCom")) Then
	If nChkAcrCom = 1 Then sStr = "activó" Else sStr = "desactivó"
	AgregarHistoria nCodigo, "CERTIFICACION: Se " & sStr & " Acreditación Comercial", 1, 0
End If
If cInt(0 & nChkCarCli) <> cInt(0 & Request.Form("ChkCarCli")) Then
	If nChkCarCli = 1 Then sStr = "activó" Else sStr = "desactivó"
	AgregarHistoria nCodigo, "CERTIFICACION: Se " & sStr & " Carpeta Cliente", 1, 0
End If

'AMP 02-02-2016
dim nvlRiesgo 
nvlRiesgo = request.Form("cbxRiesgo")
If nvlRiesgo = 2 Then
    AgregarHistoria nCodigo, "Actualiza Cliente - nivel de riesgo: CAUTELA", 1, 0

    response.Redirect "http:DetalleClienteCautela.asp?cc=" & nCodigo
Else
    AgregarHistoria nCodigo, "Actualiza Cliente - nivel de riesgo: NORMAL", 1, 0

    Response.Redirect "http:DetalleCliente.asp?cc=" & nCodigo & "&nc=" & sNombre & "&rt=" & Request.Form("txtRut")
End If


' actualiza el cliente en la sucursal que lo solicitó, solo en caso que ya exista en dicha sucursal
'if Request.Form("txtSucursalSolicitante") <> "" then
'	HabilitarCliente nCodigo, valorrut(Request.Form("txtRut")), Request.Form("txtPasaporte"), Request.Form("txtSucursalSolicitante")
'end if

'Response.Redirect "http:DetalleCliente.asp?cc=" & nCodigo & "&nc=" & sNombre & _
'										 "&rt=" & Request.Form("txtRut")
'FIN AMP 02-02-2016

Private Function ActualizarCliente
	Dim afxCliente
	Dim nEnvioGiro, nRecepcionGiro, nInformeGiro
	Dim nEnvioTransfer, nInformeTransfer
	Dim nCompraVenta, nAlarmas, nNoticias
	Dim bContacto
	
	ActualizarCliente = False
	
	Set afxCliente = Server.CreateObject("AfexCorporativo.Cliente")

	If Err.number <> 0 Then
		response.Redirect "http:../Compartido/Error.asp?Titulo=Error en ActualizarCliente&Number=" & Err.number   & "&Source=" & Err.Source & "&Description=" & Err.description
		Exit Function
	End If

	nEnvioGiro = 0
	nRecepcionGiro = 0
	nInformeGiro = 0
	nEnvioTransfer = 0
	nInformeTransfer = 0
	nAlarmas = 0
	nNoticias = 0
	nCompraVenta = 0
	
	
	If Not Actualizar(Session("afxCnxCorporativa"), "WB", _
					  Request.Form("txtRut"), Request.Form("txtPasaporte"), nTipoCliente, _
					  Request.Form("txtApellidoP"), Request.Form("txtApellidoM"), _
					  Request.Form("txtNombres"), Request.Form("txtFechaNacimiento") , _
					  Request.Form("txtDireccionPersonal"), Request.Form("txtNumero"), Request.Form("txtDepto") , _							 
					  Request.Form("cbxPaisPersonal"), _
					  Request.Form("cbxCiudadPersonal"), request.Form("cbxComunaPersonal"), _
					  CInt(0 & Request.Form("txtPaisFonoPersonal")), CInt(0 & request.Form("txtAreaFonoPersonal")), CCur(0 & request.Form("txtFonoPersonal")), _
					  CInt(0 & Request.Form("txtPaisFaxPersonal")), CInt(0 & request.Form("txtAreaFaxPersonal")), CCur(0 & request.Form("txtFaxPersonal")), CCur(0 & request.Form("txtCelular")), _
					  Request.Form("txtCorreo"), Request.Form("txtUsuario"), Request.Form("txtPassword"), _
					  Trim(Request.Form("txtRazonSocial") & Request.Form("txtNombreEmpresa")), _
					  Request.Form("txtRepresentante"), Request.Form("txtDireccionLaboral"), Request.Form("cbxCiudadLaboral"), _
					  Request.Form("cbxComunaLaboral"), _
					  CInt(0 & request.Form("txtPaisFonoLaboral")), cInt(0 & request.Form("txtAreaFonoLaboral")), CCur(0 & request.Form("txtFonoLaboral")), _
					  CInt(0 & request.Form("txtPaisFaxLaboral")), cInt(0 & request.Form("txtAreaFaxLaboral")), CCUr(0 & request.Form("txtFaxLaboral")), _
					  "", "",  bContacto, "", "",  _
					  Request.Form("txtApellidoPContacto"), Request.Form("txtApellidoMContacto"), _
					  Request.Form("txtNombresContacto"), "", nEnvioGiro, nRecepcionGiro, nInformeGiro, _
					  nEnvioTransfer, nInformeTransfer, nCompraVenta, nAlarmas, nNoticias, 0, 0, _
					  Request.Form("cbxSucursal"), 1, CInt(Request.Form("cbxBanco")), Request.Form("txtCuentaCorriente"), Request.Form("txtCuentaAhorro"), _
					  nCodigo, nChkInfCom, nChkInfTri, nChkPon, nChkAcrCom, nChkCarCli, cCur(0 & Request.Form("txtCredito")), _
					  Request.Form("cbxEjecutivos"), Request.Form("cbxNacionalidad"), CInt(0 & Trim(request.Form("cbxRubro"))), request.Form("txtPregunta"),request.Form("txtRespuesta"), _
					  request.Form("cbxRiesgo"), Request.Form("cbxPerfilPEP"), _
					  Request.Form("cbxPerfilZona"), Request.Form("cbxPerfilRS"), _
					  Request.Form("cbxPerfilACT"), Request.Form("cbxPerfilIndustria"), _
					  Request.Form("txtContacto1") , Request.Form("txtPorcentageContacto1") , _
					  Request.Form("txtContacto2"), Request.Form("txtPorcentageContacto2") , _
					  Request.Form("txtFechaActivacionComision"), Request.Form("txtProfesion"), Request.Form("cbxSexo"), nClienteAgencia, _
					  Request.form("txtPropositoTransacciones"), _
				      Request.form("txtReferenciaBancaria"), _
				      Request.form("txtOrigenFondos"), _
				      Request.Form("txtMontoPatrimonio"), _
				      Request.form("txtIngresoAnual"), _
				      Request.form("txtCantidadTransaccionesMesCLIENTE"), _
				      Request.form("txtCantidadTransaccionesMesAFEX"), _
				      Request.form("txtMontoTransaccionesMesCLIENTE"), _
				      Request.form("txtMontoTransaccionesMesAFEX"), _
				      Request.Form("txtActividadEconomica"), _
				      Request.Form("txtNombreEmpleador"), _
				      Request.Form("txtGiroComercial"), _
                      Request.Form("cbxPerfilCliente"), _
                      Request.Form("cbxOcupacion")) Then 'INTERNO-9263 MS 18-01-2017
		'FIN INTERNO-6830 MS 21-06-2016			        
        
 		response.Redirect "http:../compartido/Error.asp?Titulo=Error en Guardar Actualizacion Cliente 1&Number=" & Err.Number  & "&Source=" & Err.Source & "&Description=" & Err.Description
		Exit Function
	End If
 
	If Err.number <> 0 Then
		response.Redirect "http:../compartido/Error.asp?Titulo=Error en Guardar Actualizacion Cliente 2&Number=" & Err.Number  & "&Source=" & Err.Source & "&Description=" & Err.Description
		Exit Function
	End If
	If afxCliente.ErrNumber <> 0 Then
		response.Redirect "http:../compartido/Error.asp?Titulo=Error en Guardar Actualizacion Cliente 3&Number=" & afxCliente.ErrNumber  & "&Source=" & afxCliente.ErrSource & "&Description=" & replace(afxCliente.ErrDescription, vbCrLf , "^")
		Exit Function
	End If


	If Err.number <> 0 Then
		response.Redirect "http:../compartido/Error.asp?Titulo=Error en Guardar Actualizacion Cliente 4&Number=" & Err.Number  & "&Source=" & Err.Source & "&Description=" & Err.Description
		Exit Function
	End If
			
	Set afxCliente = Nothing
	
	ActualizarCliente = True
End Function	

Function Actualizar(ByVal Conexion, ByVal Sucursal, ByVal Rut, ByVal Pasaporte, ByVal Tipo, ByVal ApellidoPaterno,  _
					ByVal ApellidoMaterno, ByVal Nombres,  ByVal FechaNacimiento, ByVal DireccionParticular,  _
					Byval numero, ByVal Depto, _
					ByVal PaisParticular, ByVal CiudadParticular,  ByVal ComunaParticular, ByVal PaisFonoParticular,  _
					ByVal AreaFonoParticular,  ByVal NumeroFonoParticular, ByVal PaisFaxParticular,  ByVal AreaFaxParticular,  _
					ByVal NumeroFaxParticular, ByVal Celular,  ByVal CorreoElectronicoCliente, ByVal NombreUsuario,  _
					ByVal Password, ByVal RazonSocial,ByVal Representante, ByVal DireccionComercial,  ByVal CiudadComercial, _
					ByVal ComunaComercial, ByVal PaisFonoComercial,  ByVal AreaFonoComercial,  ByVal NumeroFonoComercial, _
					ByVal PaisFaxComercial,  ByVal AreaFaxComercial,  ByVal NumeroFaxComercial, ByVal CargoCliente,  _
					ByVal ProfesionCliente, ByVal Contacto,  ByVal CargoContacto,  ByVal RutContacto, ByVal ApellidoPaternoContacto,  _
					ByVal ApellidoMaternoContacto,  ByVal NombresContacto, ByVal CorreoElectronicoContacto, ByVal EnvioGiro,  _
					ByVal RecepcionGiro, ByVal InformeGiro,  ByVal EnvioTransferencia,  ByVal InformeTransferencia, ByVal CompraVenta,  _
					ByVal Alarma,  ByVal Noticias, ByVal ContratoProducto,  ByVal Auditoria, ByVal SucursalOrigen, ByVal IngresoWeb, _
					ByVal Banco, ByVal CuentaCorriente, ByVal CuentaAhorro, ByVal CodigoCliente, ByVal CtfInfCom, ByVal CtfInfTri, _
					ByVal CtfPon, ByVal CtfAcrCom, ByVal CtfCarCli, ByVal Credito, ByVal CodigoEjecutivo, ByVal Nacionalidad, ByVal Rubro, _
					ByVal Pregunta, ByVal Respuesta, Byval Riesgo, Byval PerfilPEP, Byval PerfilZona, Byval PerfilRS, _
					Byval PerfilACT, Byval PerfilIndustria, _
					Byval EmpleadoContacto1, Byval PorcentageContacto1, Byval EmpleadoContacto2, Byval PorcentageContacto2, _
					Byval FechaActivacionComision, Byval Profesion, Byval Sexo, ByVal ClienteAgencia, _
				 Byval PropositoTransacciones, _
				 Byval ReferenciaBancaria, _
				 ByVal OrigenPatrimonialFondos, _
				 ByVal MontoPatrimonio, _
				 ByVal MontoIngresoAnual, _
				 ByVal CantidadTransaccionesMesCLIENTE, _
				 ByVal CantidadTransaccionesMesAFEX, _
				 ByVal MontoTransaccionesUSDMesCLIENTE, _
				 ByVal MontoTransaccionesUSDMesAFEX, _
				 ByVal ActividadEconomica, _
				 ByVal NombreEmpleador, _
				 byval GiroComercial, _
                 byval PerfilCliente, _
                 ByVal OcupacionCliente) 'INTERNO-9263 MS 18-01-2017

	Dim sSQL
	Dim BD
	Dim NombreCompleto 
	Dim objContext
	Dim RutCliente    
	Actualizar = False
   
	' controla los errores
	On Error Resume Next
   
	If Rut <> Empty Then
		RutCliente = ValorRut(Rut)
	Else
		RutCliente = Rut
	End If
	
	CodigoClienteXc = Empty
	CodigoClienteXp = Empty
	
	NombreCompleto = Trim(Nombres) & " " & Trim(ApellidoPaterno) & " " & Trim(ApellidoMaterno)

	If RazonSocial <> Empty Then
		NombreCompleto = RazonSocial
	End If
	
	If Banco = 0 Then
		Banco = "Null"
	End If

	sSQL = " exec ActualizarClienteCorporativo " & _
	EvaluarStr(RutCliente) & ", " & EvaluarStr(Pasaporte) & ", " & Tipo & ", " & _
	EvaluarStr(NombreUsuario) & "," & EvaluarStr(Password) & "," & EvaluarStr(NombreCompleto) & ", " & _
	EnvioGiro & ", " & RecepcionGiro & ", " & InformeGiro & ", " & _
	EnvioTransferencia & ", " & InformeTransferencia & ", " & CompraVenta & ", " & _
	Alarma & ", " & Noticias & ", " & ContratoProducto & ", " & _
	Auditoria & ", " & EvaluarStr(CodigoClienteXp) & ", " & EvaluarStr(CodigoClienteXc) & ", " & _
	EvaluarStr(SucursalOrigen) & ", " & IngresoWeb & ", " & Banco & ", " & _
	EvaluarStr(CuentaCorriente) & ", " & EvaluarStr(CuentaAhorro) & ", " & _
	EvaluarStr(CorreoElectronicoCliente) & ", " & EvaluarStr(RazonSocial) & ", " & _
	EvaluarStr(Representante) & ", " & PaisFonoParticular & ", " & AreaFonoParticular & ", " & NumeroFonoParticular & ", " & _	
	PaisFaxParticular & ", " & AreaFaxParticular & ", " & NumeroFaxParticular & ", " & _
	Celular & ",null , " & EvaluarStr(PaisParticular) & ", " & EvaluarStr(CiudadParticular) & ", " & EvaluarStr(ComunaParticular) & ", " & _
	EvaluarStr(Nombres) & ", " & EvaluarStr(ApellidoPaterno) & ", " & EvaluarStr(ApellidoMaterno) & ", " & _
	CodigoCliente & ", " & CtfInfCom & ", " & CtfInfTri & ", " & CtfPon & ", " & CtfAcrCom & ", " & CtfCarCli & ", " & _
	FormatoNumeroSQL(Credito) & ", " & EvaluarStr(CodigoEjecutivo) & ", " & EvaluarStr(Nacionalidad) & ", " & _
	Rubro & ", " & shabilita & ", " & EvaluarStr(Pregunta) & ", " & EvaluarStr(Respuesta) & ", " & Riesgo & ", " & cint("0" & PerfilPEP) & ", " & cint("0" & PerfilZona) & ", " & _
	cint("0" & PerfilRS) & ", " & cint("0" & PerfilACT) & ", " & cint("0" & PerfilIndustria) & ", " & _
	evaluarstr(right(FechaNacimiento,4) & mid(FechaNacimiento,4,2) & left(FechaNacimiento,2)) & _
	", " & cint("0" & EmpleadoContacto1) & ", " & cint("0" & PorcentageContacto1) & ", " & cint("0" & EmpleadoContacto2) & _
	", " & cint("0" & PorcentageContacto2) & ", " & evaluarstr(FechaActivacionComision) & ", " & evaluarstr(session("NombreUsuarioOperador")) & _
	", " & EvaluarStr(direccionparticular) & ", " & EvaluarStr(numero) & ", " & EvaluarStr(Depto) & ", " & EvaluarStr(Profesion)& ", " & sexo & _
	", " & ClienteAgencia & _
								", " & Evaluarstr(PropositoTransacciones) & _
								", " & Evaluarstr(ReferenciaBancaria) & _
								", " & Evaluarstr(OrigenPatrimonialFondos) & _
								", " & ccur("0" & MontoPatrimonio) & _
				                ", " & ccur("0" & MontoIngresoAnual) & _
				                ", " & ccur("0" & CantidadTransaccionesMesCLIENTE) & _
				                ", " & ccur("0" & CantidadTransaccionesMesAFEX) & _
				                ", " & ccur("0" & MontoTransaccionesUSDMesCLIENTE) & _
				                ", " & ccur("0" & MontoTransaccionesUSDMesAFEX) & _
				                ", " & EvaluarStr(ActividadEconomica) & _
				                ", " & EvaluarStr(NombreEMpleador)& _
	                            ", " & EvaluarStr(GiroComercial) & _
                                ", " & PerfilCliente & _
                                ", " & cint("0" & OcupacionCliente)  'INTERNO-9263 MS 18-01-2017

	                            ' JFMG 15-06-2011 ultimos 8 campos por OPTIMA. 9 = ActividadEconomica
	'Conexion
	Set BD = Server.CreateObject("ADODB.Connection")
	BD.Open Conexion                          'Abre la conexion
	If Err.Number <> 0 Then 
		Set BD = Nothing
		Response.Redirect "http:../compartido/error.asp?description=" &  Err.number & "<br>" & Err.Description
	End If
   
   
	BD.BeginTrans
    
	'Consulta
	BD.Execute sSQL                           'Ejecuta la consulta
		
        	
	    	
    If Err.Number <> 0 Then 
		BD.RollbackTrans
		Set BD = Nothing
		Response.Redirect "http:../compartido/error.asp?description=" &  Err.number & "<br>" & Err.Description
		Exit Function	
    End If
   
	BD.CommitTrans
	
	Actualizar = True
   
	Set BD = Nothing
	
    if sEstadoOriginal <> shabilita then
		Dim Cuerpo, Asunto, tipoDocumento, nroDocumento, estadoCliente, estadoAsunto
		If RutCliente <> Empty Then
			tipoDocumento = "Rut"
            nroDocumento = RutCliente
		Else
			 tipoDocumento = "Pasaporte"
             nroDocumento = Pasaporte
		End If

		if shabilita = 0 then 
			estadoCliente="deshabilitado"
			estadoAsunto="Deshabilitación de cliente "					
		else
			estadoCliente="habilitado"
			estadoAsunto="Habilitación de cliente "			
		end if
		
        Asunto = estadoAsunto & NombreCompleto
		Cuerpo = "Estimado [nombreDestinatario]<br /><br />El siguiente cliente:<br /><br /><b>Nombre:</b>&nbsp;" & _
		         NombreCompleto & "<br /><b>Código:</b>&nbsp;" & _
		         CodigoCliente + "<br /><b>" & tipoDocumento & ":</b>&nbsp;" & nroDocumento & _
		         "<br /><br /> ha sido " & estadoCliente & " en Atención Cliente.<br /><br />Atte.<br /><br />Servicio de Mensajería Afex."
		EnviarEMailBD 8,34,Session("AmbienteServidorCorreo"),Asunto, Cuerpo
   
    end if
    
End Function

Function AgregarHistoriaCredito(Conexion)
	Dim sSql
	Dim BD
	
	On Error Resume Next
	
	AgregarHistoriaCredito = False
	
	sSQL = "InsertarHistoriaCredito " & nCodigo & ", '" & FormatoFechaSQL(Date) & "', " & _
										cCur(0 & Request.Form("txtCredito")) & ", " & EvaluarStr(Session("NombreUsuarioOperador"))
	'Conexion
	Set BD = Server.CreateObject("ADODB.Connection")
	BD.Open Conexion                          'Abre la conexion
	If Err.Number <> 0 Then 
		Set BD = Nothing
		Response.Redirect "http:../compartido/error.asp?description=" &  Err.number & "<br>" & Err.Description
		Exit Function
	End If
   
	BD.BeginTrans
    
	'Consulta
	BD.Execute sSQL                           'Ejecuta la consulta
    If Err.Number <> 0 Then 
		BD.RollbackTrans
		Set BD = Nothing
		Response.Redirect "http:../compartido/error.asp?description=" &  Err.number & "<br>" & Err.Description
		Exit Function	
    End If
   
	BD.CommitTrans

	AgregarHistoriaCredito = True
	
	Set BD = Nothing
End Function

Sub HabilitarCliente(Corporativa, Rut, Pasaporte, Sucursal)
	dim sSQL
	dim rs
	dim rsSucursales
	dim Conexion
	dim sUsuario, sPassword, sBD
		
	set rs = nothing
	set rsSucursales = nothing		
	
		
	' saca las sucursal
	sSQL = "select nombre, ip_sucursal, basedatos from sucursal where ip_sucursal = " & evaluarstr(Sucursal) & " order by nombre "
	set rsSucursales = ejecutarsqlcliente(session("afxCnxAFEXchange"), sSQL)
	if err.number <> 0 then
		Response.Write "Error al sacar la lista de sucursales para actualizar el cliente. " & err.Description
		Response.End 
	end if	
	
	do while not rsSucursales.eof
		sUsuario = "cambios"
		sPassword = "cambios"
		sBD = "cambios"
		Conexion = "Provider=SQLOLEDB;User ID=" & sUsuario & ";Password=" & sPassword & ";Initial Catalog=" & sBD & ";Data Source=" & trim(rsSucursales("ip_sucursal"))
		if instr(ucase(rsSucursales("nombre")), ucase("casa matriz")) > 0 then 
			Conexion = Session("afxCnxAFEXchange")
		end if
		if instr(ucase(rsSucursales("nombre")), ucase("moneda")) > 0 then 
			Conexion = Session("afxCnxAFEXchangeMoneda")			
		end if		
			
		' ejecuta la actualizacion en todas las sucursales
		sSQL = " update cliente set estado_cliente = 1, codigo_corporativa = " & Corporativa
		if Rut <> "" then
			sSQL = sSQL & " where rut_cliente = " & evaluarstr(Rut)
		elseif Pasaporte <> "" then
			sSQL = sSQL & " where pasaporte_cliente = " & evaluarstr(Pasaporte)
		end if		
		
		set rs = ejecutarsqlcliente(Conexion, sSQL)
		if err.number <> 0 then
			Response.Write "Error al ejecutar el Script en la Sucursal " & rsSucursales("nombre") & ". " & err.Description & Conexion & "//" & ssql
			Response.End 
		end if			
		
		rsSucursales.movenext
	loop
		
	set rs = nothing
	set rsSucursales = nothing				
End Sub

%>