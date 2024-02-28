<%@ LANGUAGE = VBScript %>
<% 
	'response.Buffer = True
	'response.Clear 
	
%>
<!--#INCLUDE virtual="/Compartido/Errores.asp" -->
<!--#INCLUDE virtual="/Compartido/Rutinas.asp" -->
<!--#INCLUDE virtual="/sucursal/Rutinas.asp" -->
<%	
	
Dim CodigoClienteXc  
Dim CodigoClienteXp  
Dim sCodigo, sNombre
Dim nTipoCliente
dim nClienteAgencia

nTipoCliente = CInt(0& Request("TCl"))


nCLienteAgencia = Request("CA")
If nClienteagencia Then
	nClienteAgencia = 1
Else
	nClienteAgencia= 0
End If

AgregarCliente
AgregaHistoria

' actualiza el cliente en la sucursal que lo solicitó, solo en caso que ya exista en dicha sucursal
if Request.Form("txtSucursalSolicitante") <> "" then
	HabilitarCliente sCodigo, valorrut(Request.Form("txtRut")), Request.Form("txtPasaporte"), Request.Form("txtSucursalSolicitante")
end if


sNombre = TRIM(Request.Form("txtNombres")) & " " & TRIM(Request.Form("txtApellidoM")) & " " & TRIM(Request.Form("txtApellidoP"))

Response.Redirect "http:DetalleCliente.asp?cc=" & sCodigo & "&nc=" & sNombre & _
						 "&rt=" & Request.Form("txtRut")

Private Sub AgregaHistoria
	AgregarHistoria sCodigo, "Agrega Cliente", 1, 0
End Sub

Private Sub AgregarCliente
	'Dim afxCliente
	
	On Error Resume Next
	'Set afxCliente = Server.CreateObject("AfexCorporativo.Cliente")
	'If Err.number <> 0 Then
	'	response.Redirect "http:../Compartido/Error.asp?Titulo=Error en HágaseCliente&Number=" & Err.number   & "&Source=" & Err.Source & "&Description=" & Err.description
	'End If
	
	Dim TipoCliente
	Dim nEnvioGiro, nRecepcionGiro, nInformeGiro
	Dim nEnvioTransfer, nInformeTransfer
	Dim nCompraVenta, nAlarmas, nNoticias
	Dim bContacto

	
		nEnvioGiro = 0
	
		nRecepcionGiro = 0
	
		nInformeGiro = 0
	
		nEnvioTransfer = 0
	
		nInformeTransfer = 0
	
		nAlarmas = 0
	
		nNoticias = 0
	
		nCompraVenta = 0
	
	If Err.number <> 0 Then
		response.Redirect "http:../compartido/Error.asp?Titulo=Error en HágaseCliente 0&Number=" & Err.Number  & "&Source=" & Err.Source & "&Description=" & Err.Description
	End If
	
	sCodigo = Agregar(Session("afxCnxCorporativa"), "WB", _
					  Request.Form("txtRut"), Request.Form("txtPasaporte"), nTipoCliente, _
					  Request.Form("txtApellidoP"), Request.Form("txtApellidoM"), _
					  Request.Form("txtNombres"), Request.Form("txtFechaNacimiento"), _
					  Request.Form("txtDireccionPersonal"), _
					  Request.Form("txtnumero"), _
					  Request.Form("txtdepto"), _
			          Request.Form("cbxPaisPersonal"), _
					  Request.Form("cbxCiudadPersonal"), request.Form("cbxComunaPersonal"), _
			          CInt(0 & Request.Form("txtPaisFonoPersonal")), CInt(0 & request.Form("txtAreaFonoPersonal")), CCur(0 & request.Form("txtFonoPersonal")), _
			          CInt(0 & Request.Form("txtPaisFaxPersonal")), CInt(0 & request.Form("txtAreaFaxPersonal")), CCur(0 & request.Form("txtFaxPersonal")), CCur(0 & request.Form("txtCelular")), _
			          Request.Form("txtCorreo"), Request.Form("txtUsuario"), Request.Form("txtPassword"), _
			          Trim(Request.Form("txtRazonSocial") & Request.Form("txtNombreEmpresa")), _
			          Request.Form("txtRepresentante"), Request.Form("txtDireccionLaboral"), Request.Form("cbxCiudadLaboral"), _
					  Request.Form("cbxComunaLaboral"), _
			          CInt(0 & request.Form("txtPaisFonoLaboral")), cInt(0 & request.Form("txtAreaFonoLaboral")), CCur(0 & request.Form("txtFonoLaboral")), _
			          CInt(0 & request.Form("txtPaisFaxLaboral")), cInt(0 & request.Form("txtAreaFaxLaboral")), CCUr(0 & request.Form("txtFaxLaboral")),  _
			          "", "",  bContacto, "", "",  _
			          Request.Form("txtApellidoPContacto"), Request.Form("txtApellidoMContacto"), _
			          Request.Form("txtNombresContacto"), "", nEnvioGiro, nRecepcionGiro, nInformeGiro, _
			          nEnvioTransfer, nInformeTransfer, nCompraVenta, nAlarmas, nNoticias, 0, 0, _
			          Request.Form("cbxSucursal"), 1, CInt(Request.Form("cbxBanco")), _
					  Request.Form("txtCuentaCorriente"), Request.Form("txtCuentaAhorro"), _
			          Request.Form("cbxPaisPasaporte"), CInt(0 & Request.Form("cbxEjecutivo")), _
					  Request.Form("cbxNacionalidad"), CInt(0 & Request.Form("cbxRubro")), _
					  request.Form("txtPregunta"), request.Form("txtRespuesta"), _
					  Request.Form("cbxRiesgo"), Request.Form("cbxPerfilPEP"), _
					  Request.Form("cbxPerfilZona"), Request.Form("cbxPerfilRS"), _
					  Request.Form("cbxPerfilACT"), Request.Form("cbxPerfilIndustria"), _
					  Request.Form("txtContacto1"), Request.Form("txtPorcentageContacto1") , _
					  Request.Form("txtContacto2"), Request.Form("txtPorcentageContacto2"), _
					  Request.Form("txtFechaActivacionComision"), Request.Form("txtNumeroSerie"), Request.form("txtIDConsultada") , _
					  Request.form("txtIDValida"), Request.Form("txtMensajeRegistro"), Request.Form("txtProfesion"), _
					  Request.Form("cbxSexo"),  nClienteAgencia, _
					  Request.form("txtPropositoTransacciones"), _
				      Request.form("txtReferenciaBancaria"), _
				      Request.form("txtOrigenFondos"), _
				      Request.form("txtMontoPatrimonio"), _
				      Request.form("txtIngresoAnual"), _
				      Request.form("txtCantidadTransaccionesMesCLIENTE"), _
				      Request.form("txtCantidadTransaccionesMesAFEX"), _
				      Request.form("txtMontoTransaccionesMesCLIENTE"), _
				      Request.form("txtMontoTransaccionesMesAFEX"), _
				      Request.form("txtActividadEconomica"), _
				      Request.form("txtNombreEmpleador"), _
				      Request.Form("txtGiroComercial"), _
                      Request.Form("cbxOcupacion"))  'INTERNO-9263 MS 18-01-2017
	            ' JFMG 15-06-2011 se agregan los últimos 8 campos por las modificaciones solicitadas por Optima) 9 = ActividadEconomica
 	
	'If afxCliente.ErrNumber <> 0 Then
	'	response.Redirect "http:../compartido/Error.asp?Titulo=Error en HágaseCliente 2&Number=" & afxCliente.ErrNumber  & "&Source=" & afxCliente.ErrSource & "&Description=" & replace(afxCliente.ErrDescription, vbCrLf , "^")
	'End If
    
    'Set afxCliente = Nothing
End Sub	

Function Agregar(ByVal Conexion, ByVal Sucursal,  ByVal Rut,  ByVal Pasaporte, _
                 ByVal Tipo, ByVal ApellidoPaterno,  ByVal ApellidoMaterno, _
                 ByVal Nombres,  ByVal FechaNacimiento, _
                 ByVal DireccionParticular, ByVal Numero , ByVal Depto, _
                 ByVal PaisParticular, _
                 ByVal CiudadParticular,  ByVal ComunaParticular, _
                 ByVal PaisFonoParticular,  ByVal AreaFonoParticular,  ByVal NumeroFonoParticular, _
                 ByVal PaisFaxParticular,  ByVal AreaFaxParticular,  ByVal NumeroFaxParticular, _
                 ByVal Celular,  ByVal CorreoElectronicoCliente, _
                 ByVal NombreUsuario,  ByVal Password, _
                 ByVal RazonSocial, ByVal Representante, _
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
                 ByVal Nacionalidad, ByVal Rubro, ByVal Pregunta, ByVal Respuesta, Byval Riesgo, _
				 Byval PerfilPEP, Byval PerfilZona, Byval PerfilRS, Byval PerfilACT, Byval PerfilIndustria, _
				 Byval EmpleadoContacto, Byval PorcentageContacto, Byval EmpleadoContacto1, _
				 Byval PorcentageContacto1, Byval FechaActivacionComision, byval NumeroSerie, byval IDConsultada, _
				 Byval IDValida, Byval MensajeRegistro, _
				 Byval Profesion, Byval Sexo, ByVal ClienteAgencia, _
				 Byval PropositoTransacciones, _
				 Byval ReferenciaBancaria, _
				 ByVal OrigenPatrimonialFondos, _
				 ByVal MontoPatrimonio, _
				 ByVal MontoIngresoAnual, _
				 ByVal CantidadTransaccionesMesCLIENTE, _
				 ByVal CantidadTransaccionesMesAFEX, _
				 ByVal MontoTransaccionesUSDMesCLIENTE, _
				 ByVal MontoTransaccionesUSDMesAFEX, _
				 Byval ActividadEconomica, _
				 ByVal NombreEmpleador, _
				 byval GiroComercial, _
                 byval Ocupacion) 'INTERNO-9263 MS 18-01-2017 'CCP-208 MS 21-10-2016
	' JFMG 15-06-2011 se agregan los últimos 8 campos por las modificaciones solicitadas por Optima. 9 = ActividadEconomica
				

	Dim sSQL
	Dim BD
	Dim rsCodigoCliente
	Dim Mensaje
	Dim NombreCompleto 
	Dim objContext
	Dim RutCliente
	Dim Raya
	Dim sApellidos
	Dim Cnn,rs3,SQL3
	dim sHistoriaID
      
	Agregar = 0
   
	' controla los errores
	On Error Resume Next
   
	'Leer vista de la base de datos corporativa
	
   
	If Rut <> Empty Then
		Rut = ValorRut(Rut)
		
	End If
	
	   CodigoClienteXc = Empty
	   CodigoClienteXp = Empty
	
	
	NombreCompleto = Trim(Nombres) & " " & Trim(ApellidoPaterno) & " " & Trim(ApellidoMaterno)
	If RazonSocial <> Empty Then
		NombreCompleto = RazonSocial
	else
		Sexo = null	
	End If
	
	sexo= Request.Form("cbxSexo")
	
	If Banco = 0 Then
		Banco = "Null"
	End If
	
	if ConsultaID then
		sHistoriaID = 1
	else
		sHistoriaID = 0
	end if
	
	if IDConsultada = "true" then 
		IDConsultada = 1
	else
		IDConsultada = 0
	end if
	if IDValida = "true" then 
		IDValida = 1
	else
		IDValida = 0
	end if
    Dim  ipUsuario, nombrePC
	ipusuario = request.servervariables("REMOTE_ADDR")
    nombrePC = request.servervariables("REMOTE_HOST")

	sSQL = " exec InsertarClienteCorporativo " & _
								EvaluarStr(Rut) & ", " & EvaluarStr(Pasaporte) & ", " & Tipo & ", " & _
								EvaluarStr(NombreUsuario) & "," & EvaluarStr(Password) & "," & EvaluarStr(NombreCompleto) & ", " & _
								EnvioGiro & ", " & RecepcionGiro & ", " & InformeGiro & ", " & _
					    		EnvioTransferencia & ", " & InformeTransferencia & ", " & CompraVenta & ", " & _
								Alarma & ", " & Noticias & ", " & ContratoProducto & ", " & _
				    			Auditoria & ", " & EvaluarStr(CodigoClienteXp) & ", " & EvaluarStr(CodigoClienteXc) & ", " & _
								EvaluarStr(SucursalOrigen) & ", " & IngresoWeb & ", " & Banco & ", " & _
								EvaluarStr(CuentaCorriente) & ", " & EvaluarStr(CuentaAhorro) & ", " & _
								EvaluarStr(CorreoElectronicoCliente) & ", " & EvaluarStr(RazonSocial) & ", " & EvaluarStr(Representante) & ", " & _
								PaisFonoParticular & ", " & AreaFonoParticular & ", " & NumeroFonoParticular  & ", " & _
								PaisFaxParticular & ", " & AreaFaxParticular & ", " & NumeroFaxParticular & ", " & _
	            				Celular & ",null "  &  _
	            				", " & EvaluarStr(PaisParticular) & ", " & EvaluarStr(CiudadParticular) & ", " & EvaluarStr(ComunaParticular) & ", " & _
								EvaluarStr(Nombres) & ", " & EvaluarStr(ApellidoPaterno) & ", " & EvaluarStr(ApellidoMaterno) & ", " & _
								EvaluarStr(PaisPasaporte) & ", " & CodigoEjecutivo & ", " & EvaluarStr(Nacionalidad) & ", " & _
								Rubro & ", " & EvaluarStr(Pregunta) & ", " & EvaluarStr(Respuesta) & _
								", " & Riesgo & ", " & cint("0" & PerfilPEP) & ", " & _
								cint("0" & PerfilZona) & ", " & cint("0" & PerfilRS) & _
								", " & cint("0" & PerfilACT) & ", " & cint("0" & PerfilIndustria) & ", " & evaluarstr(FechaNacimiento) & _
								", " & cint("0" & EmpleadoContacto) & ", " & cint("0" & PorcentageContacto) & _
								", " & cint("0" & EmpleadoContacto1) & ", " & cint("0" & PorcentageContacto1) & ", " & _
								evaluarstr(FechaActivacionComision) & ", " & Evaluarstr(session("NombreUsuarioOperador")) & ", " & _
								evaluarstr(NumeroSerie) & ", " & IDConsultada & ", " & IDValida & ", " & evaluarstr(MensajeRegistro) & ", " & _
								Evaluarstr(Profesion) & ", " & Sexo & "," & evaluarstr(direccionparticular)&"," & EvaluarStr(numero) & ", " & EvaluarStr(depto) & _
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
				                ", " & Evaluarstr(ActividadEconomica) & _
				                ", " & Evaluarstr(NombreEmpleador) & _
				                ", " & evaluarstr(GiroComercial) & _
                                ", " & evaluarstr(ipusuario) & _
                                ", " & evaluarstr(nombrePC) & _
                                ", " & cint("0" & Ocupacion) 'INTERNO-9263 MS 18-01-2017 'CCP-208 MS 21-10-2016
				                ' JFMG 15-06-2011 ultimos 8 campos por OPTIMA. 9 = ActividaEconomica
		
	Dim rs		
	Set rs = EjecutarSQLCliente(Conexion, sSQL)
	
    If Err.Number <> 0 Then 
		
		Response.Redirect "http:../compartido/error.asp?description=" &  Err.number & "<br>" & Err.Description
		Exit Function	
    End If
    
   Agregar = rs("CodigoCliente") ' JFMG 15-06-2011 rsCodigoCliente("Codigo")
   If Err.Number <> 0 Then 
		
		Response.Redirect "http:../compartido/error.asp?description=" &  Err.number & "<br>" & Err.Description
		Exit Function	
    End If
    	  
    Set rs = Nothing
End Function

Function EvaluarStr(ByVal Valor)
	Dim Devuelve
	
	If Valor="" Then 
		EvaluarStr = "Null"	
	Else
		EvaluarStr = "'" & Valor & "'"
	End If

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
			Response.Write "Error al ejecutar el Script en la Sucursal " & rsSucursales("nombre") & ". " & err.Description & Conexion
			Response.End 
		end if			
		
		rsSucursales.movenext
	loop
		
	set rs = nothing
	set rsSucursales = nothing				
End Sub
	

%>
