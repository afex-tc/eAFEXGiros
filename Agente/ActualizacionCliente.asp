<%@ Language=VBScript %>
<!-- #INCLUDE virtual="/Compartido/TimeOut.asp" -->
<!-- #INCLUDE virtual="/Compartido/Errores.asp" -->
<!-- #INCLUDE virtual="/Compartido/Rutinas.asp" -->
<!-- #INCLUDE virtual="/agente/Constantes.asp" -->
<%
    If Session("NombreUsuario") = "" Then
		response.Redirect "http:../Compartido/TimeOut.htm"
		response.end
	End If

	Dim sApellidoP, sApellidoM, sNombres, sRazonSocial, sRut, sPasaporte
	Dim sDireccion, sPais, sCiudad, sComuna, sPaisPass
	Dim sNombrePais, sNombreCiudad, sNombreComuna, sNombrePaisPass
	Dim sPaisFono, sAreaFono, sFono
	Dim sPaisFono2, sAreaFono2, sFono2
	Dim nAccion, sAFEXchange, sAFEXpress
	Dim nTipoCliente, sDisabled
	dim sTarjeta, sSexo, sNacionalidad, sCorreoElectronico, sFechaNacimiento, sTarjetas
	Dim sNumeroCelular ' APPL-9009
	
	dim sSQL, rs
	
	Dim sMensajeUsuario ' JFMG 16-05-2013 para indicar mensajes al usuario(cajero)
	
	nAccion = cInt(0 & Request("Accion"))
	
	sAFEXpress = Request.form("txtExpress")
	sAFEXchange = Request.form("txtExchange")
	
	If sAFEXpress = "" Then
		sAFEXpress = Request("AFEXpress")
	End If
	If sAFEXchange = "" Then
		sAFEXchange = Request("AFEXchange")
	End If
	
	' JFMG 06-08-2008 valida si la identificación que se ingresó ya existe
	if nAccion = 99 then		
		if Request.Form("txtRut") <> "" then
			Set rs = BuscarCliente(1, Request.Form("txtRut"), "", "")
		elseif Request.Form("txtPasaporte") <> "" then
			Set rs = BuscarCliente(2, Request.Form("txtPasaporte"), "", "")
		end if				
		if not rs.eof then
			dim ClienteMantener
			' si el cliente existe lo envia a la página de asociación
			ClienteMantener = rs("Express")
			rs.close
			set rs = nothing
			if ClienteMantener <> "" then
				Response.Redirect "AsociarCliente.asp?Nuevo=" & ClienteMantener & "&Eliminar=" & sAFEXpress
			end if
		end if
		nAccion = afxAccionPais
	end if
	' *************************** FIN *****************************
	
	sTarjetas = Request.Form("cbxTarjetas")
	if trim(sTarjetas) = "" then sTarjetas = "0008010"	
		
	If nAccion = afxAccionPais Then
		CargarActualizacion
	Else
		CargarCliente 
	End If

	Sub CargarActualizacion()
		sRut = request.Form("txtRut")
		sPasaporte = request.Form("txtPasaporte")
		sNombres = request.Form("txtNombres")
		sApellidoP = request.Form("txtApellidoP")
		sApellidoM = request.Form("txtApellidoM")
		sRazonSocial = request.Form("txtRazonSocial")
		sDireccion = request.Form("txtDireccion")
		sPaisPass = Trim(Ucase(Request.Form("cbxPaisPasaporte")))
		sPais = Trim(UCase(Request.Form("cbxPais")))
		sCiudad = Trim(UCase(Request.form("cbxCiudad")))
		sComuna = Trim(Ucase(Request.form("cbxComuna")))
		sPaisFono = request.Form("txtPaisFono")
		sAreaFono = request.Form("txtAreaFono")
		sFono = request.Form("txtFono")
		sPaisFono2 = request.Form("txtPaisFono2")
		sAreaFono2 = request.Form("txtAreaFono2")
		sFono2 = request.Form("txtFono2")
		sTarjeta = request.Form("txtTarjeta")
		sSexo = request.Form("cbxSexo")
		sNacionalidad = request.Form("cbxNacionalidad")
		sCorreoElectronico = request.Form("txtCorreoElectronico")
		sFechaNacimiento = request.Form("txtFechaNacimiento")		       
		sNumeroCelular = request.Form("txtNumeroCelular") ' APPL-9009
		If Request.Form("optPersona") = "on" Then
			nTipoCliente = 1
		Else
			nTipoCliente = 2
		End If
	End Sub


	Sub CargarCliente()
		Dim nCampo, sArgumento, rs
		
		' Jonathan Miranda G. 06-06-2007		
		'------------------- Fin ------------------
		
		If sAFEXpress <> "" Then
			nCampo = 6
			sArgumento = sAFEXpress		
		elseIf sAFEXchange <> "" Then
			nCampo = 5
			sArgumento = sAFEXchange 		
		End If	
		
		Set rs = BuscarCliente(nCampo, sArgumento, "", "")
		If Not rs.EOF Then
			nTipoCliente = cInt(0 & rs("tipo"))
			If nTipoCliente = 1 Then
				sApellidoP = MayMin(EvaluarVar(rs("paterno"), ""))
				sApellidoM = MayMin(EvaluarVar(rs("materno"), ""))
				sNombres = MayMin(EvaluarVar(rs("nombre"), ""))
			Else
				sRazonSocial = MayMin(EvaluarVar(rs("nombre_completo"), ""))
			End If	
			sRut = FormatoRut(EvaluarVar(rs("rut"), ""))
			sPasaporte = EvaluarVar(rs("pasaporte"), "")
			sPaisPass = Trim(Ucase(EvaluarVar(rs("codigo_paispas"), "")))
			sDireccion = MayMin(EvaluarVar(rs("direccion"), ""))
			sPais = Trim(Ucase(EvaluarVar(rs("codigo_pais"), "")))
			sCiudad = Trim(Ucase(EvaluarVar(rs("codigo_ciudad"), "")))
			sComuna = Trim(Ucase(EvaluarVar(rs("codigo_comuna"), "")))
			sPaisFono = EvaluarVar(rs("ddi_pais"), "")
			sAreaFono = EvaluarVar(rs("ddi_area"), "")
			sFono = trim(EvaluarVar(rs("telefono"), ""))
			sPaisFono2 = EvaluarVar(rs("ddi_pais2"), "")
			sAreaFono2 = EvaluarVar(rs("ddi_area2"), "")
			sFono2 = trim(EvaluarVar(rs("telefono2"), ""))
			sNombrePais = MayMin(EvaluarVar(rs("pais"), ""))
			sNombreCiudad = MayMin(EvaluarVar(rs("ciudad"), ""))
			sNombreComuna = MayMin(EvaluarVar(rs("comuna"), ""))
			sNombrePaisPass = MayMin(EvaluarVar(rs("paispas"), ""))
			sTarjetas = trim(EvaluarVar(rs("tarjeta"),""))
			if trim(sTarjetas) <> "" then
				' JFMG 15-05-2013
			    if Len(sTarjetas) < 6 then
			        sMensajeUsuario = "El Cliente presenta problemas con su tarjeta Giro Club. Favor actualizar los datos."
			    else
			    ' FIN JFMG 15-05-2013
				    sTarjetas = left(sTarjetas, len(sTarjetas) - 6)
				' JFMG 15-05-2013
				end if
				' FIN JFMG 15-05-2013
			end if			
			sTarjeta = right(EvaluarVar(rs("tarjeta"),""),6)
			ssexo = EvaluarVar(rs("sexo"),"")
			snacionalidad = EvaluarVar(rs("codigonacionalidad"),"")
			scorreoelectronico = EvaluarVar(rs("correoelectronico"),"")
			sFechaNacimiento = EvaluarVar(rs("fecha_nacimiento"),"")		       
		    sNumeroCelular = EvaluarVar(rs("NumeroCelularCliente"),"") ' APPL-9009
			
			sAFEXchange = rs("Exchange")
			sAFEXpress = rs("Express")
			If Err.number <> 0 Then
				MostrarErrorMS ""
			End If
		End If
		Set rs = Nothing
	End Sub
	
	if sNacionalidad = "" then sNacionalidad = "CL"
%>
<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link rel="stylesheet" type="text/css" href="../Estilos/Cliente.css">
</head>
<script LANGUAGE="VBScript">
<!--
	'Variables de módulo
	'Variables para encabezado
	Dim sEncabezadoFondo
	Dim sEncabezadoTitulo
	
	Const afxCodigoCliente = 3
	Const afxRut = 1
	Const afxPasaporte = 2
	COnst afxNombres = 4
	
	sEncabezadoFondo = "Principal"
	sEncabezadoTitulo = "Actualización de Cliente"

	Sub imgBuscar_onMouseOver()
		 imgBuscar.style.cursor = "Hand"
	End Sub	

	Sub window_onLoad()
		
		CargarMenuActualizar 
		CargarCliente
		HabilitarCampos
		
		frmCliente.cbxTarjetas.value = "<%=sTarjetas%>"
		
		' JFMG 16-05-2013
		<%If sMensajeUsuario <> "" then %>
		    msgbox "<%=sMensajeUsuario%>"
		<%End If %>
		' FIN JFMG 16-05-2013
		
	End Sub

	Sub txtRut_onBlur()
		Dim sRut
		
		If frmCliente.txtRut.value = "" Then Exit Sub
        sRut = ValidarRut(frmCliente.txtRut.value)
		If sRut = Empty Then
			msgbox "El número de Rut no es válido"
            frmCliente.txtRut.value = "" '02-07-2015 APPL-13902 MM
			frmCliente.txtRut.select
			frmCliente.txtRut.focus()
		Else
			frmCliente.txtRut.value = sRut
		End If
		
		' JFMG 06-08-2008 se rellama para ver si el rut existe
		if frmCliente.txtRut.value = "" then exit sub
		
		HabilitarTipo

		frmCliente.action = "ActualizacionCliente.asp?Accion=99"
		frmCliente.submit 
		frmCliente.action = ""
		' **************************** Fin *********************
	End Sub
	
	Sub txtPasaporte_onBlur()
		if frmCliente.txtPasaporte.value = "" then exit sub
				
		' JFMG 06-08-2008 se rellama para ver si el pasaporte existe
		HabilitarTipo
		frmCliente.action = "ActualizacionCliente.asp?Accion=99"
		frmCliente.submit 
		frmCliente.action = ""
		' **************************** Fin *********************
	End Sub

	Sub optRut_onClick()
		window.frmcliente.optpasaporte.checked=False
		window.frmcliente.txtpasaporte.style.display="none"
		window.frmcliente.cbxPaisPasaporte.style.display="none"
		lblPaisPasaporte.style.display="none"
		window.frmCliente.txtRut.style.display = ""		
		frmCliente.txtPasaporte.value = ""
	End Sub
	
	Sub optPasaporte_onClick()
		window.frmcliente.optRut.checked=False
		window.frmCliente.txtRut.style.display = "none"		
		window.frmcliente.txtpasaporte.style.display=""
		window.frmcliente.cbxPaisPasaporte.style.display=""
		lblPaisPasaporte.style.display=""
		frmCliente.txtRut.value = ""
	End Sub


	Sub CargarCliente()
		frmCliente.txtExchange.value = "<%=sAFEXchange%>"
		frmCliente.txtExpress.value = "<%=sAFEXpress%>"
		frmCliente.txtApellidoM.value = "<%=sApellidoM%>"
		frmCliente.txtApellidoP.value = "<%=sApellidoP%>"
		frmCliente.txtNombres.value = "<%=sNombres%>"
		frmCliente.txtRazonSocial.value = "<%=sRazonSocial%>"
		frmCliente.txtDireccion.value = "<%=sDireccion%>"
		frmCliente.txtPaisFono.value = "<%=ObtenerDDI(1, sPais)%>"
		frmCliente.txtPaisFono2.value = frmCliente.txtPaisFono.value '"<%=ObtenerDDI(1, sPais)%>"  
		If frmCliente.cbxComuna.value <> "" Then 
			frmCliente.txtAreaFono.value = "<%=ObtenerDDI(3, sComuna)%>"
		Else
			frmCliente.txtAreaFono.value = "<%=ObtenerDDI(2, sCiudad)%>"		
		End If
		frmCliente.txtAreaFono2.value = frmCliente.txtAreaFono.value
		frmCliente.txtFono.value = "<%=sFono%>"
		frmCliente.txtFono2.value = "<%=sFono2%>"
		frmCliente.txtRut.value = "<%=sRut%>"
		frmCliente.txtPasaporte.value = "<%=sPasaporte%>"
		If frmCliente.txtPasaporte.value <> "" Then
			window.frmcliente.optpasaporte.checked=True
			window.frmcliente.optRut.checked=False
			window.frmCliente.txtRut.style.display = "none"		
			window.frmcliente.txtpasaporte.style.display=""
			window.frmcliente.cbxPaisPasaporte.style.display=""
			lblPaisPasaporte.style.display=""
		End If	
		<% If nTipoCliente <> 1 Then %>
			window.frmCliente.optEmpresa.checked = True
			trEmpresa.style.display = ""
			trPersona.style.display = "none"
			trsexo.style.display ="none"
			window.frmCliente.cbxSexo.style.display ="none"
			trNacionalidad.style.display = "none"
			window.frmCliente.cbxnacionalidad.style.display ="none"
			trFecha.style.display ="none"
			window.frmCliente.txtfechanacimiento.style.display ="none"
			frmCliente.optPersona.checked = 0
		<% End If %>
		
		frmCliente.cbxtarjetas.value = "<%=sTarjetas%>"
		frmCliente.txttarjeta.value = "<%=sTarjeta%>"
		frmCliente.cbxsexo.value = "<%=sSexo%>"
		frmCliente.cbxnacionalidad.value = "<%=sNacionalidad%>"
		frmCliente.txtcorreoelectronico.value = "<%=sCorreoElectronico%>"
		frmCliente.txtFechaNacimiento.value = "<%=sFechaNacimiento%>"	
		frmCliente.txtNumeroCelular.value = "<%=sNumeroCelular%>" ' APPL-9009
	End Sub


	Sub CargarMenuActualizar()
		Dim sId
		
		frmCliente.objmenu.bgColor = document.bgColor 
		frmCliente.objmenu.stylesheet = "../Estilos/Cliente.css"
		sId = frmCliente.objmenu.addparent("Opciones")
		frmCliente.objMenu.addchild sId, "Guardar los cambios", "Guardar", "Principal"
		frmCliente.objMenu.addchild sId, "Agregar beneficiario", "AgregarB", "Principal"				
	End Sub
	
	Sub HabilitarId()
		frmCliente.txtRut.disabled = False
		frmCliente.optRut.disabled = False
		frmCliente.txtPasaporte.disabled = False
		frmCliente.optPasaporte.disabled = False
		frmCliente.cbxPaisPasaporte.disabled = False
	End Sub

	Sub HabilitarTipo()
		frmCliente.optPersona.disabled = False
		frmCliente.optEmpresa.disabled = False
	End Sub
	
	Sub HabilitarCampos()
		If frmCliente.txtRut.value <> "" Then
			frmCliente.txtRut.disabled = True
			frmCliente.optRut.disabled = True
			frmCliente.optPasaporte.disabled = True			
		End If
		If frmCliente.txtPasaporte.value <> "" and frmCliente.cbxPaisPasaporte.value <> "" Then
			frmCliente.txtPasaporte.disabled = True
			frmCliente.optRut.disabled = True
			frmCliente.optPasaporte.disabled = True
			frmCliente.cbxPaisPasaporte.disabled = True
		End If
		if frmCliente.txtTarjeta.value <> "" then
		end if
	End Sub

	Sub cbxPais_onblur()
		Dim sCiudad
		
		If frmCliente.cbxPais.value = "" Then Exit Sub
		If frmCliente.cbxPais.value = "<%=sPais%>" Then Exit Sub
		HabilitarId
		HabilitarTipo
		frmCliente.action = "ActualizacionCliente.asp?Accion=<%=afxAccionPais%>"
		frmCliente.submit 
		frmCliente.action = ""
	End Sub
	
	Sub cbxCiudad_onblur()
		Dim sComuna
		
		If frmCliente.cbxCiudad.value = "" Then Exit Sub
		If frmCliente.cbxCiudad.value = "<%=sCiudad%>" Then Exit Sub		
		HabilitarId
		HabilitarTipo
		frmCliente.action = "ActualizacionCliente.asp?Accion=<%=afxAccionPais%>"
		frmCliente.submit 
		frmCliente.action = ""
	End Sub

	Sub cbxComuna_onblur()
		Dim sComuna
		
		If frmCliente.cbxComuna.value = "" Then Exit Sub
		If frmCliente.cbxComuna.value = "<%=sComuna%>" Then Exit Sub		
		HabilitarId
		HabilitarTipo
		frmCliente.action = "ActualizacionCliente.asp?Accion=<%=afxAccionPais%>"
		frmCliente.submit 
		frmCliente.action = ""
	End Sub
	
	Sub optEmpresa_onClick()
		trEmpresa.style.display = ""
		trPersona.style.display = "none"
		frmCliente.optPersona.checked = 0
		frmCliente.txtApellidoM.value = ""
		frmCliente.txtApellidoP.value = ""
		frmCliente.txtNombres.value = ""
	End Sub
	
	Sub optPersona_onClick()		
		trEmpresa.style.display = "none"
		trPersona.style.display = ""
		frmCliente.optEmpresa.checked = 0
		frmCliente.txtRazonSocial.value = ""
	End Sub

	sub txtTarjeta_onBlur()
		if trim(frmCliente.txtTarjeta.value) <> "" then
			frmCliente.txtTarjeta.value = right("000000" & frmCliente.txtTarjeta.value,6)
		end if
	end sub
-->
</script>

<body >
<!-- #INCLUDE virtual="/Compartido/Encabezado.htm" -->
<form id="frmCliente" method="post">
<input type="hidden" name="txtExchange">
<input type="hidden" name="txtExpress">
<table class="Borde" ID="tabPaso1" CELLSPACING="0" border="0" height="200px" style="position: relative; top: 0px; left: 2px;">
	<tr HEIGHT="15">
		<td colspan="3" class="titulo">Datos del Cliente</td>
	</tr>
	<tr HEIGHT="2">
		<td></td>
		<td COLSPAN="2"></td>		
	</tr>
	<tr>
	<td width="1px"></td>
	<td>
	<table width="100%" border="0" style="HEIGHT: 200px; WIDTH: 300px">
		<tr HEIGHT="40">
			<td></td>
			<td VALIGN="center" colspan="3"><br>
				<table border="0" cellpadding="1">
				<tr>
					<td>
						<input TYPE="radio" name="optRut" CHECKED >Rut
						<input TYPE="radio" name="optPasaporte" >Pasaporte
					</td>
					<td id="lblPaisPasaporte" style="display: none">Pais</td>
				</tr>
				<tr>
					<td>
						<input name="txtRut" style="text-align: right; width: 150px">
						<input name="txtPasaporte" style="width: 150px; display: none">
					</td>
					<td>
						<select name="cbxPaisPasaporte" style="width: 150px; display: none">									
							<%
							    ' INTERNO-1831 - JFMG 28-07-2014								    
							    If Trim(sPasaporte) <> "" Then
							    ' Fin INTERNO-1831 - JFMG 28-07-2014	
							        CargarUbicacion 1, "", sPaisPass
							    ' INTERNO-1831 - JFMG 28-07-2014	
							    Else							     
								    CargarPaisPasaporte sPaisPass 
								End If
								' Fin INTERNO-1831 - JFMG 28-07-2014	
							%>
						</select>
					</td>
				</tr>
				<tr>
					<td>Nº Tarjeta<br>
						<select name="cbxTarjetas">
							<% CargarPrefijoTarjeta sTarjetas %>
						</select>
						<input name="txtTarjeta" value="<%=right(sTarjeta,6)%>" style="width: 87px; text-align:right">						
					</td>
				</tr>
				</table>
			</td>
		</tr>
		<tr></tr>
		<tr id="trTipoCliente" style="display: "><td colspan="3">
		<table >
			<tr>
			<td WIDTH="1"></td>
			<td id="tdTipoCliente">			
				<input TYPE="radio" id="optPersona" name="optPersona" CHECKED disabled>Persona			
				<input TYPE="radio" id="optEmpresa" name="optEmpresa" disabled>Empresa
			</td>			
			</tr>
		</table>		
		</td>
		</tr>
		
		<tr>
			<td width=15></td>
			<td colspan="2">
				<table>
					<tr>
						<td id="trSexo">Sexo<br>
							<select name="cbxSexo">
							<%CargarTipos sSexo, "SEXO"	%>
							</select>
						</td>
					</tr>
					<tr>
						<td id="trNacionalidad">Nacionalidad<br>
							<select name="cbxNacionalidad" style="width: 150px;">
								<%
									CargarUbicacion 1, "", sNacionalidad
								%>
							</select>
						</td>
					</tr>
				</table>
			<td>
		</tr>
		
		<tr ID="trEmpresa" HEIGHT="20" STYLE="DISPLAY: none"><td colspan="3">
		<table>
			<tr>
			<td WIDTH="1"></td>
			<td colspan="3">Razón Social<br>
			<input name="txtRazonSocial" id="txtRazonSocial" SIZE="40" style="width: 350px"  onblur="frmCliente.txtRazonSocial.value=MayMin(frmCliente.txtRazonSocial.value)">*
			</tr>
		</table>
		</td></tr>
		<tr ID="trPersona" HEIGHT="20"><td colspan="3">
			<table>
			<tr>
				<td></td>
				<td colspan="2">Nombres<br>				
				<input NAME="txtNombres" id="txtNombres" SIZE="25" style="width: 350px" onkeypress="IngresarTexto(2)" onblur="frmCliente.txtNombres.value=MayMin(frmCliente.txtNombres.value)">*
			</tr>
			<tr>
				<td></td>
				<td>Apellido Paterno</td>
				<td>Apellido Materno</td>
			</tr>
			<tr>
				<td></td>
				<td><input name="txtApellidoP" id="txtApellidoP" SIZE="20" style="width: 170px" onkeypress="IngresarTexto(2)" onblur="frmCliente.txtApellidoP.value=MayMin(frmCliente.txtApellidoP.value)">*</td>
				<td><input name="txtApellidoM" id="txtApellidoM" SIZE="20" style="width: 170px" onkeypress="IngresarTexto(2)" onblur="frmCliente.txtApellidoM.value=MayMin(frmCliente.txtApellidoM.value)">*</td>			
			</tr>
			</table>
		</td>
		</tr>
		<tr>
			<td width=15></td>
			<td id="trFecha" colspan="2">Fecha de Nacimiento<br>
				<input name="txtFechaNacimiento" size="10" value="<%=sFechaNacimiento%>">(dd/mm/aaaa)
			</td>
		</tr>
		<tr>
			<td></td>
			<td COLSPAN="2">Dirección<br>
				<input STYLE="HEIGHT: 22px; WIDTH: 350px" SIZE="10" NAME="txtDireccion" onkeypress="IngresarTexto(3)" onblur="frmCliente.txtDireccion.value=MayMin(frmCliente.txtDireccion.value)">
			</td>
		</tr>
		<tr>
			<td></td>
			<td>Pais<br>
				<select name="cbxPais" style="width: 170px">
					<%	
						CargarUbicacion 1, "", sPais 	
					%>
				</select>
			</td>
			<td colspan="1">Ciudad<br>
				<select name="cbxCiudad" style="width: 170px">
					<%	
                        CargarCiudadesPais sPais, sCiudad
					%>
				</select>
			</td>
		</tr>
		<tr>
			<td></td>
			<td>Comuna<br>
				<select name="cbxComuna" style="width: 170px">
				 <script></script>
					<%	
						If sPais = "CL" Then
                            CargarComunaCiudad  sCiudad, sComuna		
						End If
					%>
				</select>
			</td>
		</tr>
		
		<!-- APPL-9009-->
		<tr>
			<td width=15></td>
			<td colspan="1">Celular </br>
				<span style="color: Gray;">(+56 9)</span>
				<input name="txtNumeroCelular" maxlength="8" style="width: 83px"onkeypress="IngresarTexto(1)" onblur="frmCliente.txtNumeroCelular.value=MayMin(frmCliente.txtNumeroCelular.value)" />
			</td>			
		</tr>
		<!-- FIN APPL-9009-->
		
		<tr>
			<td width=15></td>
			<td colspan="1">Teléfono<br>
				<input disabled name="txtPaisFono" style="width: 40px">
				<input disabled name="txtAreaFono" style="width: 40px">
				<input name="txtFono" style="width: 83px"onkeypress="IngresarTexto(1)" onblur="frmCliente.txtFono.value=MayMin(frmCliente.txtFono.value)" maxlength="10"><!--APPL-8897 MS 14-05-2015-->
			</td>
			<td colspan="1">Teléfono<br>
				<input disabled name="txtPaisFono2" style="width: 40px">
				<input disabled name="txtAreaFono2" style="width: 40px">
				<input name="txtFono2" style="width: 83px"onkeypress="IngresarTexto(1)" onblur="frmCliente.txtFono2.value=MayMin(frmCliente.txtFono2.value)" maxlength="10"><!--APPL-8897 MS 14-05-2015-->
			</td>
		</tr>
		<tr>
			<td width=15></td>
			<td colspan="2">Correo Electrónico<br>				
				<input name="txtCorreoElectronico" style="width: 250px">
			</td>
		</tr>
		<tr HEIGHT="2">
			<td></td>
		</tr>
	</table>
	</td>
	<td valign="top">
	<table border="0" width="100px" height="100px">
	<tr><td colspan="2"><object align="left" id="objMenu" style="HEIGHT: 46px; LEFT: 0px; POSITION: relative; TOP: 0px; WIDTH: 190px" type="text/x-scriptlet" width="190" VIEWASTEXT border="0"><param NAME="Scrollbar" VALUE="0"><param NAME="URL" VALUE="http:../Scriptlets/Menu.htm"></object></td></tr>
	<tr><td></div></td></tr>
	<tr height="100%"><td></td></tr>
	</table>
	</td>
	</tr>
</table>	
</form>
</body>
<!--#INCLUDE virtual="/Compartido/Rutinas.htm" -->
<script>

	Sub objMenu_OnScriptletEvent(strEventName, varEventData)
		
	   Select Case strEventName
	   
			Case "linkClick"
				If Right(varEventData, 7) = "Guardar" Then
					If Not ValidarDatos Then
						Exit Sub
					End If
					HabilitarId
					HabilitarTipo
					HabilitarControles

					frmCliente.action = "GrabarActualizacionCliente.asp?AFEXchange=<%=sAFEXchange%>&AFEXpress=<%=sAFEXpress%>&TipoCliente=<%=nTipoCliente%>&Giro=<%=Request("Giro")%>&Accion=<%=Request("Accion")%>&Tipo=<%=Request("Tipo")%>"
					frmCliente.submit 
					frmCliente.action = "" 
				
				ElseIf Right(varEventData, 8) = "AgregarB" Then					
					window.open "agregarbeneficiario.asp?cliente=" & frmCliente.txtExpress.value, "", "height= 500, width= 400"
				
				ElseIF Right(varEventData, 6) = "Volver" Then
					If "<%=sAFEXchange%>" <> "" Then
						window.navigate "AtencionClientes.asp?Accion=1&Campo=5&Argumento=" & "<%=sAFEXchange%>"
					End If
	
					If "<%=sAFEXpress%>" <> "" Then
						window.navigate "AtencionClientes.asp?Accion=1&Campo=6&Argumento=" & "<%=sAFEXpress%>"
					End If

				ElseIF Right(varEventData, 3) = "ASC" Then
					window.navigate "AsociarCliente.asp?Nuevo=<%=sAFEXpress%>"
					
				Else
					window.open varEventData, "Principal"
					
				End If
				
		End Select
		
	End Sub

	Function ValidarDatos()
    Dim sRut
		ValidarDatos = False


		If Trim(frmCliente.txtRut.value) = "" And Trim(frmCliente.txtPasaporte.value) = "" Then
			MsgBox "Debe ingresar la identificación del cliente",,"AFEX"
			Exit Function
		End If
		If frmCliente.optPasaporte.checked Then
			If Trim(frmCliente.cbxPaisPasaporte.value) = "" Then
				MsgBox "Debe ingresar el país del pasaporte del cliente",,"AFEX"
				Exit Function
			End If
		End If
		If frmCliente.optPersona.checked  Then
			If Trim(frmCliente.txtNombres.value) = "" Then
				MsgBox "Debe ingresar el nombre del cliente",,"AFEX"
				Exit Function
			End If
			If Trim(frmCliente.txtApellidoP.value) = "" Then
				MsgBox "Debe ingresar el apellido paterno del cliente",,"AFEX"
				Exit Function
			End If
			If Trim(frmCliente.txtApellidoM.value) = "" Then
				MsgBox "Debe ingresar el apellido materno del cliente",,"AFEX"
				Exit Function
			End If
			If Trim(frmCliente.cbxSexo.value) = "" Then
				MsgBox "Debe ingresar el sexo del cliente",,"AFEX"
				frmCliente.cbxSexo.focus 
				Exit Function
			End If
			If Trim(frmCliente.cbxNacionalidad.value) = "" Then
				MsgBox "Debe ingresar la nacionalidad del cliente",,"AFEX"
				frmCliente.cbxNacionalidad.focus 
				Exit Function
			End If
			If Trim(frmCliente.txtFechaNacimiento.value) = "" Then
				MsgBox "Debe ingresar la fecha de nacimiento del cliente",,"AFEX"
				frmCliente.txtFechaNacimiento.focus 
				Exit Function
			End If
		Else
			If Trim(frmCliente.txtRazonSocial.value) = "" Then
				MsgBox "Debe ingresar la razón social del cliente",,"AFEX"
				Exit Function
			End If		
		End If				
		If Trim(frmCliente.txtDireccion.value) = "" Then
			MsgBox "Debe ingresar la dirección del cliente",,"AFEX"
			Exit Function
		End If
		
		If Trim(frmCliente.cbxPais.value) = "" Then
			MsgBox "Debe ingresar el pais del cliente",,"AFEX"
			Exit Function
		End If
		If Trim(frmCliente.cbxCiudad.value) = "" Then
			MsgBox "Debe ingresar la ciudad del cliente",,"AFEX"
			Exit Function
		End If
		
		If Trim(frmCliente.txtFono.value) = "0" Then frmCliente.txtFono.value = ""
		
		'If Trim(frmCliente.txtFono.value) = "" Then
		'	MsgBox "Debe ingresar el teléfono del cliente",,"AFEX"
		'	Exit Function
		'End If
				
		if trim(frmcliente.txtfono.value) <> "" then
		
			If trim(frmcliente.txtfono.value)="1234567" or trim(frmcliente.txtfono.value)="0000000"  or _
			   Trim(frmcliente.txtfono.value)="1111111" or Trim(frmcliente.txtfono.value)="111111" or _
			   trim(frmcliente.txtfono.value)="2222222" or trim(frmcliente.txtfono.value)="222222" or _
			   trim(frmcliente.txtfono.value)="3333333" or trim(frmcliente.txtfono.value)="333333" or _
			   trim(frmcliente.txtfono.value)="4444444" or trim(frmcliente.txtfono.value)="444444" or _ 
			   trim(frmcliente.txtfono.value)="5555555" or trim(frmcliente.txtfono.value)="555555" or _
			   trim(frmcliente.txtfono.value)="6666666" or trim(frmcliente.txtfono.value)="666666" or _
			   trim(frmcliente.txtfono.value)="7777777" or trim(frmcliente.txtfono.value)="777777" or _
			   trim(frmcliente.txtfono.value)="8888888" or trim(frmcliente.txtfono.value)="888888" or _
			   trim(frmcliente.txtfono.value)="9999999" or trim(frmcliente.txtfono.value)="999999" or _
			   trim(frmCliente.txtFono.value)="123456" then
    			MsgBox "El teléfono ingresado no es válido",,"AFEX"
				Exit Function

			End If  
		end if
		
		' APPL-9009
		if trim(frmcliente.txtNumeroCelular.value) <> "" then
		
			If trim(frmcliente.txtNumeroCelular.value)="12345678" or trim(frmcliente.txtNumeroCelular.value)="00000000" or _
			   Trim(frmcliente.txtNumeroCelular.value)="11111111" or trim(frmcliente.txtNumeroCelular.value)="22222222" or _
			   trim(frmcliente.txtNumeroCelular.value)="33333333" or trim(frmcliente.txtNumeroCelular.value)="44444444" or _
			   trim(frmcliente.txtNumeroCelular.value)="55555555" or trim(frmcliente.txtNumeroCelular.value)="66666666" or _
			   trim(frmcliente.txtNumeroCelular.value)="77777777" or trim(frmcliente.txtNumeroCelular.value)="88888888" or _
			   trim(frmcliente.txtNumeroCelular.value)="99999999" or _ 
			   (ccur("0" & trim(frmCliente.txtNumeroCelular.value)) < 30000000 or ccur("0" & trim(frmCliente.txtNumeroCelular.value)) > 99999999) then 'APPL-44816 MS 29-06-2017
    			MsgBox "El celular ingresado no es válido",,"AFEX"
				Exit Function

			End If  
		end if
		' FIN APPL-9009

		ValidarDatos = True
	End Function
	
</script>
</html>

