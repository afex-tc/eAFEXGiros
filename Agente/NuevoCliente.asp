<%@ Language=VBScript %>
<%

    ' JFMG 14-11-2012
    If Request("ClienteeAFN") <> "" Then ' esto porque se está llamando desde otra aplicación (por ejemplo la de VIGO)
        Session("CodigoCliente") = Request("ClienteeAFN") 
    End If
    ' FIN 14-11-2012

 %>

<!-- #INCLUDE virtual="/Compartido/TimeOut.asp" -->
<!-- #INCLUDE virtual="/Compartido/Errores.asp" -->
<!-- #INCLUDE virtual="/Compartido/Rutinas.asp" -->
<%
	Dim sApellidoP, sApellidoM, sNombres, sRazonSocial, sRut, sPasaporte
	Dim sDireccion, sPais, sCiudad, sComuna, sPaisPass
	Dim sNombrePais, sNombreCiudad, sNombreComuna, sNombrePaisPass
	Dim sPaisFono, sAreaFono, sFono, sPaisFono2, sAreaFono2, sFono2
	Dim nAccion, sAFEXchange, sAFEXpress, nMenu
	Dim nTipoCliente, sDisabled, sDisplay, sId
	dim sTarjeta, sSexo, sNacionalidad, sCorreoElectronico, sFechaNacimiento, sTarjetas
	Dim sNumeroCelular ' APPL-9009
	
	Dim sMensajeUsuario, sClienteExisteCorporativa ' JFMG 13-06-2013
	sClienteExisteCorporativa = "0" ' JFMG 13-06-2013
	
	Const afxAccionNada = 0
	Const afxAccionBuscar = 1
	Const afxAccionNuevo = 2
	Const afxAccionActualizar = 3
	Const afxAccionPais = 4
	
	nAccion = cInt(0 & Request("Accion"))
	
	' JFMG 06-08-2008 valida si la identificación que se ingresó ya existe
	if nAccion = 99 then			
		if Request.Form("txtRut") <> "" then
			Set rs = BuscarCliente(1, Request.Form("txtRut"), "", "")
		elseif Request.Form("txtPasaporte") <> "" then
			Set rs = BuscarCliente(2, Request.Form("txtPasaporte"), "", "")
		end if				
		if not rs.eof then
		    ' JFMG 13-06-2013 el cliente ya existe en corporativa pero no en giros		    
		    IF isnull(rs("Express")) then
		        sMensajeUsuario = "El cliente no se encuentra en el Sistema de Giros pero si en Atención Clientes, por lo tanto los datos han sido copiados. Favor verificar su veracidad y Grabar."
		        sClienteExisteCorporativa = "1"
		        ' llena los datos desde corporativa	
		        nTipoCliente = cInt(0 & rs("tipo"))
			    If nTipoCliente = 1 or nTipoCliente = 7 or nTipoCliente = 8 Then
			        nTipoCliente = 1
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
			            'sMensajeUsuario = "El Cliente presenta problemas con su tarjeta Giro Club. Favor actualizar los datos."
			        else
			        ' FIN JFMG 15-05-2013
				        sTarjetas = left(sTarjetas, len(sTarjetas) - 6)
				    ' JFMG 15-05-2013
				    end if
				    ' FIN JFMG 15-05-2013
			    end if			
			    sTarjeta = trim(right(EvaluarVar(rs("tarjeta"),""),6))
			    ssexo = EvaluarVar(rs("sexo"),"")
			    snacionalidad = EvaluarVar(rs("codigonacionalidad"),"")
			    scorreoelectronico = EvaluarVar(rs("correoelectronico"),"")
			    sFechaNacimiento = EvaluarVar(rs("fecha_nacimiento"),"")		       
		        sNumeroCelular = EvaluarVar(rs("NumeroCelularCliente"),"") ' APPL-9009
		    ELSE	 
		    ' FIN JFMG 13-06-2013
		    
			    ' si el cliente existe lo envia a la página de actualización
			    sAFEXpress = rs("Express")
			    rs.close
			    set rs = nothing
			    Response.Redirect "ActualizacionCliente.asp?AFEXpress=" & sAFEXpress
			
			END IF
		    ' FIN JFMG 13-06-2013
		
		ELSE ' JFMG 13-06-2013		
		
		    nAccion = afxAccionPais		
		end if
	end if
	' *************************** FIN *****************************
	
	
	'Response.Redirect "../Compartido/error.asp?description=" & Request("Menu")
	sDisabled = "disableds"
	Select Case nAccion
		Case afxAccionPais
			CargarActualizacion
			sDisabled = ""
			
		Case Else
			sPais = Session("PaisCliente")
			sCiudad = Session("CiudadCliente")
			sPaisFono = Session("PaisCliente")
			sAreaFono = Session("CiudadCliente")			
			
	End Select
	If Session("Categoria") = 4 Then 
		sDisplay = "none"
		sId = "Id"
		sPaisPass = Session("PaisCliente")
	Else
		sDisplay = ""
		sId = "Pasaporte"
	End If
	
	if trim(sTarjetas) = "" then sTarjetas = "0008010"
	
		
	Sub CargarActualizacion()
		sAFEXpress = request("AFEXpress")
		sAFEXchange = request("AFEXchange")
		sRut = request.Form("txtRut")
		sPasaporte = request.Form("txtPasaporte")
		sNombres = request.Form("txtNombres")
		sApellidoP = request.Form("txtApellidoP")
		sApellidoM = request.Form("txtApellidoM")
		sRazonSocial = Request.Form("txtRazonSocial")
		sDireccion = request.Form("txtDireccion")
		sPaisPass = Request.Form("cbxPaisPasaporte")
		sPais = Request.Form("cbxPais")
		sCiudad = Request.form("cbxCiudad")
		sComuna = Request.form("cbxComuna")			
		sPaisFono = Request.form("txtPaisFono")
		sAreaFono = Request.form("txtAreaFono")
		sFono = request.Form("txtFono")
		sPaisFono2 = Request.form("txtPaisFono2")
		sAreaFono2 = Request.form("txtAreaFono2")
		sFono2 = request.Form("txtFono2")
		sTarjetas = request.Form("cbxTarjetas")
		sTarjeta = request.Form("txtTarjeta")
		sSexo = request.Form("cbxSexo")
		sNacionalidad = request.Form("cbxNacionalidad")
		sCorreoElectronico = request.Form("txtCorreoElectronico")
		sFechaNacimiento = request.Form("txtFechaNacimiento")
		sNumeroCelular = request.Form("txtNumeroCelular") ' APPL-9009
		'Response.Redirect "../Compartido/Error.asp?Titulo=" & Request.Form("optPersona")
		If Request.Form("optPersona") = "on" Then
			nTipoCliente = 1
		Else
			nTipoCliente = 2
		End If
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
	sEncabezadoTitulo = "Nuevo Cliente"

	Sub imgBuscar_onMouseOver()
		 imgBuscar.style.cursor = "Hand"
	End Sub	

	Sub window_onLoad()
		
		CargarMenuNuevo
		<% 
			Select Case nAccion 
				Case afxAccionNada %>
					<% If Session("Categoria") = 4 Then %>
							optPasaporte_onClick
							frmCliente.txtPasaporte.focus 
					<% Else %>
							frmCliente.txtRut.focus
					<%	End If %>
					
		<%		Case afxAccionPais %>
					CargarCliente					
		<% End Select %>
		
		<% If Session("Categoria") = 4 Then %>
				frmCliente.cbxNegocio.selectedIndex = 1
		<% End If %>
		
		frmCliente.txtPaisFono.value = "<%=ObtenerDDI(1, sPais)%>" 
		If frmCliente.cbxComuna.value <> "" Then 
			frmCliente.txtAreaFono.value = "<%=ObtenerDDI(3, sComuna)%>"
		Else
			frmCliente.txtAreaFono.value = "<%=ObtenerDDI(2, sCiudad)%>"
		End If
		
		frmCliente.txtPaisFono2.value = frmCliente.txtPaisFono.value
		frmCliente.txtAreaFono2.value = frmCliente.txtAreaFono.value
		
		frmCliente.cbxTarjetas.value = "<%=sTarjetas%>"
		
		' JFMG 13-06-2013
		If "<%=sMensajeUsuario%>" <> "" Then
		   msgbox "<%=sMensajeUsuario%>",,"AFEX" 
		End If
		If "<%=sClienteExisteCorporativa%>" = "1" Then
		    CargarCliente
		End If
		' FIN JFMG 13-06-2013
	End Sub

	Sub optRut_onClick()
		window.frmcliente.optpasaporte.checked=False
		window.frmcliente.optRut.checked=true
		window.frmcliente.txtpasaporte.style.display="none"
		window.frmcliente.cbxPaisPasaporte.style.display="none"
		lblPaisPasaporte.style.display="none"
		window.frmCliente.txtRut.style.display = ""		
	End Sub
	
	Sub optPasaporte_onClick()
		window.frmcliente.optRut.checked=False
		window.frmCliente.optPasaporte.checked = True
		window.frmCliente.txtRut.style.display = "none"		
		window.frmcliente.txtpasaporte.style.display=""
		window.frmcliente.cbxPaisPasaporte.style.display=""
		lblPaisPasaporte.style.display=""
	End Sub


	Sub CargarCliente()
		frmCliente.txtApellidoM.value = "<%=sApellidoM%>"
		frmCliente.txtApellidoP.value = "<%=sApellidoP%>"
		frmCliente.txtNombres.value = "<%=sNombres%>"
		frmCliente.txtRazonSocial.value = "<%=sRazonSocial%>"
		frmCliente.txtDireccion.value = "<%=sDireccion%>"
		frmCliente.txtFono.value = "<%=sFono%>"
		frmCliente.txtFono2.value = "<%=sFono2%>"
		frmCliente.txtTarjeta.value = "<%=sTarjeta%>"
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
		<% If nTipoCliente <> afxPersona Then %>
			window.frmCliente.optEmpresa.checked = True
			trEmpresa.style.display = ""
			trPersona.style.display = "none"
			frmCliente.optPersona.checked = 0
			trfecha.style.display = "none"
			trsexo.style.display="None"
			trnacionalidad.style.display="none"
			window.frmCliente.txtfechanacimiento.style.display="none"
			window.frmCliente.cbxsexo.style.display="none"
			window.frmCliente.cbxnacionalidad.style.display="none"
		
		<% End If %>
		frmCliente.cbxsexo.value = "<%=sSexo%>"
		frmCliente.cbxnacionalidad.value = "<%=sNacionalidad%>"
		frmCliente.txtcorreoelectronico.value = "<%=sCorreoElectronico%>"
		frmCliente.txtFechaNacimiento.value = "<%=sFechaNacimiento%>"	
		frmCliente.txtNumeroCelular.value = "<%=sNumeroCelular%>" ' APPL-9009
	End Sub

	Sub CargarMenuNuevo()
		Dim sId
		
		frmCliente.objmenu.bgColor = document.bgColor 
		frmCliente.objmenu.stylesheet = "../Estilos/Cliente.css"
		sId = frmCliente.objmenu.addparent("Opciones")
		frmCliente.objMenu.addchild sId, "Agregar el cliente", "Agregar", "Principal"
				
	End Sub

	Sub HabilitarDireccion()
		frmCliente.txtDireccion.disabled = False
		frmCliente.cbxPais.disabled = False
		frmCliente.cbxCiudad.disabled = False
		frmCliente.cbxComuna.disabled = False
		frmCliente.txtPaisFono.disabled = False
		frmCliente.txtAreaFono.disabled = False
		frmCliente.txtFono.disabled = False
		frmCliente.txtPaisFono2.disabled = False
		frmCliente.txtAreaFono2.disabled = False
		frmCliente.txtFono2.disabled = False
		frmCliente.txtNumeroCelular.disabled = False ' APPL-9009
	End Sub
	
	Sub HabilitarId()
		frmCliente.txtRut.disabled = False
		frmCliente.optRut.disabled = False
		frmCliente.txtPasaporte.disabled = False
		frmCliente.optPasaporte.disabled = False
		frmCliente.cbxPaisPasaporte.disabled = False
	End Sub
		
	Sub HabilitarCampos()
		HabilitarDireccion 
		If frmCliente.txtRut.value <> "" And <%=(nMenu<>afxMenuNuevo)%> Then
			frmCliente.txtRut.disabled = True
			frmCliente.optRut.disabled = True
			frmCliente.optPasaporte.disabled = True
		End If
		If frmCliente.txtPasaporte.value <> "" And <%=(nMenu<>afxMenuNuevo)%> Then
			frmCliente.txtPasaporte.disabled = True
			frmCliente.optRut.disabled = True
			frmCliente.optPasaporte.disabled = True
			frmCliente.cbxPaisPasaporte.disabled = True
		End If
	End Sub

	Sub cbxPais_onblur()
		Dim sCiudad
		
		If frmCliente.cbxPais.value = "" Then Exit Sub
		If frmCliente.cbxPais.value = "<%=sPais%>" Then Exit Sub
		HabilitarDireccion
		HabilitarId 
		frmCliente.action = "NuevoCliente.asp?Accion=<%=afxAccionPais%>&Referencia=" & RemplazaLetra("<%=request("Referencia")%>")
		frmCliente.submit 
		frmCliente.action = ""
	End Sub
	
	Sub cbxCiudad_onblur()
		
		If frmCliente.cbxCiudad.value = "" Then Exit Sub
		If frmCliente.cbxCiudad.value = "<%=sCiudad%>" Then Exit Sub		
		HabilitarDireccion
		HabilitarId
		frmCliente.action = "NuevoCliente.asp?Accion=<%=afxAccionPais%>&Referencia=" & RemplazaLetra("<%=request("Referencia")%>")
		frmCliente.submit 
		frmCliente.action = ""
	End Sub

	Sub cbxComuna_onblur()		
		If frmCliente.cbxComuna.value = "" Then Exit Sub
		If frmCliente.cbxComuna.value = "<%=sComuna%>" Then Exit Sub		
		HabilitarDireccion
		HabilitarId
		frmCliente.action = "NuevoCliente.asp?Accion=<%=afxAccionPais%>&Referencia=" & RemplazaLetra("<%=request("Referencia")%>")
		frmCliente.submit 
		frmCliente.action = ""
	End Sub
	
	Sub optEmpresa_onClick()
		trEmpresa.style.display = ""
		trPersona.style.display = "none"
		trfecha.style.display = "none"
		trsexo.style.display="None"
		trnacionalidad.style.display="none"
		window.frmCliente.txtfechanacimiento.style.display="none"
		window.frmCliente.cbxsexo.style.display="none"
		window.frmCliente.cbxnacionalidad.style.display="none"
		frmCliente.optPersona.checked = 0
			LimpiarControles
	End Sub
	
	Sub optPersona_onClick()		
		trEmpresa.style.display = "none"
		trPersona.style.display = ""
		trfecha.style.display = ""
		trsexo.style.display=""
		trnacionalidad.style.display=""
		frmCliente.optEmpresa.checked = 0
		window.frmCliente.txtfechanacimiento.style.display=""
		window.frmCliente.cbxsexo.style.display=""
		window.frmCliente.cbxnacionalidad.style.display=""
		
			LimpiarControles
	End Sub

	Sub LimpiarControles
		frmCliente.txtApellidoM.value = ""
		frmCliente.txtApellidoP.value = ""
		frmCliente.txtNombres.value = ""
		frmCliente.txtRazonSocial.value = ""
	End Sub	
	
	Sub txtRut_onBlur()
		Dim sRut
		
		If frmCliente.txtRut.value = "" Then Exit Sub
		sRut = ValidarRut(frmCliente.txtRut.value)
		If sRut = Empty Then
			msgbox "El número de Rut no es válido"
			frmCliente.txtRut.value = ""
			frmCliente.txtRut.focus()
			exit sub
		Else
			frmCliente.txtRut.value = sRut
		End If
		
		
		' JFMG 06-08-2008 se rellama para ver si el rut existe
		if frmCliente.txtRut.value = "" then exit sub
		
		frmCliente.action = "NuevoCliente.asp?Accion=99&Referencia=" & RemplazaLetra("<%=request("Referencia")%>")
		frmCliente.submit 
		frmCliente.action = ""
		' **************************** Fin *********************
	End Sub
	
	Sub txtPasaporte_onBlur()
		if frmCliente.txtPasaporte.value = "" then exit sub
				
		' JFMG 06-08-2008 se rellama para ver si el pasaporte existe
		frmCliente.action = "NuevoCliente.asp?Accion=99&Referencia=" & RemplazaLetra("<%=request("Referencia")%>")
		frmCliente.submit 
		frmCliente.action = ""
		' **************************** Fin *********************
	End Sub
	
	sub txtTarjeta_onBlur()
		if trim(frmCliente.txtTarjeta.value) <> "" then
			frmCliente.txtTarjeta.value = right("000000" & frmCliente.txtTarjeta.value,6)
		end if
	end sub

    function RemplazaLetra(Cadena)
		cadena = replace(cadena, "Ñ", "%c3%91")
		cadena = replace(cadena, "Ã‘", "%c3%91")
		RemplazaLetra = replace(cadena, "ñ", "%c3%b1")
	end function
	
	function RemplazaPorciento(Cadena)
		cadena = replace(cadena, "%c3%91", "Ñ")
		RemplazaPorciento = replace(cadena, "%c3%b1", "ñ")
	end function

	
-->
</script>

<body >
<!-- #INCLUDE virtual="/Compartido/Encabezado.htm" -->
<form id="frmCliente" method="post">
<table class="Borde" ID="tabPaso1" CELLSPACING="0" border="0" height="200px" style="position: relative; top: 0px; left: 2px;">
	<tr HEIGHT="15">
		<td colspan="3" class="titulo">Nuevo Cliente.</td>
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
						<a style="display: <%=sDisplay%>"><input TYPE="radio" name="optRut" CHECKED >Rut</a>
						<input TYPE="radio" name="optPasaporte"><%=sId%>
					</td>
					<td id="lblPaisPasaporte" style="display: none">Pais</td>
				</tr>
				<tr>
					<td>
						<input name="txtRut" style="width: 150px; text-align:right" OnKeyPress="" OnMouseOver="frmCliente.txtRut.value=FormatoRut(frmCliente.txtRut.value)">
						<input name="txtPasaporte" style="width: 150px; display: none" onblur="frmCliente.txtPasaporte.value=MayMin(frmCliente.txtPasaporte.value)">
					</td>
					<td>
						<select name="cbxPaisPasaporte" style="width: 150px; display: none">									
							<%
									CargarPaisPasaporte sPaisPass ' INTERNO-1831 - JFMG 28-07-2014								
							%>
						</select>*
					</td>
				</tr>
				<tr>
					<td>Nº Tarjeta<br>
						<input type="hidden" name="txtTarjeta1">
						<select name="cbxTarjetas">		
							<% CargarPrefijoTarjeta sTarjetas %>
						</select>
						<input name="txtTarjeta" style="width: 87px; text-align:right">
					</td>
				</tr>
				</table>
			</td>
		</tr>
		<tr></tr>
		<tr id="trTipoCliente" style="display: "><td colspan="3">
		<table id="tbTipoCliente" >
			<tr>
			<td WIDTH="1"></td>
			<td>
				<input TYPE="radio" id="optPersona" name="optPersona" CHECKED>Persona			
				<input TYPE="radio" id="optEmpresa" name="optEmpresa" >Empresa
			</td>			
			</tr>
		</table>		
		</td></tr>
		
		<tr>
			<td width=15></td>
			<td colspan="2">
				<table>
					<tr>
						<td id="trSexo" >Sexo<br>
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
			<input name="txtRazonSocial" id="txtRazonSocial" SIZE="40" style="width: 350px" onkeypress="IngresarTexto(2)" onblur="frmCliente.txtRazonSocial.value=MayMin(frmCliente.txtRazonSocial.value)">*
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
		</td></tr>
		
			<tr>
				<td width=15></td>
				<td id="trfecha" colspan="2">Fecha de Nacimiento<br>
					<input name="txtFechaNacimiento" size="10" value="<%=sFechaNacimiento%>">(dd/mm/aaaa)
				</td>
			</tr>
		
	
		<tr style="display: <%=sDisplay%>">
			<td></td>
			<td COLSPAN="2">Dirección<br>
				<input STYLE="HEIGHT: 22px; WIDTH: 350px" SIZE="10" NAME="txtDireccion" onblur="frmCliente.txtDireccion.value=MayMin(frmCliente.txtDireccion.value)">
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
				<select  name="cbxCiudad" style="width: 170px">
					<%	
							CargarCiudadesPais sPais, sCiudad 
					%>
				</select>
			</td>
		</tr>
		<tr style="display: <%=sDisplay%>">
			<td></td>
		<td>Comuna<br>
			<select  name="cbxComuna" style="width: 170px">
			 <script></script>
				<%	
						If sPais = "CL" Then						
							CargarComunaCiudad sCiudad, sComuna
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
				<input name="txtFono" style="width: 83px"onkeypress="IngresarTexto(1)" onblur="frmCliente.txtFono.value=MayMin(frmCliente.txtFono.value)">
			</td>
			<td colspan="1">Teléfono<br>
				<input disabled name="txtPaisFono2" style="width: 40px">
				<input disabled name="txtAreaFono2" style="width: 40px">
				<input name="txtFono2" style="width: 83px"onkeypress="IngresarTexto(1)" onblur="frmCliente.txtFono2.value=MayMin(frmCliente.txtFono2.value)">
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
				<tr><td colspan="2"><object align="left" id="objMenu" style="HEIGHT: 34px; LEFT: 0px; POSITION: relative; TOP: 0px; WIDTH: 190px" type="text/x-scriptlet" width="190" VIEWASTEXT border="0"><param NAME="Scrollbar" VALUE="0"><param NAME="URL" VALUE="http:../Scriptlets/Menu.htm"></object></td></tr>
				<tr style="display: <%=sDisplay%>">
					<td colspan="2">Negocio<br>
						<select name="cbxNegocio" style="width: 100px;">
							<option selected value="0"></option>
							<option value="1">Giros</option>
							<option value="2">Cambios</option>
						</select>
					</td>					
				</tr>
	<tr height="100%"><td></td></tr>
	</table>
	</td>
	</tr>
</table>

</form>
</body>
<!--#INCLUDE virtual="/Compartido/Rutinas.htm" -->

<script language="vbscript" type="text/vbscript">

    
	Sub objMenu_OnScriptletEvent(strEventName, varEventData)
		
	   Select Case strEventName
	   
			Case "linkClick"
				If Right(varEventData, 7) = "Agregar" Then
					If Not ValidarDatos Then
						Exit Sub
					End If
					frmCliente.txtAreaFono.disabled = False
					frmCliente.txtPaisFono.disabled = False
					frmCliente.txtAreaFono2.disabled = False
					frmCliente.txtPaisFono2.disabled = False
					frmCliente.txtTarjeta1.disabled = False
					frmCliente.txtTarjeta1.value = frmCliente.cbxTarjetas.value
					
					dim referencia
					referencia = replace("<%=request("Referencia")%>", "Ñ", "%c3%91")
		            referencia = replace(referencia, "Ã‘", "%c3%91")
		            referencia = replace(referencia, "ñ", "%c3%b1")
					
					frmCliente.action = "GrabarNuevoCliente.asp?Refrencia=" & referencia
					frmCliente.submit 
					frmCliente.action = "" 
					
				Else
					window.open varEventData, "Principal"
					
				End If
				
		End Select
		
	End Sub

	Function ValidarDatos()
		
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
		<% If Session("Categoria") = 4 Then %>
				ValidarDatos = True
				Exit Function
		<% End If %>
		
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
		'If Trim(frmCliente.txtFono.value) = "" Then
		'	MsgBox "Debe ingresar el teléfono del cliente",,"AFEX"
		'	Exit Function
		'End If				
		If Trim(frmCliente.cbxNegocio.value) = "0" Then
			MsgBox "Debe seleccionar un negocio para el cliente",,"AFEX" 
			Exit Function
		End If
		
		'APPL-5076 MS 25-04-2014: Si falla la validación del evento onBlur del rut, se realiza una doble validación.
        If frmCliente.optRut.checked Then
			sRut = ValidarRut(frmCliente.txtRut.value)
			If sRut = "" Then
			    msgbox "Debe ingresar un número de rut válido ", vbOKOnly + vbInformation, "Ingreso de Cliente"
			    frmCliente.txtRut.focus
			    Exit function
		    End If
		End if
        'FIN APPL-5076 MS 25-04-2014
		
		' APPL-9009
		If trim(frmcliente.txtNumeroCelular.value) <> "" Then		
			If trim(frmcliente.txtNumeroCelular.value)="12345678" or trim(frmcliente.txtNumeroCelular.value)="00000000" or _
			   Trim(frmcliente.txtNumeroCelular.value)="11111111" or trim(frmcliente.txtNumeroCelular.value)="22222222" or _
			   trim(frmcliente.txtNumeroCelular.value)="33333333" or trim(frmcliente.txtNumeroCelular.value)="44444444" or _
			   trim(frmcliente.txtNumeroCelular.value)="55555555" or trim(frmcliente.txtNumeroCelular.value)="66666666" or _
			   trim(frmcliente.txtNumeroCelular.value)="77777777" or trim(frmcliente.txtNumeroCelular.value)="88888888" or _
			   trim(frmcliente.txtNumeroCelular.value)="99999999" or _ 
			   (ccur("0" & trim(frmCliente.txtNumeroCelular.value)) < 30000000 or ccur("0" & trim(frmCliente.txtNumeroCelular.value)) > 99999999) Then 'APPL-44816 MS 29-06-2017
    			MsgBox "El celular ingresado no es válido",,"AFEX"
				Exit Function
			End If  
		End If
		' FIN APPL-9009
		
		ValidarDatos = True
	End Function
	
</script>
</html>
