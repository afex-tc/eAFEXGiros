<%@ Language=VBScript %>
<%
	Response.Expires = 0
	If Session("CodigoCliente") = "" Then
		response.Redirect "../Compartido/TimeOut.htm"
		response.end
	End If
	estado=request.QueryString("estado")
	dd=request.form("txtExpress")
		
%>
<!-- #INCLUDE virtual="/Compartido/Errores.asp" -->
<!-- #INCLUDE virtual="/Compartido/Rutinas.asp" -->
<!-- #INCLUDE virtual="/Compartido/Constantes.asp" -->
<!-- #INCLUDE virtual="/Compartido/RutinasEncriptar.asp" -->
<%
	Dim sMensajeUsuario ' JFMG 16-05-2013 para indicar mensajes al usuario(cajero)

	' JFMG 17-12-2009 saca caracter para pasar parametros
	function RemplazaLetra(Cadena)							
		cadena = replace(cadena, "Ñ", "%c3%91")
		RemplazaLetra = replace(cadena, "ñ", "%c3%b1")
	end function
	' *************** FIN 17-12-2009 ****************

	
	Dim sApellidoP, sApellidoM, sNombres, sRazonSocial, sRut, sPasaporte
	Dim sDireccion, sPais, sCiudad, sComuna, sPaisPass, sNombreCompleto,sFechaNac
	Dim sNombrePais, sNombreCiudad, sNombreComuna, sNombrePaisPass, sNacionalidad
	Dim sPaisFono, sAreaFono, sFono, sPaisFono2, sAreaFono2, sFono2
	Dim nAccion, sAFEXchange, sAFEXpress, sArgumento, sArgumento2, sArgumento3
	Dim nTipoCliente, sDisabled, bCliente, sDisplay
	Dim nCampo, sId, sDisplay2
	dim sTarjeta, sTarjetas
	
	Dim sNumeroCelular, sSexo ' APPL-9009
	
	nCampo = cInt(0 & request("Campo"))
	sArgumento = request("Argumento")
	sArgumento2 = request("Argumento2")
	sArgumento3 = request("Argumento3")	
	
	nAccion = cInt(0 & Request("Accion"))
	nTipoCliente = 1
	
	Select Case nAccion
	    Case 91
	        dim parametros
	        parametros = RemplazaLetra(EncriptarCadena(Session("NombreUsuarioOperador"))) & "|" & _
	                    RemplazaLetra(EncriptarCadena(Session("CodigoAgente"))) & "|" & _
	                    RemplazaLetra(EncriptarCadena(Session("CodigoCliente"))) & "|" & _
	                    RemplazaLetra(EncriptarCadena(trim(request.Form("txtGiro"))))

	        if Session("Categoria") = 3 then
                response.Redirect (Session("URLeAFEXNetAGENTES") & "?Referencia=" & parametros)
            else
                response.Redirect (Session("URLeAFEXNet") & "?Referencia=" & parametros)
            end if
	        
		Case afxAccionBuscar
			nCampo = cInt(0 & request("Campo"))
			sArgumento = request("Argumento")
			sArgumento2 = request("Argumento2")
			sArgumento3 = request("Argumento3")		
			Session("ATCAFEXpress") = ""
			Session("ATCAFEXchange") = ""
			CargarCliente
			
		Case afxAccionIngresarMG, afxAccionIngresarTX, afxAccionIngresarUT
			nCampo = cInt(0 & request("Campo"))
			sArgumento = request("Argumento")
			sArgumento2 = request("Argumento2")
			sArgumento3 = request("Argumento3")		
			Session("ATCAFEXpress") = ""
			Session("ATCAFEXchange") = ""
			sDisplay = "none"
			CargarCliente
			
		Case afxAccionNuevo 
			Session("ATCAFEXpress") = ""
			Session("ATCAFEXchange") = ""
			CargarActualizacion
			
		Case afxAccionActualizar 
			Session("ATCAFEXpress") = ""
			Session("ATCAFEXchange") = ""
			CargarActualizacion 
			
		Case afxAccionClienteActual
			sArgumento2 = ""
			sArgumento3 = ""
			If Trim(Session("ATCAFEXchange")) <> "" Then
				nCampo = afxCampoCodigoExchange
				sArgumento = Session("ATCAFEXchange")
			Else
				nCampo = afxCampoCodigoExpress
				sArgumento = Session("ATCAFEXpress")
			End If
			CargarCliente
			
		Case Else
			Session("ATCAFEXchange") = ""
			Session("ATCAFEXpress") = ""
						
	End Select
	If Trim(sRut & sPasaporte) <> "" Then
		Session("IdCliente") = 1
	Else
		Session("IdCliente") = 0
	End If

	If sAFEXchange <> "" Or sAFEXpress <> "" Then
		bCliente = True
	Else
		bCliente = False
	End If
	If Session("Categoria") = 4 Then
		sDisplay2 = "none"
		sId = "Id"
		If sPaisPass = "" Then sPaisPass = Session("PaisCliente")		
	Else
		sDisplay2 = ""
		sId = "Pasaporte"
	End If
	
	if trim(sTarjetas) = "" then sTarjetas = "0008010"
	
	Sub CargarCliente()
		Dim rs 
		
		If nCampo = 0 Then Exit Sub	
				
		Set rs = BuscarCliente(nCampo, sArgumento, sArgumento2, sArgumento3)		
        		
		If rs.EOF Then
			rs.Close
			Set rs = Nothing
                Exit Sub
            End If
		If rs.RecordCount > 1 Then
			rs.Close
			Set rs = Nothing
    
                Response.Redirect "ListaClientes.asp?Accion=1&Campo=" & nCampo & _
									"&Argumento=" & sArgumento & _
								   "&Argumento2=" & sArgumento2 & _
									"&Argumento3=" & sArgumento3 & _
									"&Titulo=Lista de Clientes"
                End If
    
		If Not rs.EOF Then
			    If Session("Categoria") = 4 Then
				If Trim(rs("codigo_pais")) <> Trim(Session("PaisCliente")) Then
					rs.Close
					Set rs = Nothing
					    Exit Sub
				    End If
			    End If
			nTipoCliente = cInt(0 & rs("tipo"))
			    If nTipoCliente = 1 Then
				sApellidoP = MayMin(EvaluarVar(rs("paterno"), ""))
				sApellidoM = MayMin(EvaluarVar(rs("materno"), ""))
				sNombres = MayMin(EvaluarVar(rs("nombre"), ""))
                    Else
				sRazonSocial = MayMin(EvaluarVar(rs("nombre_completo"), ""))
                    End If
                
			sNombreCompleto = MayMin(EvaluarVar(rs("nombre_completo"), ""))
			sRut = FormatoRut(EvaluarVar(rs("rut"), ""))
			sPasaporte = EvaluarVar(rs("pasaporte"), "")
			sPaisPass = EvaluarVar(rs("codigo_paispas"), "")
            'CUM-505 MS 15-02-2016
            if nCampo = 1 and Trim(sPasaporte) <> "" then
                sPasaporte = ""
                sPaisPass = ""
            End if
                 
            if nCampo = 2 and Trim(sRut) <> "" then
                sRut = ""
            End if
            'FIN CUM-505 MS 15-02-2016
			sDireccion = MayMin(EvaluarVar(rs("direccion"), ""))
			sFechaNac= (EvaluarVar(rs("fecha_nacimiento"),""))
			sNacionalidad = (EvaluarVar(rs("nacionalidad"),""))
			sPais = EvaluarVar(rs("codigo_pais"), "")
			sCiudad = EvaluarVar(rs("codigo_ciudad"), "")
			sComuna = EvaluarVar(rs("codigo_comuna"), "")
			sPaisFono = EvaluarVar(rs("ddi_pais"), "")
			sAreaFono = EvaluarVar(rs("ddi_area"), "")
			sFono = EvaluarVar(rs("telefono"), "")
			sPaisFono2 = EvaluarVar(rs("ddi_pais2"), "")
			sAreaFono2 = EvaluarVar(rs("ddi_area2"), "")
			sFono2 = EvaluarVar(rs("telefono2"), "")
			sTarjetas = trim(EvaluarVar(rs("tarjeta"),""))
			sNumeroCelular = EvaluarVar(rs("NumeroCelularCliente"),"") ' APPL-9009
			sSexo = EvaluarVar(rs("Sexo"),"") ' APPL-9009
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
			sTarjeta = right(EvaluarVar(rs("tarjeta"), ""),6)
			sNombrePais = MayMin(EvaluarVar(rs("pais"), ""))
			sNombreCiudad = MayMin(EvaluarVar(rs("ciudad"), ""))
			sNombreComuna = MayMin(EvaluarVar(rs("comuna"), ""))
			sNombrePaisPass = MayMin(EvaluarVar(rs("paispas"), ""))
			sAFEXchange = rs("Exchange")
			sAFEXpress = rs("Express")
			Session("ATCAFEXpress") = rs("Express")
			Session("ATCAFEXchange") = rs("Exchange")
			If Err.number <> 0 Then
				MostrarErrorMS ""
				    End If
			    End If
		rs.Close
		Set rs = Nothing
	End Sub
	
	Sub CargarActualizacion()
		sAFEXpress = Request.form("txtExpress")
		sAFEXchange = Request.form("txtExchange")
		sNombreCompleto = Request.form("txtNombreCompleto")
		sRut = request.Form("txtRut")
		sPasaporte = request.Form("txtPasaporte")
		sNombres = request.Form("txtNombres")
		sApellidoP = request.Form("txtApellidoP")
		sApellidoM = request.Form("txtApellidoM")
		sNacionalidad=Request.Form ("txtNacionalidad")
		sFechaNac = Request.Form ("txtfnac")
		sDireccion = request.Form("txtDireccion")
            sPaisPass = Request.Form("cbxPaisPasaporte")
		sPais = Request.Form("cbxPais")
		sCiudad = Request.form("cbxCiudad")
		sComuna = Request.form("cbxComuna")			
		sPaisFono = request.Form("txtPaisFono")
		sAreaFono = request.Form("txtAreaFono")
		sFono = request.Form("txtFono")
		sPaisFono2 = request.Form("txtPaisFono2")
		sAreaFono2 = request.Form("txtAreaFono2")
		sFono2 = request.Form("txtFono2")
		sTarjetas = Request.form("cbxTarjetas")
		sTarjeta = Request.form("txtTarjeta")
		sNumeroCelular = Request.form("txtNumeroCelular") ' APPL-9009
		sSexo = Request.form("cbxSexo") ' APPL-9009
		If Request.Form("optPersona") = "on" Then
			nTipoCliente = 1
		Else
			nTipoCliente = 2
		End If
	End Sub

%>
<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link rel="stylesheet" type="text/css" href="../Estilos/Cliente.css">


    <style >
        
        .ContenedorModal
        {
           display: none;
           width: 100%;
           height: 100%; 
           top: 0;
           left: 0;
           
           position: fixed;
           z-index: 5000;
        }
    
        .Modal
        {
           width: 100%;
           height: 100%; 
           top: 0;
           left: 0;
           
           position: fixed;
           background-color: White;
           z-index: 5000;
           opacity: .5;
           filter: alpha(opacity=50);
        }

        .InteriorModal300
        {
           position: absolute;
           background-color: #eeeeee;
           border-radius: 4px;
           
           padding: 2px;
           
           width: 300px;
           top: 55%;
           left: 50%;
           margin-top: -100px;
           margin-left: -200px;   
           
           z-index: 5001;
        }
    </style>


</head>
<script LANGUAGE="VBScript">
<!--
	'Variables de módulo
	'Variables para encabezado
	Dim sEncabezadoFondo
	Dim sEncabezadoTitulo
	Dim sId
	Dim sExchange, sExpress, bOk
	
	sExchange = "<%=sAFEXchange%>"
	<% If nAccion = afxAccionBuscar Then %>	
	<% End If %>
	
	Const afxCodigoCliente = 3
	Const afxRut = 1
	Const afxPasaporte = 2
	COnst afxNombres = 4
	Const afxCodigoExchange = 5
	Const afxCodigoExpress = 6
	Const afxTelefono = 7
	Const afxTelefono2 = 8
		
	sEncabezadoFondo = "Principal"
	sEncabezadoTitulo = "Atención de Clientes"

	Sub window_onLoad()
		<% 
			Select Case nAccion 
				Case afxAccionNada
		%>
		<%		Case afxAccionBuscar %>					
					CargarCliente
					
		<%		Case afxAccionNuevo %>					
					CargarCliente
					
		<%		Case afxAccionActualizar %>
					CargarCliente
					
		<%		Case afxAccionIngresarMG %>
					HabilitarDireccion 
					HabilitarId 
					frmCliente.action = "IngresarGiroMG.asp?Tipo=1&ag=<%=Session("CodigoMGPago")%>"
					frmCliente.submit 
					frmCliente.action = "" 
		<%		Case afxAccionClienteActual %>
					CargarClienteActual	'miki SMC-80 MM 2016-04-25						
		<% End Select %>
		
		<% If bCliente Then %>
				<% If Session("Categoria") = 4 Then %>
						CargarMenuClienteInternacional
				<% Else %>
						CargarMenuCliente
				<% End If %>
		<% Else %>
				<% If Session("Categoria") = 4 Then %>
						CargarMenuInternacional
				<% Else %>
						CargarMenu 
				<% End If %>
		<% End If %>
		<% If Session("Categoria") = 4 Then %>
				frmCliente.optPasaporte.checked=True
				frmCliente.optRut.checked=False
				frmCliente.txtRut.style.display = "none"		
				frmcliente.txtpasaporte.style.display=""
				frmcliente.cbxPaisPasaporte.style.display=""
				lblPaisPasaporte.style.display=""
				frmCliente.txtPasaporte.select
		<% Else %>
				frmCliente.txtRut.select
		<% End If %>
		
		if trim(frmCliente.txtTarjeta.value) <> "" then
			window.showModalDialog "../compartido/tarjeta.asp"
		end if
		
		frmCliente.cbxTarjetas.value = "<%=sTarjetas%>"
		
		' JFMG 16-05-2013
		<%If sMensajeUsuario <> "" then %>
		    msgbox "<%=sMensajeUsuario%>",,"AFEX"
		<%End If %>
		' FIN JFMG 16-05-2013
		
	End Sub

	Sub optRut_onClick()
		window.frmcliente.optRut.checked=True
		window.frmcliente.optpasaporte.checked=False
		window.frmcliente.txtpasaporte.style.display="none"
		window.frmcliente.cbxPaisPasaporte.style.display="none"
		lblPaisPasaporte.style.display="none"
		window.frmCliente.txtRut.style.display = ""		
		If Trim(frmCliente.txtExchange.value  & frmCliente.txtExpress.value) <> "" Then
			frmCliente.action = "AtencionClientes.asp"
			frmCliente.submit 
			frmCliente.action = "" 
		End If
		<% If nMenu = afxMenuNormal Then %> 
			LimpiarControles
		<% End If %>
	End Sub
	
	Sub optPasaporte_onClick()
		window.frmcliente.optpasaporte.checked=True
		window.frmcliente.optRut.checked=False
		window.frmCliente.txtRut.style.display = "none"		
		window.frmcliente.txtpasaporte.style.display=""
		window.frmcliente.cbxPaisPasaporte.style.display=""
		lblPaisPasaporte.style.display=""
		If Trim(frmCliente.txtExchange.value  & frmCliente.txtExpress.value) <> "" Then
			frmCliente.action = "AtencionClientes.asp"
			frmCliente.submit 
			frmCliente.action = "" 
		End If
		<% If nMenu = afxMenuNormal Then %> 
			LimpiarControles
		<% End If %>
	End Sub
    
	Sub txtRut_onBlur()
		Dim sRut
		
		If frmCliente.txtRut.value = "" Then Exit Sub
		sRut = ValidarRut(frmCliente.txtRut.value)
		If sRut = Empty Then
			msgbox "El número de Rut no es válido"
			frmCliente.txtRut.select
			frmCliente.txtRut.focus()
		Else
			frmCliente.txtRut.value = sRut
		End If
		
	End Sub

	Sub imgBuscar_onClick()
		frmCliente.imgBuscar.disabled = true
		Buscar <%=afxAccionBuscar%>				
	End Sub
	
	Sub Buscar(ByVal Accion)
	    window.divProcesando.style.display = "block"	    
	
	    Dim nCampo, sArgumento, sArgumento2, sArgumento3
		nCampo = 0
		If Trim(frmCliente.txtRut.value) <> "" Then			
			nCampo = afxRut 
			sArgumento = frmCliente.txtRut.value
		
		elseIf Trim(frmCliente.txtTarjeta.value) <> "" Then
			nCampo = 9			
			sArgumento = trim(frmCliente.cbxTarjetas.value) & trim(frmCliente.txtTarjeta.value)			
			
		ElseIf Trim(frmCliente.txtPasaporte.value) <> "" Then
                nCampo = afxPasaporte 
			sArgumento = frmCliente.txtPasaporte.value	
			
		ElseIf trim(frmCliente.txtFono.value) <> "" Then
			nCampo = afxTelefono
			sArgumento = Trim(frmCliente.txtFono.value)	
			
		ElseIf trim(frmCliente.txtFono2.value) <> "" Then
			nCampo = afxTelefono2
			sArgumento = Trim(frmCliente.txtFono2.value) 	
		ElseIf Trim(frmCliente.txtRazonSocial.value) <> "" Then
			nCampo = afxNombres 
			sArgumento = Trim(frmCliente.txtRazonSocial.value)			
		ElseIf trim(frmCliente.txtGiro.value) <> "" Then		        
	                if isnumeric(frmCliente.txtGiro.value) and len(frmCliente.txtGiro.value) = 10 then
	                        
	                        frmCliente.action =  "AtencionClientes.asp?Accion=91"
	                        frmCliente.submit 
				            frmCliente.action = "" 
	                        exit sub
	                
	                else
	                    frmCliente.action = "ListaGiros.asp?Tipo=<%=afxListaGirosCodigo%>&Giro=" & trim(frmCliente.txtGiro.value)
				        frmCliente.submit 
				        frmCliente.action = ""	                
	                end if				
			Exit Sub						
		Elseif len(Trim(frmCliente.txtNombres.value))>=3 or len(trim(frmCliente.txtApellidoP.value)) >=3 or len(trim(frmCliente.txtApellidoM.value)) >=3  and sw <> 1 then		 
		    if len(trim(frmCliente.txtApellidoP.value)) < 3 and  len(trim(frmCliente.txtApellidoM.value))< 3 then		  
		      MsgBox "Debe ingresar mínimo 3 letras en alguno de los apellidos", ,"AFEX"
		      frmCliente.imgBuscar.disabled = false
		      
		      window.divProcesando.style.display = "none"
		         
		     else
		      nCampo = afxNombres 
			  sArgumento  = Trim(frmCliente.txtNombres.value)
			  sArgumento2 = Trim(frmCliente.txtApellidoP.value)
			  sArgumento3 = Trim(frmCliente.txtApellidoM.value)					  
			end if  	
		elseIf len(Trim(frmCliente.txtNombres.value))< 3 and len(trim(frmCliente.txtApellidoP.value)) < 3 and len(trim(frmCliente.txtApellidoM.value)) < 3 then
			MsgBox "Debe ingresar mínimo 3 letras en alguno de los campos para la busqueda.", ,"AFEX"
			frmCliente.imgBuscar.disabled = false
			
			window.divProcesando.style.display = "none"
				
		End If
							
 		Select Case Accion
		Case <%=afxAccionIngresarMG%>
			If nCampo <> 0 Then
				window.navigate "IngresarGiroMG.asp?Accion=" & Accion & "&Campo=" & nCampo & _
									 "&Argumento=" & sArgumento & _
									 "&Argumento2=" & sArgumento2 & _
									 "&Argumento3=" & sArgumento3 & _
									 "&ag=<%=Session("CodigoMGPago")%>"
			Else
				MsgBox "Debe buscar un cliente para ingresar un giro MG", , "AFEX En Línea"
			End If

		Case <%=afxAccionIngresarTX%>
			If nCampo <> 0 Then
				window.navigate "IngresarGiroMG.asp?Accion=" & Accion & "&Campo=" & nCampo & _
									 "&Argumento=" & sArgumento & _
									 "&Argumento2=" & sArgumento2 & _
									 "&Argumento3=" & sArgumento3 & _
									 "&ag=<%=Session("CodigoTXPago")%>"
			Else
				MsgBox "Debe buscar un cliente para ingresar un giro TX", , "AFEX En Línea"
			End If
		Case <%=afxAccionIngresarUT%>
			If nCampo <> 0 Then
				window.navigate "IngresarGiroMG.asp?Accion=" & Accion & "&Campo=" & nCampo & _
									 "&Argumento=" & sArgumento & _
									 "&Argumento2=" & sArgumento2 & _
									 "&Argumento3=" & sArgumento3 & _
									 "&ag=<%=Session("CodigoUTPago")%>"
			Else
				MsgBox "Debe buscar un cliente para ingresar un giro UT", , "AFEX En Línea"
			End If

		Case Else
			If nCampo <> 0 Then
				window.navigate "AtencionClientes.asp?Accion=" & Accion & "&Campo=" & nCampo & _
									 "&Argumento=" & sArgumento & _
									 "&Argumento2=" & sArgumento2 & _
									 "&Argumento3=" & sArgumento3
			End If
		End Select
				
	End Sub

	Sub CargarCliente()
		frmCliente.txtExchange.value = trim("<%=sAFEXchange%>")
		frmCliente.txtExpress.value = "<%=sAFEXpress%>"
		frmCliente.txtNombreCompleto.value = "<%=sNombreCompleto%>"
		frmCliente.txtApellidoM.value = "<%=sApellidoM%>"
		frmCliente.txtApellidoP.value = "<%=sApellidoP%>"
		frmCliente.txtNombres.value = "<%=sNombres%>"
		frmCliente.txtFNac.value = "<%=sFechaNac%>"
		frmCliente.txtNacionalidad.value = "<%=sNacionalidad%>"
		frmCliente.txtRazonSocial.value = "<%=sRazonSocial%>"
		frmCliente.txtDireccion.value = "<%=sDireccion%>"
		frmCliente.txtPaisFono.value = "<%=ObtenerDDI(1, sPais)%>" 
		frmCliente.txtAreaFono.value = "<%=ObtenerDDI(2, sCiudad)%>"
		frmCliente.txtFono.value = "<%=sFono%>"
		frmCliente.txtPaisFono2.value = "<%=ObtenerDDI(1, sPais)%>" 
		frmCliente.txtAreaFono2.value = "<%=ObtenerDDI(2, sCiudad)%>"
		frmCliente.txtFono2.value = "<%=sFono2%>"
		frmCliente.txtRut.value = "<%=sRut%>"
		frmCliente.txtPasaporte.value = "<%=sPasaporte%>"
		frmCliente.txtTarjeta.value = "<%=sTarjeta%>"
		frmCliente.txtNumeroCelular.value = "<%=sNumeroCelular%>" ' APPL-9009
		frmCliente.cbxSexo.value = "<%=sSexo%>" ' APPL-9009
		<% If sPasaporte <> "" Then %>
			    window.frmcliente.optpasaporte.checked=True
			window.frmcliente.optRut.checked=False
			window.frmCliente.txtRut.style.display = "none"		
			window.frmcliente.txtpasaporte.style.display=""			
			window.frmcliente.cbxPaisPasaporte.style.display=""
			lblPaisPasaporte.style.display=""
		<% End If %>
		
		<% If nTipoCliente <> 1 Then %>
			window.frmCliente.optEmpresa.checked = True
			trEmpresa.style.display = ""
			trPersona.style.display = "none"
			frmCliente.optPersona.checked = 0
		<% End If %>
	End Sub

	'miki SMC-80 MM 2016-04-25
    Sub CargarClienteActual()
		frmCliente.txtExchange.value = trim("<%=sAFEXchange%>")
		frmCliente.txtExpress.value = "<%=sAFEXpress%>"
		frmCliente.txtNombreCompleto.value = ""
		frmCliente.txtApellidoM.value = ""
		frmCliente.txtApellidoP.value = ""
		frmCliente.txtNombres.value = ""
		frmCliente.txtFNac.value = ""
		frmCliente.txtNacionalidad.value = ""
		frmCliente.txtRazonSocial.value = ""
		frmCliente.txtDireccion.value = ""
		frmCliente.txtPaisFono.value = "" 
		frmCliente.txtAreaFono.value = ""
		frmCliente.txtFono.value = ""
		frmCliente.txtPaisFono2.value = "" 
		frmCliente.txtAreaFono2.value = ""
		frmCliente.txtFono2.value = ""
		frmCliente.txtRut.value = "<%=sRut%>"
		frmCliente.txtPasaporte.value = "<%=sPasaporte%>"
		frmCliente.txtTarjeta.value = ""
		frmCliente.txtNumeroCelular.value = ""
		frmCliente.cbxSexo.value = ""
        frmCliente.cbxPais.value = ""
        frmCliente.cbxTarjetas.value = ""
		<% If sRut <> "" Then %>
			window.frmcliente.optpasaporte.checked=False
			window.frmcliente.optRut.checked=True
			window.frmCliente.txtRut.style.display = ""		
			window.frmcliente.txtpasaporte.style.display="none"
			window.frmcliente.cbxPaisPasaporte.style.display="none"
			lblPaisPasaporte.style.display="none"
		<% End If %>
			
		<% If nTipoCliente <> 1 Then %>
			window.frmCliente.optEmpresa.checked = True
			trEmpresa.style.display = ""
			trPersona.style.display = "none"
			frmCliente.optPersona.checked = 0
		<% End If %>
	End Sub
	'FIN miki SMC-80 MM 2016-04-25

	Sub CargarMenu()		
		frmCliente.objmenu.bgColor = document.bgColor 
		frmCliente.objmenu.stylesheet = "../Estilos/Cliente.css"
		sId = frmCliente.objmenu.addparent("Opciones")
		frmCliente.objMenu.addchild sId, "Giros Pendientes de Pago", "GPCliente", "Principal"
				
		' JFMG 05-09-2008 enlace a aplicación MoneyGram de AFEX
		'frmCliente.objMenu.addchild sId, "Agregar Pago MoneyGram", "IngresarMG", "Principal"
		frmCliente.objMenu.addchild sId, "MoneyGram", "MoneyGram", "Principal"
		' ********************************** FIN *****************************
		
		frmCliente.objMenu.addchild sId, "Agregar Pago Travelex", "IngresarTX", "Principal"
		'frmCliente.objMenu.addchild sId, "Agregar Pago Uniteller", "IngresarUT", "Principal" 'APPL-6080_MS_26-09-2014
		frmCliente.objMenu.addchild sId, "Limpiar", "Limpiar", "Principal"
		frmCliente.objMenu.addchild sId, "Compra", "Compra", "Principal"
		frmCliente.objMenu.addchild sId, "Venta", "Venta", "Principal"
		' Jonathan Miranda G. 27-02-2007
		frmCliente.objMenu.addchild sId, "Giros no Editados", "noEditados", "Principal"
		'---------------------- Fin -----------------------------
		frmCliente.objMenu.addchild sId, "", "", ""
	End Sub

	Sub CargarMenuInternacional()		
		frmCliente.objmenu.bgColor = document.bgColor 
		frmCliente.objmenu.stylesheet = "../Estilos/Cliente.css"
		sId = frmCliente.objmenu.addparent("Opciones")	
		frmCliente.objMenu.addchild sId, "Limpiar", "Limpiar", "Principal"
		frmCliente.objMenu.addchild sId, "", "", ""
		frmCliente.objMenu.addchild sId, "", "", ""
		frmCliente.objMenu.addchild sId, "", "", ""
		frmCliente.objMenu.addchild sId, "", "", ""
		frmCliente.objMenu.addchild sId, "", "", ""
	End Sub

	Sub CargarMenuCliente()
		frmCliente.objmenu.bgColor = document.bgColor 
		frmCliente.objmenu.stylesheet = "../Estilos/Cliente.css"
		sId = frmCliente.objmenu.addparent("Opciones")
		frmCliente.objMenu.addchild sId, "Giros Pendientes de Pago", "GirosPendientes", "Principal"
		frmCliente.objMenu.addchild sId, "Enviar un Giro", "EnviarGiro", "Principal"
		frmCliente.objMenu.addchild sId, "Enviar una Transferencia", "EnviarTransfer", "Principal"
		frmCliente.objMenu.addchild sId, "Solicitar un Cheque", "EnviarCheque", "Principal"
		
		' JFMG 05-09-2008 enlace a aplicación MoneyGram de AFEX
		'frmCliente.objMenu.addchild sId, "Agregar Pago MoneyGram", "IngresarMG", "Principal"
		frmCliente.objMenu.addchild sId, "MoneyGram", "MoneyGram", "Principal"
		' ********************************** FIN *****************************
				
		frmCliente.objMenu.addchild sId, "Agregar Pago Travelex", "IngresarTX", "Principal"
		'frmCliente.objMenu.addchild sId, "Agregar Pago Uniteller", "IngresarUT", "Principal" APPL-6080_MS_26-09-2014
		frmCliente.objMenu.addchild sId, "Movimientos del Cliente", "GirosCartolaLinea", "Principal"
		frmCliente.objMenu.addchild sId, "Cartola para Cliente", "GirosCartolaCliente", "Principal"
		frmCliente.objMenu.addchild sId, "Giros Enviados", "GirosEnviados", "Principal"
		frmCliente.objMenu.addchild sId, "Giros Recibidos", "GirosRecibidos", "Principal"
		frmCliente.objMenu.addchild sId, "Transferencias Enviadas", "TransferEnviadas", "Principal"
		frmCliente.objMenu.addchild sId, "Actualizar Datos Cliente", "Actualizar", "Principal"
		frmCliente.objMenu.addchild sId, "Compra", "Compra", "Principal"
		frmCliente.objMenu.addchild sId, "Venta", "Venta", "Principal"
		' Jonathan Miranda G. 27-02-2007
		frmCliente.objMenu.addchild sId, "Giros no Editados", "noEditados", "Principal"
		'---------------------- Fin -----------------------------

		frmCliente.objMenu.addchild sId, "Limpiar", "Limpiar", "Principal"
	End Sub

	Sub CargarMenuClienteInternacional()
		frmCliente.objmenu.bgColor = document.bgColor 
		frmCliente.objmenu.stylesheet = "../Estilos/Cliente.css"
		sId = frmCliente.objmenu.addparent("Opciones")
		frmCliente.objMenu.addchild sId, "Enviar un Giro", "EnviarGiro", "Principal"
		frmCliente.objMenu.addchild sId, "Limpiar", "Limpiar", "Principal"
		frmCliente.objMenu.addchild sId, "", "", ""
		frmCliente.objMenu.addchild sId, "", "", ""
		frmCliente.objMenu.addchild sId, "", "", ""
		frmCliente.objMenu.addchild sId, "", "", ""
		
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
			frmCliente.action = "AtencionClientes.asp?Accion=<%=afxAccionPais%>&Menu=<%=nMenu%>"
			frmCliente.submit 
			frmCliente.action = ""
	End Sub

	Sub cbxCiudad_onblur()
		Dim sComuna
		
		If frmCliente.cbxCiudad.value = "" Then Exit Sub
		If frmCliente.cbxCiudad.value = "<%=sCiudad%>" Then Exit Sub		
			HabilitarDireccion
			HabilitarId
			frmCliente.action = "AtencionClientes.asp?Accion=<%=afxAccionPais%>&Menu=<%=nMenu%>"
			frmCliente.submit 
			frmCliente.action = ""
	End Sub
	
	Sub optEmpresa_onClick()	
	dim sw
	sw=1
	   	trEmpresa.style.display = ""
		trPersona.style.display = "none"
		frmCliente.optPersona.checked = 0
		If Trim(frmCliente.txtExchange.value  & frmCliente.txtExpress.value) <> "" Then
			frmCliente.action = "AtencionClientes.asp"
			frmCliente.submit 
			frmCliente.action = "" 
		End If
		<% If nMenu = afxMenuNormal Then %> 
			LimpiarControles
		<% End If %>		
	End Sub
	
	Sub optPersona_onClick()		
		trEmpresa.style.display = "none"
		trPersona.style.display = ""
		frmCliente.optEmpresa.checked = 0
		If Trim(frmCliente.txtExchange.value  & frmCliente.txtExpress.value) <> "" Then
			frmCliente.action = "AtencionClientes.asp"
			frmCliente.submit 
			frmCliente.action = "" 
		End If
		<% If nMenu = afxMenuNormal Then %> 
			LimpiarControles
		<% End If %>
	End Sub
				
	Sub LimpiarControles
		frmCliente.txtPasaporte.value = ""
		frmCliente.txtRut.value = ""
		frmCliente.txtApellidoM.value = ""
		frmCliente.txtApellidoP.value = ""
		frmCliente.txtNombres.value = ""
		frmCliente.txtFNac.value ="" 
		frmCliente.txtRazonSocial.value = ""
		frmCliente.txtDireccion.value = ""
		frmCliente.txtPaisFono.value = ""
		frmCliente.txtAreaFono.value = ""
		frmCliente.txtFono.value = ""
		frmCliente.txtPaisFono2.value = ""
		frmCliente.txtAreaFono2.value = ""
		frmCliente.txtFono2.value = ""
		frmCliente.cbxPais.value = ""
		frmCliente.cbxCiudad.value = ""
		frmCliente.cbxComuna.value = ""	
		frmCliente.txtExchange.value = ""
		frmCliente.txtExpress.value = ""
		frmCliente.txtNombreCompleto.value = ""
		frmCliente.optpersona.value = ""
		frmCliente.txtNumeroCelular.value = "" ' APPL-9009
		'CargarMenu 
	End Sub

	Sub optGiros_onClick()
		frmCliente.optGiros.value = "on"
		frmCliente.optCambios.checked = false
		frmCliente.optCambios.value = ""
	End Sub
	Sub optCambios_onClick()
		frmCliente.optCambios.value = "on"
		frmCliente.optGiros.checked = false
		frmCliente.optGiros.value = ""
	End Sub	

	sub txtTarjeta_onBlur()
		if trim(frmCliente.txtTarjeta.value) <> "" then
			frmCliente.txtTarjeta.value = right("000000" & frmCliente.txtTarjeta.value,6)
		end if
	end sub

-->
</script>

<body style="display: <%=sDisplay%>">
<!-- #INCLUDE virtual="/Compartido/Encabezado.htm" -->
<!-- #INCLUDE virtual="/Compartido/Rutinas.htm" -->
<form id="frmCliente" method="post">
<input type="hidden" name="txtExchange">
<input type="hidden" name="txtExpress">
<input type="hidden" name="txtNombreCompleto">
<input type="hidden" name="cbxSexo" /> <!-- APPL-9009 -->
<table class="Borde" ID="tabPaso1" CELLSPACING="0" border="0" height="0px" style="position: relative; top: -18px; left: 2px; width: 480px;">
	<tr HEIGHT="15">
		<td colspan="2" class="titulo">Datos del cliente</td>
		<td align="right" class="titulo">
		<% If sAFEXpress <> "" Then %>
			<a align="right" onmouseout="window.status=''" onmouseover="window.status='<%=sAFEXpress%>'">Giros</a>
		<% End If %>
		<% If sAFEXchange <> "" Then %>
			<a align="right" onmouseout="window.status=''" onmouseover="window.status='<%=sAFEXchange%>'">&nbsp;&nbsp;Cambios</a>
		<% End If %>
		</td>
	</tr>
	<tr>
	<td width="1px"></td>
	<td>
	<table width="100%" border="0" style="HEIGHT: 0px; WIDTH: 300px" cellpadding="0" cellspacing="0">
		<tr HEIGHT="0">
			<!--<td></td>-->
			<td VALIGN="left" colspan="3"><br>
				<table border="0" cellpadding="0" cellspacing="0">
				
				<!-- ' JFMG 21-09-2010 busca las imagenes del cliente -->
				<tr>
					<td >
					<!-- APPL-6471 MS 20-07-2015 -->
					<!-- FIN APPL-6471 MS 20-07-2015 -->
					</td>
				</tr>
					<!-- ' FIN JFMG 21-09-2010 -->				
				<tr>
					<td>
						<a style="display: <%=sDisplay2%>">
						<input TYPE="radio" name="optRut" CHECKED style="border: 0">Rut</a>
						<input TYPE="radio" name="optPasaporte" style="border: 0"><%=sId%>
					</td>
					<td id="lblPaisPasaporte" style="display: none">Pais</td>
				</tr>
				<tr>
					<td>
						<input name="txtRut" style="width: 150px; text-align:right" OnKeyPress OnMouseOver="frmCliente.txtRut.value=FormatoRut(frmCliente.txtRut.value)">
						<input name="txtPasaporte" style="width: 150px; display: none">
					</td>
					<td>
						<select name="cbxPaisPasaporte" style="width: 150px; display: none">									
							<%
								CargarUbicacion 1, "", sPaisPass
							%>
						</select>

					</td>
				</tr>
				</table>
				
			</td>
		</tr>
		<tr>
			<td>Nº Tarjeta<br>
				<select name="cbxTarjetas">
					<% CargarPrefijoTarjeta sTarjetas%>
				</select>				
				<input name="txtTarjeta" style="width: 87px; text-align:right" onkeypress="IngresarTexto(1)">*
				</td>
            <td></td>
					<td colspan="2">Nacionalidad<br>
					<input DISABLED name="txtNacionalidad" SIZE="25" style="width: 100px" >
				</td>
		</tr>
		
		<tr id="trTipoCliente" style="display: " <%=sDisabled%>>
		<td colspan="2">
			<table id="tbTipoCliente"  cellpadding="0" cellspacing="0">
				<tr>
				<td WIDTH="1"></td>
				<td>
					<input TYPE="radio" id="optPersona" name="optPersona" style="border: 0" CHECKED>Persona			
				</td>
				<td>
					<input TYPE="radio" id="optEmpresa" name="optEmpresa" style="border: 0">Empresa
				</td>
				</tr>
			</table>
		</td>
		</tr>
		
		<tr>
			<% If Session("Categoria") = 4 Then %>		
				<td style="display: none">
			<% Else %>
				<td>
			<% End If %>
					Teléfono<br><input disabled name="txtPaisFono" style="width: 35px"><input disabled name="txtAreaFono" style="width: 35px"><input name="txtFono" style="width: 80px">*
				</td>
			<td></td>
			<% If Session("Categoria") = 4 Then %>		
				<td style="display: none">
			<% Else %>
				<td>
			<% End If %>
					Teléfono<br><input disabled name="txtPaisFono2" style="width: 35px"><input disabled name="txtAreaFono2" style="width: 35px"><input name="txtFono2" style="width: 80px">*
				</td>
		</tr>
		
		<!-- APPL-9009-->
		<tr>
		    <% If Session("Categoria") = 4 Then %>		
				<td style="display: none">
			<% Else %>
				<td>
			<% End If %>		
			    Celular <br />
				<span style="color: Gray;">(+56 9)</span>
				<input name="txtNumeroCelular" disabled style="width: 80px" />
			</td>
		</tr>
		<!-- FIN APPL-9009-->
		
		
		
		<tr ID="trEmpresa" HEIGHT="20" STYLE="DISPLAY: none"><td colspan="3">
		<table>
			<tr>
				<td WIDTH="1"></td>
				<td colspan="3">Razón Social<br>
				<input name="txtRazonSocial" id="txtRazonSocial" SIZE="40" style="width: 350px" onkeypress="IngresarTexto(2)" onblur="frmCliente.txtRazonSocial.value=MayMin(frmCliente.txtRazonSocial.value)">*
				</td>
			</tr>
		</table>
		</td></tr>
		<tr ID="trPersona" HEIGHT="20"><td colspan="3">
			<table>
			<tr>
				<td>
					<%
					if session("codigoagente") = "YA" then
						response.write request.servervariables("REMOTE_ADDR")
					end if
					%>

				</td>
					<td colspan="2">Nombres<br>
					<input NAME="txtNombres" SIZE="25" style="width: 350px" onkeypress="IngresarTexto(2)" onblur="frmCliente.txtNombres.value=MayMin(frmCliente.txtNombres.value)">* 
				</td>
			</tr>
			<tr>
			<td></td>
				<td>Apellido Paterno<br>
					<input name="txtApellidoP" SIZE="20" style="width: 170px" onkeypress="IngresarTexto(2)" >*
	                </td>
				<td>Apellido Materno<br>
					<input name="txtApellidoM" SIZE="20" style="width: 170px" onkeypress="IngresarTexto(2)" >*
	                </td>
			</tr>
			<tr>
			<td></td>
					<td colspan="2">Fecha de Nacimiento<br>
					<input DISABLED name="txtFNac" SIZE="20" style="width: 100px" onkeypress="IngresarTexto(2)" >
				</td>
			</tr>
			</table>
			
		</td></tr>
<% If Session("Categoria") = 4 Then %>		
		<tr style="display: none">
<% Else %>
		<tr>
<% End If %>
		<td colspan="3"><table>
			<tr>
				<td></td>
				<td COLSPAN="2">Dirección<br>
					<input disabled STYLE="WIDTH: 350px" SIZE="10" NAME="txtDireccion">
				</td>
			</tr>
			<tr>
				<td></td>
				<td>Pais<br>
					<select disabled name="cbxPais" style="width: 170px">
						<%	
							If nAccion = afxAccionBuscar Then							
						%>
								<option selected value="<%=sPais%>"><%=sNombrePais%>&nbsp;</option>
						<%	
							ElseIf nAccion <> afxAccionNada Then
								CargarUbicacion 1, "", sPais 	
							End If
						%>
					</select>
				</td>
				<td colspan="1">Ciudad<br>
					<select disabled name="cbxCiudad" style="width: 170px">
						<%	
							If nAccion = afxAccionBuscar Then							
						%>
								<option selected value="<%=sCiudad%>"><%=sNombreCiudad%>&nbsp;</option>
						<% 
							ElseIf nAccion <> afxAccionBuscar And nAccion <> afxAccionNada Then
								CargarCiudadesPais sPais, sCiudad 
							End If
						%>
					</select>
				</td>
			</tr>
			<tr>
				<td></td>
				<td>Comuna<br>
				<select disabled name="cbxComuna" style="width: 170px">
				 <script></script>
					<%	
						If nAccion = afxAccionBuscar Then							
					%>
							<option selected value="<%=sComuna%>"><%=sNombreComuna%>&nbsp;</option>
					<% 
						ElseIf nAccion <> afxAccionBuscar And nAccion <> afxAccionNada Then
							If sPais = "CL" Then						
								CargarComunaCiudad sCiudad, sComuna
							End If
						End If
					%>
				</select>
				</td>
				<td></td>
			</tr>
			</table>
		</td>
		</tr>
		<tr HEIGHT="2">
			<td></td>
		</tr>
	</table>
	</td>
	<td valign="top">
	<% If bCliente Then %>
			<% If Session("Categoria") = 4 Then %>
					<table border="0" width="100px" height="0px">
					<tr><td colspan="2"><object align="left" id="objMenu" style="HEIGHT: 130px; LEFT: 0px; POSITION: relative; TOP: -50px; WIDTH: 190px" type="text/x-scriptlet" width="190" VIEWASTEXT border="0"><param NAME="Scrollbar" VALUE="0"><param NAME="URL" VALUE="http:../Scriptlets/Menu.htm"></object></td></tr>
			<% Else %>
					<table border="0" width="100px" style="POSITION: relative; TOP: -90px">
					<tr height=50><td colspan="2"><object align="left" id="objMenu" style="HEIGHT: 240px; LEFT: 0px; POSITION: relative; TOP: 0px; WIDTH: 190px" type="text/x-scriptlet" width="190" VIEWASTEXT border="0"><param NAME="Scrollbar" VALUE="0"><param NAME="URL" VALUE="http:../Scriptlets/Menu.htm"></object></td></tr>
			<% End If %>
	<% Else %>
			<table border="0" width="100px" height="0px">
			<tr><td colspan="2"><object align="left" id="objMenu" style="HEIGHT: 130px; LEFT: 0px; POSITION: relative; TOP: 0px; WIDTH: 190px" type="text/x-scriptlet" width="190" VIEWASTEXT border="0"><param NAME="Scrollbar" VALUE="0"><param NAME="URL" VALUE="http:../Scriptlets/Menu.htm"></object></td></tr>	
	<% End If %>
	<tr>
		<td width="90" valign="left">
		   <input type="button" WIDTH="70" HEIGHT="20" value="Buscar"  id ="imgBuscar" ></input>			
		</td>
		<td width="110"></td>
	</tr>
	<tr style="display: ">
		<td>Giro<br>
			<input name="txtGiro" SIZE="10" style="width: 100px" >*
		</td>
	</tr>
	<tr height="100%"><td></td></tr>
	</table>
	</td>
	</tr>
</table>



    <div id="divProcesando" class="ContenedorModal">
        <div class="Modal">
        </div>
        <div class="InteriorModal300" style="text-align: center;">
           <img src="../Images/Procesando.gif" />
        </div>
    </div>



</form>
</body>
<!--#INCLUDE virtual="/Compartido/Rutinas.htm" -->
<script language="vbscript">

	Sub objMenu_OnScriptletEvent(strEventName, varEventData)
	   Select Case strEventName
	   
			Case "linkClick"
				If Right(varEventData, 10) = "EnviarGiro" Then			
					<% if session("Categoria") <> 4 then %>
					
					if trim(frmCliente.txtRut.value )= ""  and trim(frmCliente.txtPasaporte.value )="" then
								msgbox "El Cliente no puede enviar Giros sin Identificación, actualice sus datos.",,"AFEX"
								exit sub
					end if 	
					if trim(frmCliente.txtNombreCompleto.value )= "" then
						msgbox "No puede enviar un Giro sin nombre de cliente, actualice sus datos.",,"AFEX"
						exit sub
					end if
					<% If nTipoCliente = 1 Then %>						
						if trim(frmCliente.txtApellidoP.value )="" then
							msgbox "No puede enviar un Giro sin apellido paterno del cliente, actualice sus datos.",,"AFEX"
							exit sub
						end if
						If Trim(frmCliente.txtApellidoM.value) = Empty Then
							msgbox "Debe ingresar el apellido materno, actualice sus datos", ,"AFEX"
							Exit Sub
						End If					
					<% End If %>
					<% If nTipoCliente = 1 Then %>
						if trim(frmCliente.txtNacionalidad.value )="" then
							msgbox "No puede enviar un Giro sin nacionalidad del cliente, actualice sus datos.",,"AFEX"
							exit sub
						end if
					<% End If %>
					<% If nTipoCliente = 1 Then %>
						if trim(frmCliente.txtFNac.value )= "" then
							msgbox "Se requiere la fecha de nacimiento del cliente que envia, actualice sus datos.",,"AFEX"
							exit sub
						end if	
					<% End If %>
					if trim(frmCliente.txtDireccion.value )= "" then
							msgbox "No se puede enviar el giro sin la dirección del cliente, actualice sus datos.",,"AFEX"
							exit sub 
					end if
					if trim(frmCliente.txtFono.value  ) = "" or trim(ccur("0" & frmCliente.txtFono.value))=0 then
						msgbox "No se puede enviar el giro sin el teléfono del cliente, actualice sus datos.",,"AFEX"
						exit sub
					end if
					<%end if%>
							
					HabilitarDireccion 

					<% If Session("ModoPrueba") Then %>
							frmCliente.action = "EnviarGiro.asp?Accion=1&Cliente=" & frmCliente.txtExpress.value 
					<% Else %>
							frmCliente.action = "EnviarGiro.asp?Accion=1&Cliente=" & frmCliente.txtExpress.value & "&sPasap=" & trim(frmCliente.txtPasaporte.value ) & _
                                                "&sPaisPass=" & frmCliente.cbxPaisPasaporte.value & "&sRut=" & trim(frmCliente.txtRut.value)  'CUM-505 MS 02-02-2016
					<% End If %>
					frmCliente.submit 
					frmCliente.action = "" 

				ElseIf Right(varEventData, 14) = "EnviarTransfer" Then
					
					<% If NOT Session("UsuarioAutorizadoEnviarTransferencia") Then %>
						exit sub
					<% End If %>

					if frmCliente.txtExchange.value = empty then
						msgbox "Para enviar una Transferencia el cliente debe ser ingresado por el Depto. de Atención a Clientes.", , "AFEX"
						exit sub
					end if

					HabilitarDireccion 
					frmCliente.action = "EnviarTransfer.asp?Cliente=" & frmCliente.txtExchange.value
					frmCliente.submit 
					frmCliente.action = "" 
					
				ElseIf Right(varEventData, 12) = "EnviarCheque" Then
					if frmCliente.txtExchange.value = empty then
						msgbox "Para enviar un Cheque el cliente debe ser ingresado por el Depto. de Atención a Clientes.", , "AFEX"
						exit sub
					end if

					'exit sub
					HabilitarDireccion 
					frmCliente.action = "EnviarCheque.asp?Cliente=" & frmCliente.txtExchange.value
					frmCliente.submit 
					frmCliente.action = "" 

				
				ElseIf Right(varEventData, 9) = "MoneyGram" Then
					<%if Session("Categoria") = 3 then%> 
						window.open "<%= Session("URLMoneyGramAGENTES") %>" & _ 
								"?CuentaUsuario=" & trim("<%=RemplazaLetra(Session("NombreUsuarioOperador"))%>") & _		
								"&CodigoCliente=" & trim(frmCliente.txtExpress.value) & _
								"&ClienteAgente=" & "<%=Session("CodigoCliente")%>" & _
								"&CategoriaAgente=" & "<%=Session("Categoria")%>" & _
								"&CuentaSucursal=" & "<%=RemplazaLetra(Session("NombreUsuarioAgente"))%>" & _
								"&ContrasenaUsuario=" & "<%=RemplazaLetra(EncriptarCadena(trim(Session("ContrasenaOperador"))))%>" & _
								"&ContrasenaSucursal=" & "<%=RemplazaLetra(EncriptarCadena(trim(Session("ContrasenaAgente"))))%>"
						
					<%else%>
						window.open "<%= Session("URLMoneyGram") %>" & _   
								"?CuentaUsuario=" & trim("<%=RemplazaLetra(Session("NombreUsuarioOperador"))%>") & _		
								"&CodigoCliente=" & trim(frmCliente.txtExpress.value) & _
								"&ClienteAgente=" & "<%=Session("CodigoCliente")%>" & _
								"&CategoriaAgente=" & "<%=Session("Categoria")%>" & _
								"&CuentaSucursal=" & "<%=RemplazaLetra(Session("NombreUsuarioAgente"))%>" & _
								"&ContrasenaUsuario=" & "<%=RemplazaLetra(EncriptarCadena(trim(Session("ContrasenaOperador"))))%>" & _
								"&ContrasenaSucursal=" & "<%=RemplazaLetra(EncriptarCadena(trim(Session("ContrasenaAgente"))))%>"
					<%end if%>
					
					parent.close 					
									
				' *********************************** Fin *****************************************************

				ElseIf Right(varEventData, 10) = "IngresarTX" Then
					Buscar <%=afxAccionIngresarTX%>
					
				ElseIf Right(varEventData, 10) = "IngresarUT" Then
					Buscar <%=afxAccionIngresarUT%>
					
				ElseIf Right(varEventData, 15) = "GirosPendientes" Then
					If Trim(frmCliente.txtExpress.value) <> "" Then
						If <%=Session("IdCliente")%> = 0 Then
						End If
					End If

					frmCliente.action = "ListaGiros.asp?Tipo=<%=afxGirosPendientes%>&Cliente=" & frmCliente.txtExpress.value & _
							"&ApellidoCliente=" & Trim(Trim(frmCliente.txtApellidoP.value) & " " & Trim(frmCliente.txtApellidoM.value)) & _
							"&NombreCliente=" & Trim(frmCliente.txtNombres.value) & Trim(frmCliente.txtRazonSocial.value) & _
							"&AFEXpress=<%=Session("ATCAFEXpress")%>&AFEXchange=<%=Session("ATCAFEXchange")%>"
					frmCliente.submit 
					frmCliente.action = "" 


				ElseIf Right(varEventData, 17) = "GirosCartolaLinea" Then					
					frmCliente.action = "ListaGiros.asp?Tipo=<%=afxGirosCartola%>&Cliente=" & frmCliente.txtExpress.value 
					frmCliente.submit 
					frmCliente.action = "" 

				ElseIf Right(varEventData, 19) = "GirosCartolaCliente" Then					
					ImprimirCartolaCliente
				
				ElseIf Right(varEventData, 17) = "GirosCartolaLinea" Then					
					frmCliente.action = "ListaGiros.asp?Tipo=<%=afxGirosCartola%>&Cliente=" & frmCliente.txtExpress.value & "&AFEXpress=<%=Session("ATCAFEXpress")%>&AFEXchange=<%=Session("ATCAFEXchange")%>"
					frmCliente.submit 
					frmCliente.action = "" 

				ElseIf Right(varEventData, 19) = "GirosCartolaCliente" Then					
					ImprimirCartolaCliente

				ElseIf Right(varEventData, 13) = "GirosEnviados" Then
					If Trim(frmCliente.txtExpress.value) <> "" Then
						frmCliente.action = "ListaGiros.asp?Tipo=<%=afxGirosEnviados%>&Cliente=" & frmCliente.txtExpress.value & "&Agente=<%=Session("CodigoAgente")%>" & _
								"&ApellidoCliente=" & Trim(Trim(frmCliente.txtApellidoP.value) & " " & Trim(frmCliente.txtApellidoM.value)) & _
								"&NombreCliente=" & Trim(frmCliente.txtNombres.value) & Trim(frmCliente.txtRazonSocial.value) & _
								"&AFEXpress=<%=Session("ATCAFEXpress")%>&AFEXchange=<%=Session("ATCAFEXchange")%>"
					Else
						frmCliente.action = "ListaGiros.asp?Tipo=<%=afxGirosEnviados%>&Cliente=00000&Agente=<%=Session("CodigoAgente")%>" & _
								"&ApellidoCliente=" & Trim(Trim(frmCliente.txtApellidoP.value) & " " & Trim(frmCliente.txtApellidoM.value)) & _
								"&NombreCliente=" & Trim(frmCliente.txtNombres.value) & Trim(frmCliente.txtRazonSocial.value) & _
								"&AFEXpress=<%=Session("ATCAFEXpress")%>&AFEXchange=<%=Session("ATCAFEXchange")%>"					
					End If
					frmCliente.submit 
					frmCliente.action = "" 

				ElseIf Right(varEventData, 14) = "GirosRecibidos" Then
					frmCliente.action = "ListaGiros.asp?Tipo=<%=afxGirosRecibidos%>&Cliente=" & frmCliente.txtExpress.value  & "&Pagador=<%=Session("CodigoAgente")%>" & _
							"&ApellidoCliente=" & Trim(Trim(frmCliente.txtApellidoP.value) & " " & Trim(frmCliente.txtApellidoM.value)) & _
							"&NombreCliente=" & Trim(frmCliente.txtNombres.value) & Trim(frmCliente.txtRazonSocial.value) & _
							"&AFEXpress=<%=Session("ATCAFEXpress")%>&AFEXchange=<%=Session("ATCAFEXchange")%>"
					frmCliente.submit 
					frmCliente.action = "" 
				
				' Jonathan Miranda G. 27-02-2007
				ElseIf Right(varEventData, 10) = "noEditados" Then
					if Trim(frmCliente.txtNombres.value) = empty then
						exit sub
					end if
					frmCliente.action =  "ListaGiros.asp?Tipo=98&Pagador=<%=Session("CodigoAgente")%>" & _
							"&ApellidoCliente=" & Trim(Trim(frmCliente.txtApellidoP.value) & " " & Trim(frmCliente.txtApellidoM.value)) & _
							"&NombreCliente=" & Trim(frmCliente.txtNombres.value) & Trim(frmCliente.txtRazonSocial.value)
					frmCliente.submit 
					frmCliente.action = ""
				'-------------------------------------- Fin ------------------------------

				ElseIf Right(varEventData, 16) = "TransferEnviadas" Then
					frmCliente.action = "ListaTransfer.asp?Titulo=Transferencia Enviadas&Tipo=9&Cliente=" & frmCliente.txtExchange.value & "&Desde=" & cDate(Date()-30) & "&Hasta=" & Date() & "&TipoLlamada=1"
					frmCliente.submit 
					frmCliente.action = "" 

				ElseIf Right(varEventData, 10) = "Actualizar" Then
					' Jonathan Miranda G. 05-06-2007
					if frmCliente.txtExpress.value = "" then
						msgbox "Este cliente no puede ser actualizado desde esta página, solicitelo a Atención Clientes.",,"AFEX"
						exit sub
					end if
					'------------------- Fin ------------------------
					
					HabilitarDireccion 
					frmCliente.action = "ActualizacionCliente.asp?AFEXchange=<%=sAFEXchange%>&AFEXpress=<%=sAFEXpress%>"
					frmCliente.submit 
					frmCliente.action = "" 
					
				ElseIf Right(varEventData, 9) = "GPCliente" Then	
					IF trim(frmCliente.txtApellidoP.value )<>"" and trim(frmCliente.txtNombres.value  )<>"" Then
						If trim(len(frmCliente.txtApellidoP.value ))>=3 and trim(len(frmCliente.txtNombres.value ))>=3 Then						
							frmCliente.action = "ListaGiros.asp?Tipo=<%=afxGirosPendientes%>&Cliente=" & frmCliente.txtExpress.value & _
									"&ApellidoCliente=" & Trim(Trim(frmCliente.txtApellidoP.value) & " " & Trim(frmCliente.txtApellidoM.value)) & _
									"&NombreCliente=" & Trim(frmCliente.txtNombres.value) & Trim(frmCliente.txtRazonSocial.value) & _
									"&Telefono=" & Trim(frmCliente.txtFono.value) & _
									"&Telefono2=" & Trim(frmCliente.txtFono2.value) & _
									"&AFEXpress=<%=Session("ATCAFEXpress")%>&AFEXchange=<%=Session("ATCAFEXchange")%>" & _
									"&Rut=" & Trim(frmCliente.txtRut.value) & "&Pasaporte=" & Trim(frmCliente.txtPasaporte.value)
							frmCliente.submit 
							frmCliente.action = "" 
						Else
							MsgBox "Debe ingresar mínimo 3 letras para realizar la búsqueda", ,"AFEX"
						End If                        
					End IF
						
				ElseIf Right(varEventData, 7) = "Limpiar" Then
					frmCliente.action = "AtencionClientes.asp"
					frmCliente.submit 
					frmCliente.action = "" 
					
				ElseIf Right(varEventData, 6) = "Compra" Then
					frmCliente.action = "Prueba3.asp?to=1&nc=<%=sNombreCiudad%>&dc=<%=sDireccion%>&pf=<%=sPaisFono%>" & _
													"&af=<%=sAreaFono%>"
					frmCliente.submit 
					frmCliente.action = "" 
					
				ElseIf Right(varEventData, 5) = "Venta" Then
					frmCliente.action = "Prueba3.asp?to=2&nc=<%=sNombreCiudad%>&dc=<%=sDireccion%>&pf=<%=sPaisFono%>" & _
													"&af=<%=sAreaFono%>"
					frmCliente.submit 
					frmCliente.action = "" 

				Else
					window.open varEventData, "Principal"
				
				End If
				
		End Select

	End Sub
	
	Sub ImprimirCartolaCliente()			
		window.open  "ImprimirCartolaCliente.asp?init=actx" & _
						"&ca=" & "<%=Session("CodigoAgente")%>" & _
						"&fd=" & "19990101" & _
						"&fh=" & "<%=FormatoFechaSQL(Date)%>" & _
						"&cl=" & frmCliente.txtExpress.value & _
						"&user0=giros&password0=giros", _
						"", "dialogHeight= 250pxl; dialogWidth= 250pxl; " & _
					    "dialogTop= 0; dialogLeft= 0; " & _
						"status=no; scrollbars=yes"		
	End Sub

</script>
</html>

