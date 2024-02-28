<%@ LANGUAGE = VBScript %>
<%
	'option explicit	
	'Se asegura que la página no se almacene en la memoria cache
	Response.Expires = 0
	If Session("CodigoCliente") = "" Then
		response.Redirect "../Compartido/TimeOut.htm"
		response.end
	End If
%>

<%	
	Dim afx, sCliente, Giro
	Dim sApellidos, sNombres, sRut, sPasaporte
	Dim sDireccion, sPais, sCiudad, sComuna, sPaisPass, sNombreCompleto
	Dim sNombrePais, sNombreCiudad, sNombreComuna, sNombrePaisPass
	Dim sPaisFono, sAreaFono, sFono
	Dim nAccion, sAFEXpress, nDestinoBoleta
	
	On Error	Resume Next
	
	sAFEXpress = Session("AFEXpress")
	sCliente = Session("CodigoCliente")
	CargarCliente
	AgregarGiro
				
	Set afx = Nothing
	Response.Redirect "Resultado.asp"
					
	Sub CargarCliente()
		Dim nCampo, rs
				
		Set rs = BuscarCliente(6, sAFEXpress, "", "")
		
		If rs.EOF Then
			Set rs = Nothing
			Exit Sub
		End If
				
		If Not rs.EOF Then
			
			nTipoCliente = cInt(0 & rs("tipo"))
			
			If nTipoCliente = 1 Then
				sApellidos = Trim(Trim(MayMin(EvaluarVar(rs("paterno"), ""))) & " " & Trim(MayMin(EvaluarVar(rs("materno"), ""))))
				sNombres = MayMin(EvaluarVar(rs("nombre"), ""))
			Else
				sNombres = MayMin(EvaluarVar(rs("nombre_completo"), ""))
			End If	
			
			sNombreCompleto = MayMin(EvaluarVar(rs("nombre_completo"), ""))
			sRut = FormatoRut(EvaluarVar(rs("rut"), ""))
			sPasaporte = EvaluarVar(rs("pasaporte"), "")
			sPaisPass = EvaluarVar(rs("codigo_paispas"), "")
			sDireccion = MayMin(EvaluarVar(rs("direccion"), ""))
			sPais = EvaluarVar(rs("codigo_pais"), "")
			sCiudad = EvaluarVar(rs("codigo_ciudad"), "")
			sComuna = EvaluarVar(rs("codigo_comuna"), "")
			sPaisFono = cInt(0 & EvaluarVar(rs("ddi_pais"), ""))
			sAreaFono = cInt(0 & EvaluarVar(rs("ddi_area"), ""))
			sFono = cCur(0 & EvaluarVar(rs("telefono"), ""))
			sNombrePais = MayMin(EvaluarVar(rs("pais"), ""))
			sNombreCiudad = MayMin(EvaluarVar(rs("ciudad"), ""))
			sNombreComuna = MayMin(EvaluarVar(rs("comuna"), ""))
			sNombrePaisPass = MayMin(EvaluarVar(rs("paispas"), ""))
			If Err.number <> 0 Then
				MostrarErrorMS ""
			End If
		End If
		Set rs = Nothing
	End Sub
	
'Métodos	
	Sub AgregarGiro()
'		Response.Redirect "../compartido/error.asp?description=" & _
'			Session("afxCnxAFEXpress") & ", " & "AM" & ", " &  _
'			Session("CodigoMatriz") & ", " &  cCur(cDbl(Request.Form("txtMonto"))) & ", " &  _
'			cCur(cDbl(Request.Form("txtTarifa"))) & ", " &  0 & ", " &  1 & ", " &  0 & ", " &  "USD" & ", " &  _
'			Request.Form("cbxMoneda") & ", " &  Request.Form("txtMensaje") & ", " &  _
'			"" & ", " &  "" & ", " &  "" & ", " &  _
'			"" & ", " &  Request.Form("txtNombre") & ", " &  Request.Form("txtApellido") & ", " &  _
'			Request.Form("txtDireccion") & ", " &  Request.Form("cbxCiudad") & ", " &  "" & ", " &  Request.Form("cbxPais") & ", " &  _
'			cInt(0 & Request.Form("txtPaisFono")) & ", " &  cInt(0 & Request.Form("txtAreaFono")) & ", " &  cCur(0 & Request.Form("txtFono")) & ", " &  _
'			sRut & ", " &  sPasaporte & ", " &  sPaisPass & ", " &  _
'			sNombres & ", " &  sApellidos & ", " &  sDireccion & ", " &  _
'			sCiudad & ", " &  sComuna & ", " &  sPais & ", " &  _
'			cInt(0 & sPaisFono) & ", " & cInt(0 & sAreaFono) & ", " &  cCur(0 & sFono) & ", " &  _
'			"AM" & ", , " &  sAFEXpress
		
		Set afx = Server.CreateObject("AfexWebXP.Web")
		If Err.number <> 0 Then
			Set afx = Nothing
			MostrarErrorMS "Grabar Envio Giro 1"
		End If
		'MostrarErrorMS Request.Form("optEnviarBoleta") & ", " & Request.Form("optGuardarBoleta")
		If Request.Form("optEnviarBoleta") = "on" Then
			nDestinoBoleta = 1
		End If
		If Request.Form("optGuardarBoleta") = "on" Then
			nDestinoBoleta = 2
		End If
		Giro = afx.EnviarGiro(Session("afxCnxAFEXpress"), _
			Session("CodigoAgente"), Session("CodigoMatriz"),  _
			cCur(cDbl(0 & Request.Form("txtMonto"))), cCur(cDbl(0 & Request.Form("txtTarifa"))),  _
			0,  1,  0,  Request.Form("cbxMoneda"), Request.Form("cbxMoneda"),  _
			Request.Form("txtMensaje"), "",  "",  "", "", _
			Request.Form("txtNombre"),  Request.Form("txtApellido"),  _
			Request.Form("txtDireccion"),  Request.Form("cbxCiudad"),  "",  Request.Form("cbxPais"),  _
			cInt(0 & Request.Form("txtPaisFono")),  cInt(0 & Request.Form("txtAreaFono")), cCur(0 & Request.Form("txtFono")),  _
			sRut, sPasaporte, sPaisPass,  _
			sNombres, sApellidos, sDireccion, sCiudad, sComuna,  sPais,  _
			cInt(0 & sPaisFono), cInt(0 & sAreaFono),  cCur(0 & sFono),  _
			Session("codigoAgente"),,sAFEXpress,,,,, True, Session("afxCnxCorporativa"), Session("CodigoCliente"), nDestinoBoleta)
		If Err.number <> 0 Then
			Set afx = Nothing
			MostrarErrorMS "Grabar Envio Giro 2"
		End If						
		If afx.ErrNumber <> 0 Then
			MostrarErrorAFEX afx, "Grabar Envio Giro 3"
		End If
		Set afx = Nothing
	End Sub
 							 	
%>
<!--#INCLUDE virtual="/Compartido/Rutinas.asp" -->
<!--#INCLUDE virtual="/Compartido/Errores.asp" -->
