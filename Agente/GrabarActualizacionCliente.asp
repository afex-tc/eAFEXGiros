<%@ LANGUAGE = VBScript %>
<%
	'Se asegura que la página no se almacene en la memoria cache
	Response.Expires = 0
	If Session("CodigoCliente") = "" Then
		response.Redirect "../Compartido/TimeOut.htm"
		response.end
	End If
%>
<!--#INCLUDE virtual="/Compartido/Constantes.asp" -->
<!--#INCLUDE virtual="/Compartido/Rutinas.asp" -->
<!--#INCLUDE virtual="/Compartido/Errores.asp" -->
<%	
    
	Dim nNegocio
	Dim sCodigo
	Dim nTipoCliente
	Dim sAFEXchange, sAFEXpress
	dim sTarjeta
	dim sFechaNacimiento, sSexo
	 				
	 sTarjeta = Request.Form("txtTarjeta")
	 if trim(sTarjeta) <> "" then sTarjeta = trim(Request.Form("cbxTarjetas")) & right(trim(Request.Form("txtTarjeta")),6)
	 	
		
	'On Error	Resume Next

	sAFEXpress = Trim(Request.Form("txtExpress"))
	sAFEXchange = Trim(Request.Form("txtExchange"))
	nTipoCliente = cInt(0 & Trim(Request("TipoCliente")))
		
	If nTipoCliente = 2 then
		sFechaNacimiento = "01/01/1900"
		sSexo = 0 
	else
		sFechaNacimiento= Request.Form("txtfechanacimiento")
		sSexo = request.Form("cbxSexo")
	end If
	
	If sAFEXpress <> "" Then		
		If Not ValidarIdentificacion Then		
		Else		
			ActualizarClienteXP 
		End If	
	End If
	
	If Trim(Request.Form("txtRut")) <> "" Or Trim(Request.Form("txtPasaporte")) <> "" Then
		Session("IdCliente") = 1
	Else
		Session("IdCliente") = 0
	End If
	
	Select Case Request("Tipo")
	Case "DetalleGiro"
		response.Redirect  "DetalleGiro.asp?Codigo=" & Request("Giro") & "&Cliente=" & request("Cliente") & "&AFEXpress=" & Request("AFEXpress") & "&AFEXchange=" & Request("AFEXchange") & "&Accion=" & Request("Accion")
			
	Case Else
		If sAFEXpress <> "" Then
			Response.Redirect "AtencionClientes.asp?Accion=1&Campo=6&Argumento=" & sAFEXpress
		End If
		If sAFEXchange <> "" Then
			Response.Redirect "AtencionClientes.asp?Accion=1&Campo=5&Argumento=" & sAFEXchange
		End If		
	End Select


	'Funciones y Procedimientos
	
	Function ValidarIdentificacion()
		Dim rs, sNuevo, nCampo, sArgumento
		
		ValidarIdentificacion = True
		If Request.Form("optRut") = "on" Then
			nCampo = afxCampoRut
			sArgumento = Request.Form("txtRut")
		Else
			nCampo = afxCampoPasaporte
			sArgumento = Request.Form("txtPasaporte")
		End If	
		Set rs = BuscarCliente(nCampo, sArgumento, "", "")
			
		'Si encuentra un cliente con la misma identificacion
		'va a Asociar Clientes
		If Not rs.EOF Then
			Do Until rs.EOF
				If Trim(rs("express")) <> Trim(sAFEXpress) Then
					Exit Do
				End If 
				rs.movenext
			Loop
			If Not rs.EOF Then
				sNuevo = rs("express")
				Set rs = Nothing
				ValidarIdentificacion = False
				response.Redirect  "AsociarCliente.asp?Nuevo=" & sNuevo & "&Eliminar=" & sAFEXpress	& "&Giro=" & Request("Giro") & "&AFEXchange=" & Request("AFEXchange") & "&Tipo=" & Request("Tipo") & "&Accion=" & Request("Accion")
				response.End 
				Exit Function
			End If
		End If
		
		Set rs = Nothing
	End Function
		
	Sub ActualizarClienteXP
	    
	    ' APPL-9009
	    Dim sSQLCliente
	    Dim rsCliente
		'Dim afxClienteXP, bOK
								 
		'Set afxClienteXP = Server.CreateObject("AfexClienteXP.Cliente")
		
		'bOK = afxClienteXP.Actualizar(Session("afxCnxAFEXpress"), sAFEXpress,  _
		'						 Request.Form("txtRut"), Request.Form("txtPasaporte"), Request.Form("cbxPaisPasaporte"), _
		'						 Request.Form("txtNombres"), _
		'						 Trim(Trim(Request.Form("txtApellidoP")) & " " & Trim(Request.Form("txtApellidoM"))), _
		'						 sfechanacimiento, request.Form("txtDireccion"), _
		'						 request.Form("cbxComuna"), request.Form("cbxCiudad"), _
		'						 request.Form("cbxPais"), CInt(0 & request.Form("txtPaisFono")), _
		'						 CInt(0 & request.Form("txtAreaFono")), CCur(0 & request.Form("txtFono")), _
	 	'						 CInt(0 & request.Form("txtAreaFono2")), CCur(0 & request.Form("txtFono2")), _
	 	'						 sTarjeta, sSexo, request.Form("cbxNacionalidad"), _
	 	'						 request.Form("txtCorreoElectronico"))
	 	Dim sRut, sPasaporte, sPaisPasaporte
	 	
	 	If Request.Form("txtRut") <> "" Then
            sRut = ValorRut(Request.Form("txtRut"))
            sPasaporte = ""
            sPaisPasaporte = ""
        Else
            sRut = ""
            sPasaporte = Request.Form("txtPasaporte")
            sPaisPasaporte = Request.Form("cbxPaisPasaporte")        
        End If
        
        If Request.Form("txtFechaNacimiento") <> "" then
            sFechaNacimiento = FormatoFechaSQL(Request.Form("txtFechaNacimiento"))
        Else
            sFechaNacimiento = ""
        End If
        
        On Error Resume Next
	 	sSQLCliente = "exec ActualizarCliente " & EvaluarStr(sAFEXpress) & ", " & EvaluarStr(Request.Form("txtNombres")) & ", " & _
                                 EvaluarStr(Trim(Trim(Request.Form("txtApellidoP")) & " " & Trim(Request.Form("txtApellidoM")))) & ", " & EvaluarStr(sFechaNacimiento) & ", " & _
                                 EvaluarStr(request.Form("txtDireccion")) & ", " & EvaluarStr(request.Form("cbxComuna")) & ", " & _
                                 EvaluarStr(request.Form("cbxCiudad")) & ", " & EvaluarStr(request.Form("cbxPais")) & ", " & _
                                 EvaluarVar(cint("0" & request.Form("txtPaisFono")), "0") & ", " & EvaluarVar(cint("0" & request.Form("txtAreaFono")), "0") & ", " & _
                                 EvaluarVar(CCur("0" & request.Form("txtFono")), "0") & ", " & EvaluarStr(sRut) & ", " & _
                                 EvaluarStr(sPasaporte) & ", " & EvaluarStr(sPaisPasaporte) & ", " & _
                                 EvaluarVar(CInt("0" & request.Form("txtAreaFono2")), "0") & ", " & EvaluarVar(CCur("0" & request.Form("txtFono2")), "0") & ", " & _
                                 EvaluarStr(sTarjeta) & ", " & EvaluarVar(sSexo, "0") & ", " & EvaluarStr(request.Form("cbxNacionalidad")) & ", " & _
                                 EvaluarStr(request.Form("txtCorreoElectronico")) & ", NULL, NULL, NULL, NULL, NULL, NULL, " & _
                                 EvaluarVar(CCur("0" & request.Form("txtNumeroCelular")), "0")
        
		Set rsCliente = EjecutarSqlCliente(Session("afxCnxAfexpress"), sSQLCliente)
		If Err.number <> 0 Then
		'	Set afxClienteXP = Nothing
			MostrarErrorMS "Error en Actualizar Cliente Giros"
		End If
		'If afxClienteXP.ErrNumber <> 0 Then
		'	MostrarErrorAFEX afxClienteXP, "Error en Actualizar Cliente Giros"
		'End If
		'Set afxClienteXP = Nothing
        ' FIN APPL-9009
	End Sub	
%>
