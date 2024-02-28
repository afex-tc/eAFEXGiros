<%@ LANGUAGE = VBScript %>
<%
	'Se asegura que la página no se almacene en la memoria cache
	Response.Expires = 0
	If Session("CodigoCliente") = "" Then
		response.Redirect "../Compartido/TimeOut.htm"
		response.end
	End If
%>
<!--#INCLUDE virtual="/agente/Constantes.asp" -->
<!--#INCLUDE virtual="/Compartido/Rutinas.asp" -->
<!--#INCLUDE virtual="/Compartido/Errores.asp" -->
<!--#INCLUDE virtual="/Compartido/RutinasEncriptar.asp" -->

<%	
	Dim sCodigo
	Dim nTipoCliente
	Dim sAFEXchange, sAFEXpress
	Dim nNegocio
	
	On Error	Resume Next
	nNegocio = cInt(0 & Request.Form("cbxNegocio"))
	
	If Err.number <> 0 Then
		response.Redirect "../Compartido/Error.asp?Titulo=Error en Agregar Cliente&Number=" & Err.number   & "&Source=" & Err.Source & "&Description=" & Err.description
	End If

	ValidarIdentificacion

	If Request.Form("optPersona") = "on" Then
		nTipoCliente = 1
	Else
		nTipoCliente = 2
	End If
	If Err.number <> 0 Then
		response.Redirect "../Compartido/Error.asp?Titulo=Error en Agregar Cliente 1&Number=" & Err.number   & "&Source=" & Err.Source & "&Description=" & Err.description
	End If
	
	Select Case nNegocio
		Case afxGiros
			sAFEXpress = AgregarClienteXP 			
			If Err.number <> 0 Then
				response.Redirect "../Compartido/Error.asp?Titulo=Error en Agregar Cliente 2&Number=" & Err.number   & "&Source=" & Err.Source & "&Description=" & Err.description & ", " & nNegocio
			End If
			
			' JFMG 20-11-2012 por VIGO			
			If Request("Refrencia") <> "" Then
			    dim referencia 
			    referencia = Split(Request("Refrencia"),"|")
			   			    
			    If DesEncriptarCadena(referencia(1)) = "eAFEXNET" Then
			        		        	        
			        dim parametros
	                parametros = EncriptarCadena(Session("NombreUsuarioOperador"))
	                parametros = parametros & "|" & EncriptarCadena(Session("CodigoAgente"))
	                parametros = parametros & "|" & EncriptarCadena(Session("CodigoCliente"))
	                parametros = parametros & "|" & referencia(0)
	                parametros = parametros & "|||" & EncriptarCadena(sAFEXpress)
	                
	                parametros = replace(parametros, "Ñ", "%c3%91")
		            parametros = replace(parametros, "Ã‘", "%c3%91")
		            parametros = replace(parametros, "ñ", "%c3%b1")
	                
	                'response.Write parametros & " ** " & referencia(0)
	                'response.end
	                   
	                response.Redirect(Session("URLeAFEXNet") & "?Referencia=" & parametros)			    
			    End If		    
	            
			Else
			' FIN JFMG 20-11-2012
			
			    Response.Redirect "AtencionClientes.asp?Accion=1&Campo=" & afxCampoCodigoExpress & "&Argumento=" & sAFEXpress 
			
			' JFMG 20-11-2012 por VIGO
		    End if
			' FIN JFMG 20-11-2012
				
		Case afxCambios
			sAFEXchange = AgregarClienteXC(nTipoCliente)
			If Err.number <> 0 Then
				response.Redirect "../Compartido/Error.asp?Titulo=Error en Agregar Cliente 3&Number=" & Err.number   & "&Source=" & Err.Source & "&Description=" & Err.description & ", " & nNegocio
			End If
			Response.Redirect "AtencionClientes.asp?Accion=1&Campo=" & afxCampoCodigoExchange & "&Argumento=" & sAFEXchange
				
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
			if rs("express") <> "" then
				Set rs = Nothing
				ValidarIdentificacion = False
				response.Redirect  "../Compartido/Error.asp?Titulo=Error en Agregar Cliente" & "&Description=Imposible agregar el nuevo cliente. La identificación ya existe"
				response.End 
				Exit Function
			end if
		End If
		
		Set rs = Nothing
	End Function
		
%>
