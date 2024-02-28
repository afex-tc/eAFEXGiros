<!--#INCLUDE virtual="/Compartido/Rutinas.asp" -->
<!--#INCLUDE virtual="/Compartido/Errores.asp" -->
<%
	'Se asegura que la página no se almacene en la memoria cache
	Response.Expires = 0
	If Session("CodigoCliente") = "" Then
		response.Redirect "../Compartido/TimeOut.htm"
		response.end
	End If
%>
<%
    ON Error Resume Next

	Dim objEMail
	Dim hora, fecha
	Dim sSQL2,sSQL3,sSQL4,rs4,sw,rs3
	Dim rs
	
	hora= Time()
	fecha= Date()
	hora1= replace(time,":","")
	
	Cod_giro = request.Form("Cod_Giro") 
	Monto= request.Form ("Monto")
	Detalle = request.Form ("Detalle")
	Motivo= request.form("txtAnular")
						
            ' Tecnova rperez 30-08-2016 - Sprint 2: Se comenta bloque de código por requerimiento INTERNO 5810
			'sSQL2 = "exec SolicitudAnulacion " & EvaluarStr(cod_giro) & ", " & EvaluarStr(Motivo) & ", " & _
			'			        EvaluarStr(Session("CodigoAgente")) & ", " & EvaluarStr(Session("NombreUsuarioOperador"))			
			'set rs= ejecutarsqlcliente(session("afxcnxAfexpress"),sSQL2)
			'If Err.number <> 0 Then	
		    '	set rs = nothing 			
			'	MostrarErrorMS "Enviar Anulación "
			'End If			
			
			'' verifica si el giro se actualizó
			'sSQL2 = "select estado_giro from giro with(nolock) where codigo_giro=" & evaluarstr(cod_giro)			
			'set rs = ejecutarsqlcliente(session("afxCnxAFEXpress"), sSQL2)
			'If Err.number <> 0 Then
			'	Set rs = Nothing
			'	MostrarErrorMS "Enviar Anulación - Consulta "
			'End If
			'if rs.eof then
			'	Set rs = Nothing
			'	MostrarErrorMS "Enviar Anulación - Validación "
			'end if
			'if not rs.eof then
			'	if rs("estado_giro") <> 7 and rs("estado_giro") <> 9 then
			'		MostrarErrorMS "El giro no se actualizó. Comuníquese con Informática para informarles del problema."
			'	end if
			'end if
	        
	    
			'Set objEMail = Server.CreateObject("CDONTS.NewMail")
			'objEMail.From = Request.Form("txtNombre") & " <" & Request.Form("txtEmail") & ">"			
			'objEMail.To = "jonathan.miranda@afex.cl" '"reclamos@afex.cl"		
			'objEMail.cc = "jonathan.miranda@afex.cl" '"transmisiones@afex.cl;fernando.barrera@afex.cl"
			'objEMail.Subject = "Anulación de Giro. Código:" & Cod_Giro &" (" & fecha &" " & hora &")"
			'objEMail.Body = Motivo & Chr(13) & "NOTA: El Giro " & Cod_Giro &" quedo Automaticamente en estado de Reclamo esta es una peticion de " & Session("NombreUsuarioOperador")
			'objEMail.Send
			'Set objEMail = Nothing
			
			' envía mail
			Dim rsListaCorreo
			sSQL = "exec Mail_Listar_Destinatarios 5, 15, " & Session("AmbienteServidorCorreo")
			set rsListaCorreo = ejecutarsqlcliente(session("afxCnxServidorCorreo"), sSQL)
			If Err.number <> 0 Then				
				MostrarErrorMS "Mail Anulación - " & err.Description 
			End If
			If not rsListaCorreo is nothing then
			    If rsListaCorreo.EOF then
			        MostrarErrorMS "Mail Anulación - Si Destinatarios"
			    End If
			    Do While not rsListaCorreo.EOF
			        sSQL = "exec Mail_Anulacion_Giro '" & Session("PerfilServidorCorreo") & "'," & evaluarstr(cod_giro) & ", " & _
			                                                    evaluarstr(rsListaCorreo("email")) & ", " & evaluarstr(rsListaCorreo("nombre")) & ", " & _
			                                                    evaluarstr(Request.Form("txtNombre")) & ", " & evaluarstr(Motivo) & ", " & _
		                                                        evaluarstr(Session("NombreOperador"))
			        set rs = ejecutarsqlcliente(session("afxCnxServidorCorreo"), sSQL)
    			
			        rsListaCorreo.MoveNext
			    Loop
			    ' envía email al usuario que anulo
			    if Request.Form("txtEmail") <> "" then
			        sSQL = "exec Mail_Anulacion_Giro '" & Session("PerfilServidorCorreo") & "'," & evaluarstr(cod_giro) & ", " & _
			                                                        evaluarstr(Request.Form("txtEmail")) & ", '', " & _
			                                                         evaluarstr(Request.Form("txtNombre")) & ", " & evaluarstr(Motivo) & ", " & _
		                                                            evaluarstr(Session("NombreOperador"))
			        set rs = ejecutarsqlcliente(session("afxCnxServidorCorreo"), sSQL)
			    End If
    			
			End If
			SET rsListaCorreo = Nothing			   
			SET rs = nothing
		'response.redirect "PeticionAnulacion.asp"
		response.Redirect "DetalleGiro.asp?Codigo=" & cod_giro & "&SA=1"
%>