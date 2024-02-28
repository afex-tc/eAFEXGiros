<%@ LANGUAGE = VBScript %>
<%
	Response.Expires = 0
	If Session("CodigoCliente") = "" Then
		response.Redirect "../Compartido/TimeOut.htm"
		response.end
	End If
%>
<% 
	Response.Buffer = True
	Response.Clear 
%>
<!--#INCLUDE virtual="/Compartido/Errores.asp" -->
<!--#INCLUDE virtual="/Compartido/Rutinas.asp" -->
<!--#INCLUDE virtual="/sucursal/Rutinas.asp" -->
<%	
	Dim nCodigoCliente, sNombreCliente, sRutCliente
		
	nCodigoCliente = Request("cc")
	sNombreCliente = Request("nc")
	sRutCliente = Request("rt")
	
	
	AgregarHistoria nCodigoCliente, Request.Form("txtDescripcion"), Request.Form("cbxTipo") _
					, Request.Form("cbxOpcionAutorizacion") ' JFMG 16-04-2009 se agrega tipo autorizacion
	
	if	 Request.Form("cbxOpcionAutorizacion")<>"" then	'Trae el codigo de la autorizacion.
		Dim Cuerpo, Asunto
		Asunto = Request.Form("hddTextoAutorizacion") '"Autorización cliente " & trim(sNombres)
				 
		Cuerpo = "Estimado [nombreDestinatario],<br /><br />Se informa que el usuario " & Session("NombreOperador") & _
				 " ha realizado la siguiente acción con el cliente&nbsp;" &  trim(sNombreCliente) & ":" & _
				 "<br /><br /><b>Autorización : " & Request.Form("hddTextoAutorizacion")  & "</b>" & _
				 "<br /><br /><b>Descripción : " & Request.Form("txtDescripcion")  & "</b>" & _
				 "<br /><br /> Atte,<br /><br />Servicio de Mensajería Afex."
		
		EnviarEMailBD 8, 29,Session("AmbienteServidorCorreo"),Asunto, Cuerpo
	end if
	
	Response.Redirect "http:DetalleCliente.asp?cc=" & nCodigoCliente & "&nc=" & sNombreCliente & _
											 "&rt=" & sRutCliente
	

%>
