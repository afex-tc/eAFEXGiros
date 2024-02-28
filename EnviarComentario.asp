<%@ Language=VBScript %>
<!--#INCLUDE virtual="/Compartido/Rutinas.asp"-->
<%
	Dim objEMail
	Dim hora, fecha
	hora= replace(time,": "," ")
	fecha= replace(date,"/ "," ")
	fecha= replace(fecha,"- "," ")
	'On Error Resume Next
	On Error Goto 0
	Set objEMail = Server.CreateObject("CDONTS.NewMail")
	objEMail.From = "Contactenos@afex.cl" 'Request.Form("txtNombre") & " <" & Request.Form("txtEmail") & ">"	
	
	' JFMG 05-01-2010 se comenta para asignar la lista desde la BD
	'objEMail.To = "julio.greene@afex.cl,andres.aguilar@afex.cl,domingo.avila@afex.cl,hugo.sepulveda@afex.cl,arturo.munoz@afex.cl; laura.greene@afex.cl"
	'if Request.Form("optComent")=1 then		
	'	objEMail.cc = "julio.greene@afex.cl,andres.aguilar@afex.cl,domingo.avila@afex.cl,hugo.sepulveda@afex.cl,arturo.munoz@afex.cl; laura.greene@afex.cl"
	'else
	'	objEMail.cc = "thomas.greene@afex.cl,domingo.avila@afex.cl,hugo.sepulveda@afex.cl, julio.greene@afex.cl,andres.aguilar@afex.cl; laura.greene@afex.cl"
	'end if	
	dim sPara, sCopia, sMensajeErrorCorreo, i
	sPara = ObtenerCorreoElectronicoContacto(1)
	if sPara = "" then
		sMensajeErrorCorreo = "No se encontró la lista de correos."
	else
		i = instr(sPara, "//")
		sCopia = mid(sPara, i + 2)
		sPara = left(sPara, i - 1)
		
		objEMail.To = sPara
		objEMail.Cc = sCopia
	end if
	' ********** FIN JFMG 05-01-2010 **********************************	
	
	Select Case Request.Form("optComent")
		Case 1
			objEMail.Subject = "AFEX Ltda.  (CN" &" "& fecha &" "& hora &")"
		Case 2
			objEMail.Subject = "AFEX Ltda.  (CN" &" "& fecha &" "& hora &")"
		Case 3
			objEMail.Subject = "AFEX Ltda.  (CN" &" "& fecha &" "& hora &")"
		Case 4
			objEMail.Subject = "AFEX Ltda.  (CN" &" "& fecha &" "& hora &")"
	
	End Select
	objEMail.Body = "Formulario de Contactenos AFEX Ltda." &" "& fecha & Chr(13) & Chr(13) & "Nombre: " & Request.Form("txtNombre") & Chr(13) & "Apellidos: " & Request.Form("txtApellidos") & Chr(13) & "Ciudad: " & Request.Form("txtCiudad") & Chr(13) & "País: " & Request.Form("Pais") & Chr(13) &  "E-mail: " & Request.Form("txtEmail") & Chr(13) & "Teléfono: " & Request.Form("txtFono") & Chr(13) &  Chr(13) &  "1: Consulta 2: Felicitaciones 3: Sugerencia 4: Reclamo" & Chr(13) & Chr(13) & "¿Que tipo de Consulta?: " & Request.Form("optComent") & Chr(13) & "¿Que tipo de Negocio?: " & Request.Form("optComent2") & Chr(13) & "Comentario: " & Request.Form("txtComentario")   				  
				  
	objEMail.Send
	Set objEMail = Nothing

	If Err.number = 0 and sMensajeErrorCorreo = "" then	' JFMG 05-01-2010 se agregó la validación de mensajeerror
		'If Session("CodigoAgente") <> "" Then
		'	Response.Redirect "http:Agente\AtencionClientes.asp"
		'	Response.End 
		'Else
			Response.Redirect "respuesta.asp"
		'End If
	End If
			
%>
<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<title>Enviar Correo</title>
</head>
	<script language="VBScript">
	<!--
	-->
	</script>
<body>

	<table>
		<tr>
			<td>
				<%=sMensajeErrorCorreo & sPara%>
			</td>
		</tr>
	</table>
	
</body>
</html>