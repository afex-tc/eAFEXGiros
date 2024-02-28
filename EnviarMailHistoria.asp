<%@ Language=VBScript %>
<!--#INCLUDE virtual="/compartido/Rutinas.asp" -->
<%
Dim sSQL
Dim rsHistoria
Dim sMensaje

Dim sDescripcion

On Error Resume Next

' agrega la historia de la operación al mail
Set rsHistoria = BuscarHistoria(Request("Giro"), Request("Agente"))
If err.number <> 0 Then
	sMensaje = "Ocurrió un error. " & err.Description 	
End If
If rsHistoria.EOF Then
	sMensaje = "El giro que Ud. busca no se encuentra entre nuestros datos, puede ser que aun el Agente Captador no lo haya enviado."

Else
	EnviarMail
End If

Sub EnviarMail()
	sDescripcion = "					AFEX tiene el agrado de enviarle la siguiente información acerca del giro que " & vbCrlf & _
						"Ud. a enviado a traves de nuestra Empresa: " & _
						vbCrlf & vbCrlf 

	sDescripcion = sDescripcion & vbCrlf & "HISTORIA: "	
	
	'Do While Not rsHistoria.EOF
	'	sDescripcion = sDescripcion & vbCrlf & Trim(rsHistoria("fecha")) & " " & Trim(rsHistoria("hora")) & " " & Trim(rsHistoria("descripcion"))
	'	rsHistoria.MoveNext
	'Loop	
	
	HTML = "<!DOCTYPE HTML PUBLIC ""-//IETF//DTD HTML//EN"">" & NL
	HTML = HTML & "<html>" 
	HTML = HTML & "<head> "	
	HTML = HTML & "<META NAME=GENERATOR Content=Microsoft Visual Studio 6.0>"
	HTML = HTML & "<title>Mail</title>" 	
	HTML = HTML & "</head>"
	HTML = HTML & "<body link=""white"" alink=""white"" vlink=""white""> "
	HTML = HTML & "<a href=""http://www.afex.cl""><img src=""http://jfmiranda/eafex/images/afexemail.jpg""></a><br><br><br><br>" 
	HTML = HTML & "<table border=0 width=100%" & " style=""font-family: Tahoma; font-size: 12px"">"
	HTML = HTML & "<tr><td width=20%" & " style=""font-family: Tahoma; font-size: 12px""><b>REMITENTE :</b></td><td width=40%" & ">" & rsHistoria("remitente") & "</td></tr>"
	HTML = HTML & "<tr><td>&nbsp;</td></tr>"
	HTML = HTML & "<tr><td><b>BENEFICIARIO :<b></td><td>" & rsHistoria("beneficiario") & "</td></tr>"
	HTML = HTML & "<tr><td><b>PAIS :</b></td><td>" & rsHistoria("paisb") & "</td><td><b>CIUDAD :</b></td><td>" & rsHistoria("ciudadb") & "</td></tr>"
	HTML = HTML & "<tr><td><b>DIRECCIÓN :</b></td><td>" & rsHistoria("direccionb") & "</td><td><b>FONO :</b></td><td>" & rsHistoria("fonob") & "</td></tr>"
	HTML = HTML & "<tr><td><b>MONTO :</b></td><td>" & rtrim(rsHistoria("moneda")) & "$ " & rsHistoria("montogiro") & "</td></tr>"
	HTML = HTML & "</table><br><br>"
	HTML = HTML & "<center style=""font-family: Tahoma; font-size: 12px""><b>HISTORIA</b></center>"		
	HTML = HTML & "<table border=1 width=100%" & " style=""font-family: Tahoma; font-size: 12px"">"
	HTML = HTML & "<tr><td style=""font-family: Tahoma; font-size: 12px""><b>Fecha</b></td><td><b>Hora</b></td><td><b>Descripción</b></td></tr>"
	Do Until rsHistoria.EOF	
		HTML = HTML & "<tr><td>" & rsHistoria("fecha") & "</td>" 
		HTML = HTML & "	 <td>" & rsHistoria("hora") & "</td>"
		HTML = HTML & "	 <td>" & rsHistoria("descripcion") & "</td>"
		HTML = HTML & "</tr>"
					
		rsHistoria.MoveNext
	Loop
	HTML = HTML & "</table>"
	HTML = HTML & "<br><br><div style=""font-family: Tahoma; font-size: 12px""><b>ATTE.  AFEX LTDA.</b></div>" 
	HTML = HTML & "</body>" 
	HTML = HTML & "</html>" 
	Set rsHistoria = Nothing
	
	'response.Write html
	'response.End 
	
	' envia el mail
	EnviarEmail "AFEX", Request("Mail"), "", "Consulta estado Giro " & Request("Giro"), HTML, 1 'sDescripcion
	If err.number <> 0 Then
		sMensaje "Ocurrió un error. " & err.Description
	
	Else
		sMensaje = "Los datos han sido enviados a su e-mail."
	End If	
End Sub

Function BuscarHistoria(ByVal Giro, ByVal Agente)
	Dim rsHistoria
	Dim sSQL
	
'	On Error Resume Next
	
	Set BuscarHistoria = Nothing
	
	' busca la historia del giro enviado
	sSQL = " BuscarHistoria " & EvaluarSTR(Giro) & ", 2, " & EvaluarStr(Agente)	
	
	Set rsHistoria = EjecutarSQLCliente(Session("afxCnxAFEXpress"), sSQL)
	If err.number <> 0 Then
		err.Raise 50000,"BuscarHistoria", "Error al ejecutar la función BuscarHistoria. " & err.Description
		exit function
	End If
	
	Set BuscarHistoria = rsHistoria
	Set rsHistoria = Nothing
End Function

%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<script language="VBScript">
<!--
	window.returnvalue = "<%=sMensaje%>"
	window.close	
-->
</script>

<BODY>

<P>&nbsp;</P>

</BODY>
</HTML>
