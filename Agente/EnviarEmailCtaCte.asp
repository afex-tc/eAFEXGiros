<%@ Language=VBScript %>
<%
	EnviarEmail "AfexMoneyWeb <afex@afex.cl>", _
					"Gerencia de Informática <juan.miranda@afex.cl>", "Cuadraturas Nacionales <cuadraturas.nacionales@afex.cl>; Cristian Martinez <cristian.martinez@afex.cl>", _
					"CtaCte." & TRIM(Session("NombreUsuario")) & " " & Request.Form("Desde") & " al " & Request.Form("Hasta"), _
					"<HTML><BODY>" & _
					Request.Form("Documento") & vbCrlf & _					
					"</BODY></HTML>" 
'					"<br>AFEX Ltda.<br>Santiago - Chile" 
'					"Cuadratura de Cuenta Corriente del " & Request.Form("Desde") & " al " & Request.Form("Hasta") & vbCrLf & vbCrLf & 
'					"			Saldo Actual: " & FormatNumber(Request.Form("SaldoActual"), 2) & vbCrLf & 
	Response.Redirect "http:AtencionClientes.asp"
	
Sub EnviarEMail(ByVal Desde, Byval Para, Byval CC, ByVal Asunto, ByVal Mensaje)
	Dim objEMail
		
	On Error Resume Next
	Set objEMail = Server.CreateObject("CDONTS.NewMail")
 
	objEMail.BodyFormat = 0
	objEMail.MailFormat = 0	
	objEMail.From = Desde
	objEMail.To = Para
	objEMail.cc = CC
	objEMail.Subject = Asunto
	objEMail.Body = Mensaje
	objEMail.Send
	Set objEMail = Nothing
	
End Sub
	
%>
