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
	Dim objEMail
	Dim hora, fecha
	Dim sSQL3,sSQL4,rs4,sw,cnn,rs3
	Dim rs2
	hora= Time()
	fecha= Date()
	hora1= replace(time,":","")
	'On Error Resume Next
	Cod_giro = request.Form("Cod_Giro") 
	Monto= request.Form ("Monto")
	Detalle = request.Form ("Detalle")
	Motivo= request.form("txtCorregir")
	Nombre=request.Form ("chkNombre")
	Direccion=request.Form ("chkDireccion")
	Fono=request.Form ("chkFono")
	
	On Error Goto 0

		'On Error Resume Next		
		Set Cnn = server.CreateObject("ADODB.Connection")		
		Cnn.CommandTimeout = 60
		
		sSQL2 = "exec CorregirGiro '" & cod_giro & "','" & Motivo & "','" & session("NombreUsuarioOperador") & "','COR1','" & Motivo &"','" & Session("CodigoAgente") & "'"   
		set rs2 = EjecutarSqlCliente(Session("afxCnxAFEXpress"), sSQL2)
		If Err.number <> 0 Then
			MostrarErrorMS "Solicitud de corrección de datos"
        else
            response.redirect "PeticionCorrecion.asp"
		End If				
		Set rs2 = Nothing	
			
		Set objEMail = Server.CreateObject("CDONTS.NewMail")
		objEMail.From = Request.Form("txtNombre") & " <" & Request.Form("txtEmail") & ">"	
		objEMail.To = "reclamos@afex.cl"
		'objEMail.To = "patricia.sierra@afex.cl"
		'objEMail.cc = "jeannette.barrera@afex.cl"
		objEMail.cc = "transmisiones@afex.cl;"
		objEMail.Subject = "Correccion de Giro. Código:" & Cod_Giro &" (" & fecha &" " & hora &")"
		objEMail.Body = "Se Debe Corregir "&Nombre&", "&Direccion&", "&Fono& Chr(13) & Motivo & Chr(13) &"NOTA: El Giro " & Cod_Giro &" quedo Automaticamente en estado de Reclamo esta es una peticion de " & Session("NombreUsuarioOperador")
		objEMail.Send
		Set objEMail = Nothing

		'response.Write sw
		'response.redirect "PeticionCorrecion.asp"
	

%>