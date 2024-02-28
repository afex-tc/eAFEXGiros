<%@ LANGUAGE = VBScript %>
<% 
	response.Buffer = True
	response.Clear 
%>
<!--#INCLUDE virtual="/Compartido/Errores.asp" -->
<!--#INCLUDE virtual="/Compartido/Rutinas.asp" -->
<!--#INCLUDE virtual="/sucursal/Rutinas.asp" -->
<%	
	Dim rsFicha
	
	Set rsFicha = ObtenerFichaCliente()
	ActualizarFicha
	Set rsFicha = Nothing		
	
	Response.Redirect "http:NuevoCliente.asp"


	Private Function ActualizarFicha
		Dim nValor
		
		ActualizarFicha = False
			
		Do Until rsFicha.EOF
			If Request.Form(rsFicha("campo")) = "on" Then nValor = 1 Else	nValor = 0
			If Not Actualizar(Session("afxCnxCorporativa"), rsFicha("correlativo"), nValor) Then						  
				response.Redirect "http:../compartido/Error.asp?Titulo=Guardar Actualizacion Ficha 1&Number=" & Err.Number  & "&Source=" & Err.Source & "&Description=" & Err.Description
				Exit Function
			End If
	 		'Session("afxCnxCorporativa") = "DSN=AfexCorporativa;UID=corporativa;PWD=afxsqlcor;"

			If Err.number <> 0 Then
				response.Redirect "http:../compartido/Error.asp?Titulo=Guardar Actualizacion Ficha 2&Number=" & Err.Number  & "&Source=" & Err.Source & "&Description=" & Err.Description
				Exit Function
			End If
			rsFicha.MoveNext
		Loop
	
		ActualizarFicha = True
	End Function	


	Function Actualizar(ByVal Conexion, ByVal Correlativo, ByVal Estado)

		Dim sSQL
		Dim BD
	      
		Actualizar = False
	   
		' controla los errores
		On Error Resume Next
	   

		sSQL = "ActualizarConfiguracionFicha @Correlativo = " & Correlativo & ", @Estado = " & Estado
									
		
		'Conexion
		Set BD = Server.CreateObject("ADODB.Connection")
		BD.Open Conexion                          'Abre la conexion
		If Err.Number <> 0 Then 
			Set BD = Nothing
			Response.Redirect "http:../compartido/error.asp?description=" &  Err.number & "<br>" & Err.Description
		End If
	   
		BD.BeginTrans
	    
		'Consulta
		BD.Execute sSQL                           'Ejecuta la consulta
	    If Err.Number <> 0 Then 
			BD.RollbackTrans
			Set BD = Nothing
			Response.Redirect "http:../compartido/error.asp?description=" &  Err.number & "<br>" & Err.Description
			Exit Function	
	    End If
	   
		BD.CommitTrans
		
		Actualizar = True
	   
		Set BD = Nothing
	End Function

%>
<!--#INCLUDE virtual="/Compartido/Rutinas.asp" -->
