<%@ LANGUAGE = VBScript %>
<% 
	Response.Buffer = True
	Response.Clear 
%>
<!--#INCLUDE virtual="/Compartido/Errores.asp" -->
<!--#INCLUDE virtual="/Compartido/Rutinas.asp" -->
<%	
	Dim nCodigoCliente, sNombreCliente, sRutCliente
	
	nCodigoCliente = Request("cc")
	sNombreCliente = Request("nc")
	sRutCliente = Request("rt")
	
	'busco el nombre de documento segun el codigo
	sqldoc="select nombre from tipo_documento where codigo="& Request.Form("cbxTipoDocumento") &""
	Set rs = EjecutarSQLCliente(Session("afxCnxCorporativa"), sqldoc)
	if not rs.eof then 
		nom_documento=rs("nombre") 
		response.write("es: "& nom_documento)
	end if
	'hasta aca busqueda del nombre de documento
	 AgregarDocumento
	'CrearHistorico

	Response.Redirect "http:DetalleCliente.asp?cc=" & nCodigoCliente & "&nc=" & sNombreCliente & _
											 "&rt=" & sRutCliente

Private Sub AgregarDocumento
	Dim afxConexion 
	Dim sSQL
	Dim BD
	
	On Error Resume Next
	
	afxConexion = Session("afxCnxCorporativa") '"DSN=AfexCorporativa;UID=corporativa;PWD=afxsqlcor;"	
	
	Set BD = Server.CreateObject("ADODB.Connection")
		
	BD.Open afxConexion                          'Abre la conexion
	If Err.Number <> 0 Then 
		Set BD = Nothing
		Response.Redirect "http:../compartido/error.asp?description=" &  "GuardarNuevoDocumento"
	End If
	
	Dim  ipUsuario, nombrePC
	ipusuario = request.servervariables("REMOTE_ADDR")
	nombrePC = request.servervariables("REMOTE_HOST")
	
	sSQL = "InsertarDocumentoCliente " & nCodigoCliente & ", " & Request.Form("cbxTipoDocumento") & ", " & _
										 EvaluarStr(Request.Form("txtNumero")) & ", " & _
										 EvaluarStr(Session("NombreUsuarioOperador")) & ", '" & _
										 FormatoFechaSQL(Date) & "', " & EvaluarStr(Left(Time, 8)) & ", " & _
										 EvaluarStr(Request.Form ("txtnombre")) & ", NULL /*imagen */ , '" &_
										  ipusuario & "', '" & nombrePC & "','" & Session("NombreOperador") & "',1"
										 

	BD.BeginTrans
	
	BD.Execute sSQL

    If Err.Number <> 0 Then 
		BD.RollbackTrans
		Set BD = Nothing
		Response.Redirect "http:../compartido/error.asp?description=" &  Err.number & "<br>" & Err.Description
		Exit Sub
    End If
    
    BD.CommitTrans
   
    
    Set BD = Nothing
End Sub	

Private Sub CrearHistorico
	Dim afxConexion2 
	Dim sSQL1
	Dim BD2
	
	On Error Resume Next
	
	afxConexion2 = Session("afxCnxCorporativa") '"DSN=AfexCorporativa;UID=corporativa;PWD=afxsqlcor;"	
	
	Set BD2 = Server.CreateObject("ADODB.Connection")
		
	BD2.Open afxConexion2                          'Abre la conexion
	If Err.Number <> 0 Then 
		Set BD2 = Nothing
		Response.Redirect "http:../compartido/error.asp?description=" &  "GuardarNuevoDocumento"
	End If
	
	det_descripcion="Creación de " & nom_documento
	
	sSQL2="INSERT INTO Historia(fecha, hora, codigo_cliente, sistema, descripcion, usuario, tipo) VALUES ('"& FormatoFechaSQL(Date) &"', "& EvaluarStr(Left(Time, 8)) &", "& nCodigoCliente &", '', '"& det_descripcion &"', "& EvaluarStr(Session("NombreUsuarioOperador")) &", '1')" 
	
	BD2.BeginTrans
	
	BD2.Execute sSQL2

    If Err.Number <> 0 Then 
		BD2.RollbackTrans
		Set BD2 = Nothing
		Response.Redirect "http:../compartido/error.asp?description=" &  Err.number & "<br>" & Err.Description
		Exit Sub
    End If
    
    BD2.CommitTrans
    
    Set BD2 = Nothing

end sub

Function EvaluarStr(ByVal Valor)
	Dim Devuelve
	
	If Valor="" Then 
		EvaluarStr = "Null"	
	Else
		EvaluarStr = "'" & Valor & "'"
	End If

End Function

%>
<!--#INCLUDE virtual="/Compartido/Rutinas.asp" -->
