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
	
	Set rsFicha = ObtenerProductoFicha()
	ActualizarFicha
	Set rsFicha = Nothing		
	
	Response.Redirect "http:ConfiguracionFicha.asp"


	Private Function ActualizarFicha
		Dim nValor
		
		ActualizarFicha = False
		
		Eliminar Session("afxCnxCorporativa")
		If Err.number <> 0 Then
			response.Redirect "http:../compartido/Error.asp?Titulo=Eliminar Ficha Producto 1&Number=" & Err.Number  & "&Source=" & Err.Source & "&Description=" & Err.Description
			Exit Function
		End If
		
		Do Until rsFicha.EOF
			If Request.Form(rsFicha("campo") & "_" & rsFicha("Producto")) = "on" Then
				If Not Insertar(Session("afxCnxCorporativa"), rsFicha("campo"), rsFicha("producto"), cCur(0 & Request.Form("txt" & rsFicha("campo") & "_" & rsFicha("producto")))) Then	  
					response.Redirect "http:../compartido/Error.asp?Titulo=Insertar Ficha Producto 1&Number=" & Err.Number  & "&Source=" & Err.Source & "&Description=" & Err.Description
					Exit Function
				End If
			End If
	 		'Session("afxCnxCorporativa") = "DSN=AfexCorporativa;UID=corporativa;PWD=afxsqlcor;"

			If Err.number <> 0 Then
				response.Redirect "http:../compartido/Error.asp?Titulo=Insertar Ficha Producto 2&Number=" & Err.Number  & "&Source=" & Err.Source & "&Description=" & Err.Description
				Exit Function
			End If
			rsFicha.MoveNext
		Loop
	
		ActualizarFicha = True
	End Function	


	Function Eliminar(ByVal Conexion)

		Dim sSQL
		Dim BD
	      
		Eliminar = False
	   
		' controla los errores
		On Error Resume Next
	   

		sSQL = "DELETE Ficha_Producto WHERE tipo = 'MCS' "
									
		
		'Conexion
		Set BD = Server.CreateObject("ADODB.Connection")
		BD.Open Conexion                          'Abre la conexion
		If Err.Number <> 0 Then 
			Set BD = Nothing
			Response.Redirect "http:../compartido/error.asp?titulo=Eliminar Ficha Producto 1&description=" &  Err.number & "<br>" & Err.Description
		End If
	   
		BD.BeginTrans
	    
		'Consulta
		BD.Execute sSQL                           'Ejecuta la consulta
	    If Err.Number <> 0 Then 
			BD.RollbackTrans
			Set BD = Nothing
			Response.Redirect "http:../compartido/error.asp?titulo=Eliminar Ficha Producto 2&description=" &  Err.number & "<br>" & Err.Description
			Exit Function	
	    End If
	   
		BD.CommitTrans
		
		Eliminar = True
	   
		Set BD = Nothing
	End Function


	Function Insertar(ByVal Conexion, ByVal Correlativo, ByVal Producto, ByVal MontoDesde)

		Dim sSQL
		Dim BD
	      
		Insertar = False
	   
		' controla los errores
		On Error Resume Next
	   
		sSQL = "InsertarFichaProducto @Campo = " & Correlativo & ", @Producto = " & Producto & ", @Monto_Desde=" & FormatoNumeroSQL(MontoDesde)
		
		'Conexion
		Set BD = Server.CreateObject("ADODB.Connection")
		BD.Open Conexion                          'Abre la conexion
		If Err.Number <> 0 Then 
			Set BD = Nothing
			Response.Redirect "http:../compartido/error.asp?titulo=Insertar Ficha Producto 1&description=" &  Err.number & "<br>" & Err.Description
		End If
	   
		BD.BeginTrans
	    
		'Consulta
		BD.Execute sSQL                           'Ejecuta la consulta
	    If Err.Number <> 0 Then 
			BD.RollbackTrans
			Set BD = Nothing
			Response.Redirect "http:../compartido/error.asp?titulo=Insertar Ficha Producto 2&description=" &  Err.number & "<br>" & Err.Description
			Exit Function	
	    End If
	   
		BD.CommitTrans
		
		Insertar = True
	   
		Set BD = Nothing
	End Function

%>
<!--#INCLUDE virtual="/Compartido/Rutinas.asp" -->
