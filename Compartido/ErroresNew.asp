<!-- Errores.asp -->
<%


	Sub MostrarErrorAFEX(ByRef objAFEX, ByVal Titulo)
		
		If Titulo = "" Then Titulo = "Associated Foreign Exchange"
		If Session("ModoPrueba") Then
			Response.Redirect  "http://192.168.111.12/afexmoneyweb/Compartido/Error.asp?Titulo=" & Titulo & _
						"&Number=" & objAFEX.ErrNumber  & _
						"&Source=" & objAFEX.ErrSource & _
						"&Description=" & replace(objAFEX.ErrDescription, vbCrLf , "^")			
		Else
			Response.Redirect  "http://200.72.160.51/afexmoneyweb/Compartido/Error.asp?Titulo=" & Titulo & _
						"&Number=" & objAFEX.ErrNumber  & _
						"&Source=" & objAFEX.ErrSource & _
						"&Description=" & replace(objAFEX.ErrDescription, vbCrLf , "^")			
		End If
		Set objAFEX = Nothing	
	End Sub


	Function MostrarErrorMS(ByVal Titulo)
		
		If Titulo = "" Then Titulo = "Associated Foreign Exchange"
		If Session("ModoPrueba") Then
			Response.Redirect "http://192.168.111.12/afexmoneyweb/Compartido/Error.asp?Titulo=" & Titulo & _
						"&Number=" & Err.Number  & _
						"&Source=" & Err.Source & _
						"&Description=" & Err.Description
		Else
			Response.Redirect "http://200.72.160.51/afexmoneyweb/Compartido/Error.asp?Titulo=" & Titulo & _
						"&Number=" & Err.Number  & _
						"&Source=" & Err.Source & _
						"&Description=" & Err.Description
		End If		
	End Function


%>
