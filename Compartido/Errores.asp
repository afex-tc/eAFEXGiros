<!-- Errores.asp -->
<%


	Sub MostrarErrorAFEX(ByRef objAFEX, ByVal Titulo)
		
		If Titulo = "" Then Titulo = "Associated Foreign Exchange"
		If Session("Categoria")= 3 or Session("Categoria") = 4 Then
			Response.Redirect  "http://www.moneyexpress.cl/Compartido/Error.asp?Titulo=" & Titulo & _
					"&Number=" & objAFEX.ErrNumber  & _
					"&Source=" & objAFEX.ErrSource & _
					"&Description=" & replace(objAFEX.ErrDescription, vbCrLf , "^")			
		else
		
			Response.Redirect  "/Compartido/Error.asp?Titulo=" & Titulo & _
						"&Number=" & objAFEX.ErrNumber  & _
						"&Source=" & objAFEX.ErrSource & _
						"&Description=" & replace(objAFEX.ErrDescription, vbCrLf , "^")			
		End If
		Set objAFEX = Nothing	
	End Sub



	Function MostrarErrorMS(ByVal Titulo)
		
		If Titulo = "" Then Titulo = "Associated Foreign Exchange"
		
		If Session("Categoria")= 3 or Session("Categoria") = 4 then
			Response.Redirect "http://www.moneyexpress.cl/Compartido/Error.asp?Titulo=" & Titulo & _
						"&Number=" & Err.Number  & _
						"&Source=" & Err.Source & _
						"&Description=" & Err.Description
		else
		
			Response.Redirect "/Compartido/Error.asp?Titulo=" & Titulo & _
						"&Number=" & Err.Number  & _
						"&Source=" & Err.Source & _
						"&Description=" & Err.Description
		End If
	End Function




%>
