
<%
	 'Los patrones están aquí sin terminar por falta de espacio.


     Function EncriptarCadena(ByVal cadena)
		Dim idx
        Dim result        
        For idx = 1 To len(cadena) 			
			result = result & EncriptarCaracter(mid(cadena, idx, 1), len(cadena), cint(idx) - 1)
			'EncriptarCadena = mid(cadena, idx, 1) & " // " & len(cadena) & " // " & idx & " ** " & result
		Next
        EncriptarCadena = result
     End Function

     Function EncriptarCaracter(ByVal caracter, ByVal variable, ByVal a_indice)

		Dim caracterEncriptado
        Dim indice

		dim patron_busqueda
		dim Patron_encripta 

		patron_busqueda = "ABCDEFGHIJKLMNÑOPQRSTVWXYZabcdefghijklmnñopqrstvwxyz1234567890"
		Patron_encripta = "ABCDEFGHIJKLMNÑOPQRSTVWXYZabcdefghijklmnñopqrstvwxyz1234567890"

        If instr(patron_busqueda, caracter) > 0 Then
			indice = (cint(instr(patron_busqueda, caracter)) + cint(variable) + cint(a_indice)) Mod len(patron_busqueda)
			if cint(indice) > 0 then
				EncriptarCaracter = mid(Patron_encripta, indice, 1)
			else
				EncriptarCaracter = indice
			end if
        else
			EncriptarCaracter = caracter
		End If

		'EncriptarCaracter =	instr(patron_busqueda, caracter)
     End Function


     Function DesEncriptarCadena(ByVal cadena)
		Dim idx
        Dim result

        For idx = 1 To len(cadena)
			result = result & DesEncriptarCaracter(mid(cadena, idx, 1), len(cadena), cint(idx) - 1)
        Next
        DesEncriptarCadena = result
     End Function


     Function DesEncriptarCaracter(ByVal caracter, ByVal variable, ByVal a_indice)

		Dim indice
		
		dim patron_busqueda
		dim Patron_encripta 

		patron_busqueda = "ABCDEFGHIJKLMNÑOPQRSTVWXYZabcdefghijklmnñopqrstvwxyz1234567890"
		Patron_encripta = "ABCDEFGHIJKLMNÑOPQRSTVWXYZabcdefghijklmnñopqrstvwxyz1234567890"

        If instr(Patron_encripta, caracter) > 0 Then
			If (cint(instr(Patron_encripta, caracter)) - cint(variable) - cint(a_indice)) > 0 Then
				indice = (cint(instr(Patron_encripta, caracter)) - cint(variable) - cint(a_indice)) Mod len(Patron_encripta)
            Else
				'La línea está cortada por falta de espacio
				indice = ((cint(len(patron_busqueda)) + cint(instr(Patron_encripta,caracter))) - cint(variable) - cint(a_indice)) Mod len(Patron_encripta)
            End If
			indice = indice Mod len(Patron_encripta)
			if cint(indice) > 0 then
				DesEncriptarCaracter = mid(patron_busqueda,cint(indice), 1)
			else
				DesEncriptarCaracter = indice
			end if
        Else
			DesEncriptarCaracter = caracter
        End If

     End Function

	Function CodificardesdeWS(ByVal Cadena, ByRef MensajeCodificador)
		dim sXML 
		dim i
				
		on error resume next		
		
		sXML = "<?xml version='1.0' encoding='utf-8'?>" & _
				"<soapenv:Envelope xmlns:soapenv='http://schemas.xmlsoap.org/soap/envelope/'>" & _
				"<soapenv:Body>" & _
				"<Codificar xmlns='http://tempuri.org/'>" & _
					"<strACodificar>" & Cadena & "</strACodificar>" & _
				"</Codificar>" & _
			"</soapenv:Body>" & _
			"</soapenv:Envelope>"
				
		CodificardesdeWS = ""
		
		Set WebServices = CreateObject("msxml2.serverxmlhttp")
		Set myXML = CreateObject("MSXML2.DOMDocument")
		Set XMLEnviar = CreateObject("MSXML2.DOMDocument")
		
		XMLEnviar.loadXML(sXML)
			
		myXML.Async = False		
		WebURL = Session("URLwsUtilitariosAFEX")
		WebServices.Open "POST",WebURL , false
		
		
		WebServices.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
		WebServices.setRequestHeader "Content-Length", "length"
		WebServices.setRequestHeader "SOAPAction", "http://tempuri.org/Codificar"
		webservices.send(xmlenviar)
				
		if WebServices.readyState <> 4 then
			MensajeCodificador = "Transferencia Incompleta ." & webservices.responseText & err.Description 			
								
		else
			
			if WebServices.status = 200 then ' Respuesta del Servidor OK
				myXML.loadXML(WebServices.responseText)						
			
								
				Set RSSItems = myXML.getElementsByTagName("CodificarResponse")
				RSSItemsCount = RSSItems.Length								
				if (RSSItemsCount > 0) then
				
					for i = 0 To RSSItemsCount - 1					
						Set RSSItem = RSSItems.Item(i)
						CodificardesdeWS = RSSItem.text												
					next
				end if				
				
			else			
				MensajeCodificador = WebServices.statusText & vbcrlf & err.Description & vbcrl & WebServices.responseText 
				MensajeCodificador = "Error en la conexión para ws. " & vbcrlf & _
									"Detalle del ERROR: " & vbcrlf & _
									MensajeCodificador
			end if
		end If
		
		if err.number <> 0 then
			MensajeCodificador = "Error. " & err.Description
		end if
		
		Set WebServices = Nothing 
		Set myXML = Nothing		
		set xmlenviar = nothing
	End Function


%>