<%@ Language=VBScript %>
<!--#INCLUDE virtual="/Compartido/Rutinas.asp" -->
<%
	Dim rs
	Dim sSQL
	Dim sValor
	Dim sMensaje
	Dim afx
	
	On Error Resume Next
	Consultar()
	Sub Consultar()
		' rescata el tipo de cambio para venta o compra de dolares y la tarifa para un envio
		Select Case Request("Tipo")
			Case 1 ' compra
				'sSQL = "APObtenerPRActualPro 'USD'"
				'set rs = EjecutarSQLCliente(Session("afxCnxAFEXchange"), sSQL)
				'if	err.number <> 0 then
				'	sMensaje = "Error al consultar el Tipo de Cambio. " & err.Description 
				'	exit sub
				'end if
				'if rs.eof then
				'	sValor = 0
				'else
				'	sValor = ccur(rs("pr_compra")) - ccur(Session("afxRestarTCCompraUSD"))
				'end if
				sValor = ObtenerValorMoneda(1)
				
			
			Case 2	' venta
				'sSQL = "APObtenerPRActualPro 'USD'"
				'set rs = EjecutarSQLCliente(Session("afxCnxAFEXchange"), sSQL)
				'if	err.number <> 0 then
				'	sMensaje = "Error al consultar el Tipo de Cambio. " & err.Description 
				'	exit sub
				'end if
				'if rs.eof then
				'	sValor = 0
				'else
				'	sValor = ccur(rs("pr_venta")) + ccur(Session("afxSumarTCVentaUSD"))
				'end if
				sValor = ObtenerValorMoneda(2)
		
			Case 3	' tarifa de un envío
				Set afx = Server.CreateObject("AFEXGiro.Giro")
				
				If Request("Moneda") = "CLP" Then
					sValor = afx.ObtenerTarifaxp(Session("afxCnxAFEXpress"), afxGiroNacional, Session("CodigoMatriz"), Request("Monto"), _
												Request("Pais") , Request("Ciudad"), "CLP")
				
				Else
					sValor = afx.ObtenerTarifa(Session("afxCnxAFEXpress"), Session("CodigoMatriz"), Request("Monto"), _
												Request("Pais"), Request("Ciudad"), "USD", "USD")
				End If
				
				Set afx = nothing
				If err.number <> 0 Then
					sMensaje = "Error al consultar la Tarifa. " & err.Description 
					exit sub
				End If												
		End Select
	End Sub

	Set rs = nothing

	Function ObtenerValorMoneda(ByVal Tipo)
		Dim sVenta, sCompra, sMoneda
		
		ObtenerValorMoneda = 0		
		
		Set WebServices = server.CreateObject("msxml2.xmlhttp")
		Set myXML = Server.CreateObject("MSXML2.DOMDocument")		
			
		myXML.Async = False
		WebURL = "http://www.afex.cl/eafex/webservices/eafex.asmx/Obtenervalores"
		WebServices.Open "POST",WebURL , False 
		WebServices.setRequestHeader "Content-Type", "text/xml"
		WebServices.setRequestHeader "SOAPAction", "urn:myserver/soap:ThisName/thisMethod"
		WebServices.Send() 	
				
		if WebServices.readyState <> 4 then
			sMensaje = "Transferencia Incompleta" 		
		else		
			if WebServices.status = 200 then ' Respuesta del Servidor OK
				myXML.loadXML(WebServices.responseText)
						
				Set RSSItems = myXML.getElementsByTagName("Valores")
				RSSItemsCount = RSSItems.Length-1
						
				if (RSSItemsCount > 0) then				
					sCompra = 0
					sVenta = 0
					for i = 0 To RSSItemsCount						
						Set RSSItem = RSSItems.Item(i)
						for each child in RSSItem.childNodes
							snodename = snodename & "//" & child.nodeName
							Select case lcase(child.nodeName)
								case "tipocambio"
									If child.text = "USD" Then
										sMoneda = child.text
									End If
								case "moneda"									
									'sMoneda = sMoneda & "1*" & child.text
								case "vcompra"
									If sMoneda = "USD" Then
										sCompra = child.text
									End If
								case "vventa"
									If sMoneda = "USD" Then
										sVenta = child.text
									End If
							End Select											
						next
						if sMoneda <> Empty And sCompra <> Empty And sVenta <> Empty Then exit for	
					next					
				end if			
			else			
				sMensaje = WebServices.statusText
			end if
		end If
		
		Select Case Tipo
			Case 1
				ObtenerValorMoneda = sCompra
			
			Case 2
				ObtenerValorMoneda = sVenta
		End Select
			
			
		Set WebServices = Nothing 
		Set myXML = Nothing		
	End Function

%>


<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<BODY>

<script language="vbscript">
<!--
	window.dialogWidth = 1
	window.dialogHeight = 1
	window.dialogLeft = 1
	window.dialogTop = 1

	if "<%=sMensaje%>" <> Empty then
		window.returnvalue = "<%=sMensaje%>"
	else
		window.returnvalue = "<%=sValor%>"
	end if	
	window.close() 
-->
</script>

</BODY>
</HTML>
