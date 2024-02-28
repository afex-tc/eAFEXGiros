<%@ Language=VBScript %>

<%
	CambiarEstadoTRF Request("corr")
	
	response.Redirect "..\Agente\AtencionClientes.asp"
	
	Sub CambiarEstadoTRF(CorrelativoTRF)
		Dim cnn, sSQL	

		On Error Resume Next
			
		Set cnn = CreateObject("ADODB.Connection")
		cnn.Open Session("afxCnxAFEXchange")
			
		If Err.number <> 0 Then
			cnn.Close
			Set cnn = Nothing
			MostrarErrorMS "Actualizar Estado Transferencia 1"
			Exit Sub
		End If

		sSQL = "update transferencia set estado_transferencia = " & Request("es") & " " & _
			   "where correlativo_transferencia = " & CorrelativoTRF
			
		cnn.Execute sSQL

		If Err.number <> 0 Then
			cnn.Close
			Set cnn = Nothing
			MostrarErrorMS "Actualizar Estado Transferencia 2"
			Exit Sub
		End If
			
		cnn.Close
		Set cnn = Nothing
	End Sub
%>
<!--#INCLUDE virtual="/Compartido/Errores.asp" -->