<%@ Language=VBScript %>
<%
		'Se asegura que la página no se almacene en la memoria cache
		Response.Expires = 0
		If session("afxCnxAFEXchange") = "" Then
			response.Redirect "../Compartido/TimeOut.htm"
			response.end
		End If
%>
<%
	Const ColFecha = 0
	Const ColCodigoSP = 1
	Const ColCodigoProducto = 2
	Const ColCodigoMoneda = 3
	Const ColMontoExt = 4
	Const ColTC = 5
	Const ColMontoNac = 6
	Const ColCheque = 7
	Const ColCliente = 8
	Const ColAgente = 9
	Const ColTipoOperacion = 10
		
	Dim rs, Contenido, i, Linea
	Dim Cnn, Sql
	
	'On Error Resume Next
	
	Linea = Split(request.Form("Contenido"), vbCrLf)
	
	For i = 0 To Ubound(Linea) - 1
		Contenido = Split(Linea(i), ";")
		
		Set Cnn = CreateObject("ADODB.Connection")
		Cnn.Open Session("afxCnxCorporativa")
			
		Cnn.BeginTrans

		If i = 0 Then
			Sql = "Delete	Compliance " & _
				  "Where	fecha_solicitud = convert(char, '" & Contenido(ColFecha) & "', 112) " & _
				  "and		codigo_agente = '" & Contenido(ColAgente) & "'"

			Cnn.Execute Sql
		End If
		
		Cnn.Execute "InsertarSPCompliance '" & Contenido(ColFecha) & "', " & _
											   CCur(Contenido(ColCodigoSP)) & ", " & _
											   CInt(Contenido(ColCodigoProducto)) & ", '" & _
											   Contenido(ColCodigoMoneda) & "', " & _
											   FormatoNumeroSQL(Contenido(ColMontoExt))	& ", " & _
											   FormatoNumeroSQL(Contenido(ColTC))	& ", " & _
											   FormatoNumeroSQL(Contenido(ColMontoNac))	& ", '" & _
											   Contenido(ColCheque) & "', '" & _
											   Contenido(ColCliente) & "', '" & _
											   Contenido(ColAgente) & "', '" & _
											   Contenido(ColTipoOperacion) & "'"
											  
		If Err.number <> 0 Then
			Cnn.RollbackTrans
			Set Cnn = Nothing
			MsgBox Err.number & ": " & Err.description
			Exit For		
		End If

		Cnn.CommitTrans
		Set Cnn = Nothing
	Next
	Dim sMenu
	sMenu = ObtenerMenuCliente(Session("CodigoCliente"), "")
	Response.Redirect "http:../Agente/Menu.asp?Menu=" & sMenu & "&cpl=0"
%>
<!--#INCLUDE virtual="/Compartido/Rutinas.asp" -->

