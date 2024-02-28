<%@ Language=VBScript %>
<!--#INCLUDE virtual="LoginSSL.asp" -->
<%

Dim prmResultado	'Variable de salida de tipo String, con el resultado de la accion

Dim sURL, nTipo, sErrorUsuario
Dim sCodigoGiro, prmURL
Const afxACCListaGiro = 20
Const afxACCPagarGiro = 21
Const afxACCAvisarGiro = 22
Const afxACCReclamarGiro = 23

On Error Resume Next
prmURL = Request("prmURL")
If Not ValidarClienteSSL(sURL, sErrorUsuario, nTipo, Request("sslu"), request("sslp")) Then
	prmResultado = Request("sslu") & ", " & request("sslp")
	'Response.Write prmResultado
	Response.Redirect prmURL
	Response.End 
Else	

	Select Case cInt(0 & Request("acc"))
	Case afxACCListaGiro
		prmResultado = EjemploAGT()
		If prmURL = "" Then
			Response.Write prmResultado
			Response.Write object2.GetPreviousURL 	
		Else
			Response.Redirect prmURL & "?prmOut=" & prmResultado
		End If
		Response.End 
		
	Case afxACCPagarGiro
		
		sCodigoGiro = BuscarGiroOrden(Session("afxCnxAFEXpress"), cCur(0 & Request("pg1")), Session("CodigoAgente"))
		If Err.number <> 0 Then
			Response.Write Err.description 
		ElseIf PagarGiro Then
			Response.Write ""
		Else
			Response.Write ""			
		End If		
		Response.End 

	Case afxACCAvisarGiro
		
		sCodigoGiro = BuscarGiroOrden(Session("afxCnxAFEXpress"), cCur(0 & Request("pg1")), Session("CodigoAgente"))
		If Err.number <> 0 Then
			Response.Write Err.description 			
		ElseIf AvisarGiro Then
			Response.Write ""
		Else
			Response.Write Err.description 
		End If		
		Response.End 

	Case afxACCReclamarGiro
		
		sCodigoGiro = BuscarGiroOrden(Session("afxCnxAFEXpress"), cCur(0 & Request("pg1")), Session("CodigoAgente"))
		
		If Err.number <> 0 Then
			Response.Write Err.description 			
		ElseIf ReclamarGiro Then
			Response.Write ""
		Else
			Response.Write Err.description 
		End If		
		Response.End 
		
	End Select	
	
End If


Function EjemploAGT()
	Dim obj
	Dim sLista
		
	Set obj = Server.CreateObject("AFEXtoAGT3.AGT")
	sLista = obj.ListaGiro("", Session("CodigoAgente"))
	Set obj = Nothing
	EjemploAGT = sLista

End Function


'Métodos	
Function PagarGiro()
	Dim bVoucher
	Dim afxPago

	PagarGiro=False		
	bVoucher = True
	Set afxPago = Server.CreateObject("AfexGiroXP.Giro")
	If Err.number <> 0 Then
		Set afxPago = Nothing
		Exit Function
	End If
	
	If Session("Categoria") = 3 And Session("CodigoAgente") <> Session("CodigoMoneyBroker") Then
		bVoucher = False
	End If
	Giro = afxPago.Pagar(Session("afxCnxAFEXpress"), sCodigoGiro, Session("CodigoAgente"), _
						 Request("pg2"), cInt(0 & Request("pg3")), Request("pg4"), _
						 Request("pg5"), Request("pg6"), _
						 Request("pg7"), Request("pg8"), Session("NombreUsuario"),,bVoucher)
	If Err.number <> 0 Then
		Set afxPago = Nothing
		Exit Function
	End If						
	If afxPago.ErrNumber <> 0 Then
		Err.Raise afxPago.ErrNumber, AfxPago.ErrSource, afxPago.ErrDescription
		Set afxPago = Nothing
		Exit Function
	End If
	If Not Giro Then
		Set afxPago = Nothing
		Exit Function
	End If
	Set afxPago = Nothing
	PagarGiro = True
	
End Function
 							 	

Function BuscarGiroOrden(ByVal Conexion, _
							    ByVal Orden, Byval Agente)
   Dim sSQL
   Dim rsOrden
	
   'Manejo de errores
   On Error Resume Next
	BuscarGiroOrden = ""
	
   'Crea la consulta
   sSQL = "SELECT    codigo_giro " & _
          "FROM      Giro " & _
          "WHERE     correlativo_salida = " & Orden & " " & _
          "AND		agente_pagador = '" & Agente	& "' "
	
   'Asigna al metodo el resultado de la consulta
   Set rsOrden = EjecutarSQLCliente(Conexion, sSQL)		
   'Si se produjeron errores en la consulta
   If Err.Number <> 0 Then
		Set rsOrden = Nothing
		Exit Function
   End If
	BuscarGiroOrden = rsOrden("codigo_giro")
	Set rsOrden = Nothing
	
End Function


Function EjecutarSQLCliente(ByVal Conexion, ByVal SQL)
   Dim rsESQL
   Const adUseClient = 2
   Const adOpenStatic = 3
   Const adLockBatchOptimistic = 4

   Set EjecutarSQLCliente = Nothing
   Set rsESQL = server.CreateObject("ADODB.Recordset")
   rsESQL.CursorLocation = 3
   rsESQL.Open SQL, Conexion, 3, 4
	If Err.number <> 0 Then
		Set rsESQL = Nothing
		Exit Function
	End If
   Set rsESQL.ActiveConnection = Nothing
   Set EjecutarSQLCliente = rsESQL
   Set rsESQL = Nothing
   
End Function	

Function AvisarGiro()
	Dim afxAG, bAG
	AvisarGiro = False
	
	Set afxAG = Server.CreateObject("AfexGiro.Giro")
	If Err.number <> 0 Then
		Set afxAG = Nothing
		Exit Function
	End If
	bAG = afxAG.Avisar(Session("afxCnxAFEXpress"), sCodigoGiro, _
						 Request("pg2"), Request("pg3"), Session("NombreUsuario"), Request("pg4"), Request("pg5"))
	If Err.number <> 0 Then
		Set afxAG = Nothing		
		Exit Function
	End If						
	If afxAG.ErrNumber <> 0 Then
		Err.Raise afxAG.ErrNumber, AfxAG.ErrSource, afxAG.ErrDescription
		Set afxAG = Nothing	
		Exit Function
	End If
	'Err.Raise 1, "1s", bag
	Set afxAG = Nothing
	AvisarGiro = True
End Function

Function ReclamarGiro()

	Dim afxRG, bRG
	ReclamarGiro = False
	
	Set afxRG = Server.CreateObject("AfexGiro.Giro")
	If Err.number <> 0 Then
		Set afxRG = Nothing
		Exit Function
	End If
	bAG = afxRG.Reclamar(Session("afxCnxAFEXpress"), sCodigoGiro, _
						 Request("pg2"), Request("pg3"), Session("CodigoAgente"), Session("NombreUsuario"))
	If Err.number <> 0 Then
		Set afxRG = Nothing		
		Exit Function
	End If						
	If afxRG.ErrNumber <> 0 Then
		Err.Raise afxRG.ErrNumber, AfxRG.ErrSource, afxRG.ErrDescription
		Set afxRG = Nothing	
		Exit Function
	End If
	Set afxRG = Nothing
	ReclamarGiro = True
End Function
	
%>
<script language="vbscript">
	sub window_onload()
	end sub
</script>
