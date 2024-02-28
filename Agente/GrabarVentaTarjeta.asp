<%@ Language=VBScript %>
<%
	'Se asegura que la página no se almacene en la memoria cache
	Response.Expires = 0
	If Session("CodigoAgente") = "" Then
		response.Redirect "../Compartido/TimeOut.htm"
		response.end
	End If
%>
<%
	Dim sNombreMoneda, sNumeroPIN, nCorrelativo
	Dim cObservado
	
	On Error Resume Next
	'sNombreMoneda=request.QueryString("nmn")
	sNombreMoneda = Trim(Request("nmn"))
	'tt=Request.Form("cbxMonto")
	sNumeroPIN = ""
	nCorrelativo = 0
	cObservado = 0
	'Response.Redirect "../Compartido/error.asp?Titulo=Grabar Venta de Tarjeta&Description=" & sNombreMoneda & " " & Request.Form("cbxMonto")
	
	ObtenerDatosPIN
	GrabarVenta
	'response.Write sNombreMoneda
'	'Response.Redirect "../Compartido/error.asp?Titulo=Grabar Venta de Tarjeta&Description=PIN: " & sNumeroPIN & " Correlativo: " & nCorrelativo
	
    Response.Redirect "VenderTarjeta.asp?hb=disabled&mn=" & Trim(Request.Form("cbxMoneda")) & _
									"&mto=" & Request.Form("cbxMonto") & "&bs=" & Request.Form("txtNumeroBoleta") & _
									 "&np=" & sNumeroPIN & "&tc=" & cObservado & "&sw=3"
	
	
	
	Sub GrabarVenta
		Dim afxTarjeta
		
		On Error Resume Next
		
		Set afxTarjeta = CreateObject("AFEXProducto.TarjetaTelefonica")
		'response.Write Request.Form("cbxMoneda")
		
		If Trim(Request.Form("cbxMoneda")) = "USD" Then
			cObservado = afxTarjeta.ObtenerObservado(Session("afxCnxAFEXpress"))
		End If
		 
		If Not afxTarjeta.Vender(Session("afxCnxAFEXpress"), Session("CodigoAgente"), _
								 TRIM(Request.Form("cbxMoneda")), Request.Form("cbxMonto"), _
								 cObservado, sNumeroPIN, _
								 nCorrelativo, Request.Form("txtNumeroBoleta"), _
								 Session("NombreUsuarioOperador")) Then
								' response.Write nCorrelativo
			If afxTarjeta.ErrNumber <> 0 Then
				Set afxTarjeta = Nothing
				Response.Redirect "../Compartido/error.asp?Titulo=Grabar Venta de Tarjeta&Description=" & afxTarjeta.ErrDescription
				Exit Sub
			End If
			If Err.number <> 0 Then
				Set afxTarjeta = Nothing
				Response.Redirect "../Compartido/error.asp?Titulo=Grabar Venta de Tarjeta&Description=" & Err.Description
				Exit Sub
			End If
		End If
		
		If afxTarjeta.ErrNumber <> 0 Then
			Set afxTarjeta = Nothing
			Response.Redirect "../Compartido/error.asp?Titulo=Grabar Venta de Tarjeta&Description=" & afxTarjeta.ErrDescription
			Exit Sub
		End If
		If Err.number <> 0 Then
			Set afxTarjeta = Nothing
			Response.Redirect "../Compartido/error.asp?Titulo=Grabar Venta de Tarjeta&Description=" & Err.Description
			Exit Sub
		End If
		
		Set afxTarjeta = Nothing
	End Sub
	
	Function ObtenerDatosPIN()
		Dim rs, Sql
		
		On Error Resume Next
		
		ObtenerDatosPIN = False
		
		'Sql = "select top 1 min(correlativo) as correlativo, " & _
		'	  "numero_pin from tarjeta " & _
		'	  "where codigo_giro is null and numero_boleta is null " & _
		'	  "and   tipo_PIN = 2 " & _
		'	  "and   monto = " & Request.Form("cbxMonto") & " " & _
		'	  "and   codigo_moneda = '" & Trim(Request.Form("cbxMoneda")) & "' " & _
		'	  "group by numero_pin " & _
		'	  "order by correlativo asc"
		Sql = "select correlativo, numero_pin " & _
			  "from 	tarjeta " & _
			  "where 	correlativo = (select min(correlativo) " & _
									  "from tarjeta " & _
									  "where codigo_giro is null " & _
									  "and 	numero_boleta is null " & _
									  "and  tipo_PIN = 2 " & _
									  "and 	monto = " & Request.Form("cbxMonto") & " " & _
									  "and 	codigo_moneda = '" & Trim(Request.Form("cbxMoneda")) & "')"
		'response.Write sql								  
	
		Set rs = CreateObject("ADODB.Recordset") 
		
		rs.Open Sql, Session("afxCnxAFEXpress"), 3, 1
		
		If Err.number <> 0 Then
			Set rs = Nothing
			MostrarErrorMS "ObtenerDatosPIN", Err.Description
			Exit Function
		End If
		
		If rs.EOF Then
			sNumeroPIN = ""
			nCorrelativo = 0
			
			Set rs = Nothing
			Exit Function
		End If
		
		sNumeroPIN = rs("numero_pin")
		nCorrelativo = rs("correlativo")
		
		ObtenerDatosPIN = True
		
		Set rs = Nothing
	End Function
	
	
%>
<HTML>
<HEAD>
<META name="VI60_DefaultClientScript" Content="VBScript">
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<BODY>
</BODY>
</HTML>
