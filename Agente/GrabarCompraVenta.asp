<%@ LANGUAGE = VBScript %>
<%
	'Se asegura que la página no se almacene en la memoria cache
	Response.Expires = 0
	If Session("CodigoCliente") = "" Then
		response.Redirect "../Compartido/TimeOut.htm"
		response.end
	End If
%>
<%	
	Dim afx, afxDSP, nCodigoSP, nOperacion, iLinea
	Dim sCodigo, nPagoAfexEfectivo, nPagoClienteEfectivo
	Dim sOperacion

	On Error	Resume Next	
	CompraVenta
	
	Set afx = Nothing
	Response.Redirect "CompraVentaMoneda.asp"

	'Métodos	
	Sub CompraVenta()
		Set afx = Server.CreateObject("AfexProducto.SP")
		If Err.number <> 0 Then
			Set afx = Nothing
			MostrarErrorMS "Grabar Compra/Venta 1"
		End If
		
		Select Case cInt(0 & request.Form("cbxOperacion"))
		
		Case afxOperacionCompra
			nPagoAfexEfectivo = 0
			nPagoClienteEfectivo = cCur(0 & request.Form("txtTotal"))
			sOperacion = "Compra de Moneda"

		Case afxOperacionVenta
			nPagoAfexEfectivo = cCur(0 & request.Form("txtTotal"))
			nPagoClienteEfectivo = 0
			sOperacion = "Venta de Moneda"
			
		End Select
		
		nCodigoSP = afx.AgregarEncabezado(Session("afxCnxAFEXweb"), , "Web", , Date, "CLP", "CLP", , , 0, 0,  nPagoClienteEfectivo, 0, 0, 0, nPagoAfexEfectivo, 0, 0, , , 0, 0, 0, 0, 0, afxEPEPagado, Session("CodigoCaja"), Session("NombreUsuario"), cDate(Session("FechaApertura")), , afxCierre,,,,,,,,,,, True)

		If Err.number <> 0 Then
			Set afx = Nothing
			MostrarErrorMS "Grabar Compra/Venta 2"
		End If
		If afx.ErrNumber <> 0 Then
			MostrarErrorAFEX afx, "Grabar Compra/Venta 3"
		End If
		If nCodigoSP = 0 Then
			response.Redirect "../compartido/error.asp?Titulo=Grabar Compra/Venta 4&description=Se produjo un error desconocido al intentar grabar la operación"
		End If
				
		If Err.number <> 0 Then
			Set afx = Nothing
			MostrarErrorMS "Grabar Compra/Venta 2"
		End If
		If afx.ErrNumber <> 0 Then
			MostrarErrorAFEX afx, "Grabar Compra/Venta 3"
		End If
		
		Set afxDSP = Server.CreateObject("AfexProducto.SP")
		If Err.number <> 0 Then
			Set afxDSP = Nothing
			MostrarErrorMS "Grabar Compra/Venta 5"
		End If
		
		nOperacion = request.Form("cbxOperacion")
		iLinea = 1
		If cCur(0 & Request.Form("txtSubTotal")) <> 0 Then
			DetalleCompraVenta request.Form("cbxMoneda"), request.Form("cbxProducto"), request.Form("txtMonto"), request.Form("txtTipoCambio"), request.Form("txtSubTotal")
		End If

		iLinea = 2
		If cCur(0 & Request.Form("txtSubTotal1")) <> 0 Then
			DetalleCompraVenta request.Form("cbxMoneda1"), request.Form("cbxProducto1"), request.Form("txtMonto1"), request.Form("txtTipoCambio1"), request.Form("txtSubTotal1") 
		End If
		
		iLinea = 3
		If cCur(0 & Request.Form("txtSubTotal2")) <> 0 Then
			DetalleCompraVenta request.Form("cbxMoneda2"), request.Form("cbxProducto2"), request.Form("txtMonto2"), request.Form("txtTipoCambio2"), request.Form("txtSubTotal2")
		End If
		
		Set afxDSP = Nothing
	End Sub


	Sub DetalleCompraVenta(ByVal Moneda, ByVal Producto, Byval MontoExtranjera, Byval TipoCambio, ByVal MontoNacional)
		Dim bResultado

		bResultado = afx.AgregarDetalle(Session("afxCnxAFEXweb"), nCodigoSP, nOperacion, iLinea, , Session("FechaApertura"), Session("CodigoCaja"), Session("NombreUsuario"), afxCompra, Moneda, "CLP", Producto, afxDocBoleta, 1, afxDocConGlosa, , MontoExtranjera, TipoCambio, MontoNacional,,,,,,,,,,, afxPagado,,,, Session("CodigoCaja"), Date, Time, True)
		
		If Err.number <> 0 Then
			Set afx = Nothing
			MostrarErrorMS "Grabar Compra/Venta 6"
		End If
		If afx.ErrNumber <> 0 Then
			MostrarErrorAFEX afx, "Grabar Compra/Venta 7"
		End If
		If Not bResultado Then
			response.Redirect "../compartido/error.asp?Titulo=Grabar Compra/Venta 8&description=Se produjo un error desconocido al intentar grabar la operación"
		End If
	End Sub
 							 	
%>
<!--#INCLUDE virtual="/Compartido/Rutinas.asp" -->
<!--#INCLUDE virtual="/Compartido/Errores.asp" -->
<html>
<head>
<meta name="VI60_DefaultClientScript" Content="VBScript">
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<title>AFEX Ltda.</title>
</head>
<script language="VBScript">
<!--

	Sub window_onLoad()
		ImprimirBoleta
		window.navigate "AtencionClientes.asp"
	End Sub

	Sub ImprimirBoleta()
		Dim afxTM
		
		Set afxTM = CreateObject("AfexPrinter.TM295")
		
		On Error Resume Next
		afxTM.Inicializar
	   If afxTM.ErrNumber <> 0 Then
			MostrarErrorAFEX afxTM, "Imprimir Boleta 1"	   
		End If
		
		afxTM.Habilitar MsComm1
	   If afxTM.ErrNumber <> 0 Then
			MostrarErrorAFEX afxTM, "Imprimir Boleta 2"	   
		End If
		
		
		afxTM.EncabezadoBoleta MSComm1, 0, Date, "<%=sOperacion%>", Time
	   If afxTM.ErrNumber <> 0 Then
			MostrarErrorAFEX afxTM, "Imprimir Boleta 3"	   
		End If

		afxTM.DetalleBoleta MSComm1, "<%=request.Form("cbxMoneda")%>", _
				cCur(0 & "<%=request.Form("txtMonto")%>"), cCur(0 & "<%=request.Form("txtTipoCambio")%>"), cCur(0 & "<%=request.Form("txtSubTotal")%>")
	   If afxTM.ErrNumber <> 0 Then
			MostrarErrorAFEX afxTM, "Imprimir Boleta 4"	   
		End If

		If cCur(0 & "<%=request.Form("txtSubTotal1")%>") <> 0 Then
			afxTM.DetalleBoleta MSComm1, "<%=request.Form("cbxMoneda1")%>", _
					cCur(0 & "<%=request.Form("txtMonto1")%>"), cCur(0 & "<%=request.Form("txtTipoCambio1")%>"), cCur(0 & "<%=request.Form("txtSubTotal1")%>")
			If afxTM.ErrNumber <> 0 Then
				MostrarErrorAFEX afxTM, "Imprimir Boleta 5"	   
			End If
		End If
		
		If cCur(0 & "<%=request.Form("txtSubTotal2")%>") <> 0 Then
			afxTM.DetalleBoleta MSComm1, "<%=request.Form("cbxMoneda2")%>", _
					cCur(0 & "<%=request.Form("txtMonto2")%>"), cCur(0 & "<%=request.Form("txtTipoCambio2")%>"), cCur(0 & "<%=request.Form("txtSubTotal2")%>")
			If afxTM.ErrNumber <> 0 Then
				MostrarErrorAFEX afxTM, "Imprimir Boleta 6"	   
			End If
		End If

		afxTM.PieBoleta MSComm1, cCur(0 & "<%=request.Form("txtTotal")%>")
	   If afxTM.ErrNumber <> 0 Then
			MostrarErrorAFEX afxTM, "Imprimir Boleta 7"	   
		End If
		afxTM.Deshabilitar MSComm1
				
		Set afxTM = Nothing
	End Sub
	
-->
</script>
<body>
<OBJECT classid="clsid:648A5600-2C6E-101B-82B6-000000000014" id="MSComm1" style="LEFT: 0px; TOP: 0px" VIEWASTEXT>
	<PARAM NAME="_ExtentX" VALUE="1005">
	<PARAM NAME="_ExtentY" VALUE="1005">
	<PARAM NAME="_Version" VALUE="393216">
	<PARAM NAME="CommPort" VALUE="1">
	<PARAM NAME="DTREnable" VALUE="1">
	<PARAM NAME="Handshaking" VALUE="0">
	<PARAM NAME="InBufferSize" VALUE="512">
	<PARAM NAME="InputLen" VALUE="0">
	<PARAM NAME="NullDiscard" VALUE="0">
	<PARAM NAME="OutBufferSize" VALUE="512">
	<PARAM NAME="ParityReplace" VALUE="63">
	<PARAM NAME="RThreshold" VALUE="0">
	<PARAM NAME="RTSEnable" VALUE="0">
	<PARAM NAME="BaudRate" VALUE="9600">
	<PARAM NAME="ParitySetting" VALUE="0">
	<PARAM NAME="DataBits" VALUE="8">
	<PARAM NAME="StopBits" VALUE="0">
	<PARAM NAME="SThreshold" VALUE="0">
	<PARAM NAME="EOFEnable" VALUE="0">
	<PARAM NAME="InputMode" VALUE="0">
</OBJECT>
</body>
</html>
