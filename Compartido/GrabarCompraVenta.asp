<%@ LANGUAGE = VBScript %>
<%
	'option explicit	
	'Se asegura que la página no se almacene en la memoria cache
	Response.Expires = 0
	If Session("CodigoCliente") = "" Then
		response.Redirect "TimeOut.htm"
		response.end
	End If
%>
<!--#INCLUDE virtual="/compartido/Errores.asp" -->
<!--#INCLUDE virtual="/compartido/Constantes.asp" -->
<%
	'Se asegura que la página no se almacene en la memoria cache
	Response.Expires = 0
	If session("CodigoCliente") = "" Then
		response.Redirect "TimeOut.asp"
		response.end
	End If
	
	Dim sAFEXchange, sAFEXpress
	Dim sCodigo, nTipoCliente, bExtranjero
	
	On Error	Resume Next
	'Verifica si el usuario puede grabar
	'VerificarBloqueoUsuario	
	
	sAFEXchange = Session("AFEXchange")
		
	GrabarOperacion
	Response.Redirect "..\compartido\Resultado.asp"
	
	Sub GrabarOperacion()
		Dim cMonto, sMoneda
		Dim sGlosaVoucher, bOk
		Dim nOperacion, nFormaPago, cTipoCambio, cMontoNacional

		cMonto = ccur(Cdbl(0 & Request.Form("txtMonto")))
		sMoneda = Request.Form("cbxMoneda")
		nFormaPago = cInt(0 & Request.Form("cbxFormaPago"))
		nOperacion = cInt(0 & Request.Form("cbxOperacion"))
		cTipoCambio = cCur(cDbl(0 & request.Form("txtTipoCambio")))
		cMontoNacional = cCur(cDbl(0 & request.Form("txtTotal")))
		Select Case nFormaPago			
		Case afxEfectivoCLP
		Case afxDepositoCLP
		End Select

'		response.Redirect "../compartido/error.asp?description=" & _
'			nOperacion & ", " & afxProductoTransferencia & ", " & cMonto & ", " & cTipoCambio & ", " & cMontoNacional
		'Transferencia para AFEXchange							  
		Set afxOperacion = Server.CreateObject("AfexWeb.Web")
		If Err.number <> 0 Then
			Set afxOperacion = Nothing		
			MostrarErrorMS "Grabar Compra/Venta 1"
		End If
		bOk = afxOperacion.AgregarPreSP(Session("afxCnxAFEXchange"), _
								  sAFEXchange, Date(), Time(), nOperacion, _
								  afxProductoEfectivo, sMoneda,  _
								  cMonto, cTipoCambio, cMontoNacional,,,,,,,,,nDestinoBoleta)
								  
		If Err.number <> 0 Then
			Set afxOperacion = Nothing		
			MostrarErrorMS "Grabar Compra/Venta 2"
		End If
		If afxOperacion.ErrNumber <> 0 Then
			MostrarErrorAFEX afxOperacion, "Grabar Compra/Venta 3"
		End If		
		Set afxOperacion = Nothing			
		If Not bOk Then
			response.Redirect "error.asp?Titulo=Grabar Compra/Venta 4&description=No se pudo realizar la operación"
		End If
	End Sub	

%>
<!--#INCLUDE virtual="/compartido/Rutinas.asp" -->