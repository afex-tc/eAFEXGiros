<%@ Language=VBScript %>
<!--#INCLUDE virtual="/Compartido/Constantes.asp" -->
<!--#INCLUDE virtual="/Compartido/Errores.asp" -->
<!--#INCLUDE virtual="/Compartido/Rutinas.asp" -->
<%

	' JFMG 18-08-2008 Sesión para poder ver el Visor de precios que se encontraba en la página antigua
	Session("AutorizaVisorAP") = 1
	' ***************************** FIN ************************************

	Dim sLista, bExtranjero
	   
	If Session("PaisCliente") <> Session("PaisMatriz") Then		
		bExtranjero = True
	Else
		bExtranjero = False
	End If
	
	sLista = Request("Menu")
	If sLista = "*" Then
		sLista = "10;11;12;13;14;15;16;20;21;22;23;30;31;32;33;40"
	End If
	
	Sub Abandonar
		Session.Abandon 
	End Sub

	
%>

<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<title>AFEX</title>
<link rel="stylesheet" type="text/css" href="../Estilos/MenuCliente.CSS">
</head>

<script LANGUAGE="VBScript">
<!--

	Sub window_onLoad()
		Dim sId
		
		objmenu.bgColor = document.bgColor 
		objmenu.stylesheet = "../Estilos/MenuCliente.css"
		<% If InStr(sLista, "10") <> 0 Then %>
			sId = objmenu.addparent("Consultas")
			<%	If Session("ModoPrueba") Then %>
					'objMenu.addchild sId, "Giros Pendientes de Pago New", "../Agente/ListaGirosNew.asp?Tipo=<%=afxGirosPendientes%>&Agente=<%=Session("CodigoAgente")%>", "Principal"
					'objMenu.addchild sId, "Giros Pendientes de Aviso New", "../Agente/ListaGirosNew.asp?Tipo=<%=afxGirosAviso%>&Agente=<%=Session("CodigoAgente")%>&Pagador=<%=Session("CodigoAgente")%>", "Principal"					
					'objMenu.addchild sId, "Giros Reiteración de Aviso New", "../Agente/ListaGirosNew.asp?Tipo=<%=afxGirosReiteraAviso%>&Agente=<%=Session("CodigoAgente")%>&Pagador=<%=Session("CodigoAgente")%>", "Principal"
			<%	End If %>
			<% If InStr(sLista, "12") <> 0 Then %>
				<% If Session("CiudadMatriz") = Session("CiudadCliente") Then %>
						'objMenu.addchild sId, "Giros Pendientes de Pago", "../Agente/ListaGiros.asp?Tipo=<%=afxGirosPendientes%>&Agente=<%=Session("CodigoMatriz")%>", "Principal"
				<% Else %>
						objMenu.addchild sId, "Giros Pendientes de Pago", "../Agente/ListaGiros.asp?Tipo=<%=afxGirosPendientes%>&Agente=<%=Session("CodigoAgente")%>", "Principal"
				<% End If %>
			<% End If %>
			<% If InStr(sLista, "15") <> 0 Then %>
				<% If Session("CiudadMatriz") = Session("CiudadCliente") Then %>
				<% Else %>
					objMenu.addchild sId, "Giros Pendientes de Aviso", "../Agente/ListaGiros.asp?Tipo=<%=afxGirosAviso%>&Agente=<%=Session("CodigoAgente")%>&Pagador=<%=Session("CodigoAgente")%>", "Principal"					
					objMenu.addchild sId, "Giros Reiteración de Aviso", "../Agente/ListaGiros.asp?Tipo=<%=afxGirosReiteraAviso%>&Agente=<%=Session("CodigoAgente")%>&Pagador=<%=Session("CodigoAgente")%>", "Principal"
				<% End If %>
			<% End If %>
			
			<% If InStr(sLista, "11") <> 0 Then %>
					objMenu.addchild sId, "Giros Enviados", "../Agente/ListaGiros.asp?Tipo=<%=afxGirosEnviados%>&Agente=<%=Session("CodigoAgente")%>", "Principal"
					objMenu.addchild sId, "Giros Enviados y Anulados", "../Agente/ListaGiros.asp?Tipo=<%=afxGirosEnviados%>&Agente=<%=Session("CodigoAgente")%>&st=<%=afxEstadoGiroNulo%>", "Principal"
			<% End If %>
			<% If InStr(sLista, "12") <> 0 Then %>
					objMenu.addchild sId, "Giros Recibidos", "../Agente/ListaGiros.asp?Tipo=<%=afxGirosRecibidos%>&Pagador=<%=Session("CodigoAgente")%>", "Principal"
					objMenu.addchild sId, "Giros Recibidos y Anulados", "../Agente/ListaGiros.asp?Tipo=<%=afxGirosRecibidos%>&Pagador=<%=Session("CodigoAgente")%>&st=<%=afxEstadoGiroNulo%>", "Principal"
			<% End If %>
			<% If InStr(sLista, "13") <> 0 Then %>
					objMenu.addchild sId, "Transferencias Enviadas", "../Agente/ListaTransfer.asp?Titulo=Transferencias Enviadas&Tipo=<%=afxTrfEnviadas%>&Desde=" & Date() & "&Hasta=" & Date() & "&Usuario=<%=Session("NombreUsuario")%>", "Principal"
					objMenu.addchild sId, "Cheques Solicitados", "../Agente/ListaCheque.asp?Titulo=Cheques Solicitados&Tipo=<%=afxChqSolicitado%>&Desde=" & Date() & "&Hasta=" & Date() & "&Usuario=<%=Session("NombreUsuario")%>", "Principal"
			<% End If %>
			
			
			<% If Trim(Session("CodigoAgente")) = "ZC" Or Trim(Session("CodigoAgente")) = "AX" Or _
				  Trim(Session("CodigoAgente")) = "AG" Or Trim(Session("CodigoAgente")) = "AP" Or _
				  Trim(Session("CodigoAgente")) = "AR" Or Trim(Session("CodigoAgente")) = "AH" Or _
				  Trim(Session("CodigoAgente")) = "AZ" Or Trim(Session("CodigoAgente")) = "AO" Or _
				  Trim(Session("CodigoAgente")) = "AQ" Or Trim(Session("CodigoAgente")) = "AI" Or _
				  Trim(Session("CodigoAgente")) = "AY" Or Trim(Session("CodigoAgente")) = "ZB" Or _
				  Trim(Session("CodigoAgente")) = "AE" Or Trim(Session("CodigoAgente")) = "ZD" Or _
				  Trim(Session("CodigoAgente")) = "AM" Or Trim(Session("CodigoAgente")) = "AC" Or _
				  Trim(Session("CodigoAgente")) = "ZE" Or Trim(Session("CodigoAgente")) = "ZF" Or _
				  Trim(Session("CodigoAgente")) = "AN" Or Trim(Session("CodigoAgente")) = "AA" Or _
				  Trim(Session("CodigoAgente")) = "AV" Or Trim(Session("CodigoAgente")) = "AT" Or _
				  Trim(Session("CodigoAgente")) = "AD" Or Trim(Session("CodigoAgente")) = "AJ" or _
				  Trim(Session("CodigoAgente")) = "ZH" Or Trim(Session("CodigoAgente")) = "ZL" or _ 
				  Trim(Session("CodigoAgente")) = "ZK" Or Trim(Session("CodigoAgente")) = "ZG" or _
				  Trim(Session("CodigoAgente")) = "ZM" or Trim(Session("CodigoAgente")) = "ZN" or _
				  Trim(Session("CodigoAgente")) = "ZO" Then %>
				  
					objMenu.addchild sId, "Tarjetas Vendidas", "../Agente/ListaTarjeta.asp?sc=<%=Session("CodigoAgente")%>", "Principal"
			<% End If %>
			
			<% If InStr(sLista, "14") <> 0 Then %>
				objMenu.addchild sId, "Saldo de Monedas", "../Agente/ListaSaldoMonedas.asp", "Principal"
				objMenu.addchild sId, "Compras y Ventas", "../Agente/ListaCompraVenta.asp?Tipo=5&Titulo=Compras y Ventas&Desde=" & Date & "&Hasta=" & Date, "Principal"
			<% End If %>
			
			<% If InStr(sLista, "15") <> 0 Then %>
				objMenu.addchild sId, "Oficinas de Pago", "../Agente/LugaresPago.asp", "Principal"
			<% End If %>
			
			'<% If InStr(sLista, "16") <> 0 Then %>
			'	objMenu.addchild sId, "Paridades", "../Compartido/Paridades.asp", "Principal"	
			'<% End If %>
				'objMenu.addchild sId, "Tipos de Cambio", "../Compartido/ConsultaTipoCambio.asp", "Principal"
				'objMenu.addchild sId, "Direcciones en Chile", "../Sucursales.asp?Region=13", "Principal"
		<% End If %>		
		'objMenu.addchild sId, "Tarjetas Vendidas", "../Agente/ListaTarjeta.asp?sc=<%=Session("CodigoAgente")%>", "Principal"
		
		<% If session("Categoria") <= 2 Then %>
			objMenu.addchild sId, "Precios", "PRECIOS", "_top"
			
			' JFMG 18-08-2008 Opción para ver el visor de precios de la página antigua
			objMenu.addchild sId, "Visor", "VISOR", "_top"
			' **************************** FIN ***************************************
			
			' Jonathan Miranda G. 08-08-2008 menú para agregar pagos AMEX
			objMenu.addchild sId, "Agregar Pago AMEX", "PagoAMEX", "Principal"
			objMenu.addchild sId, "Listado Pagos AMEX", "ListadoPagoAMEX", "Principal"
			'---------------------- Fin -----------------------------
			
			' JFMG 20-07-2009 se agrega lista de tarjetas no consideradas por afex para pagar
			objMenu.addchild sId, "Lista Tarjetas AMEX No Consideradas", "ListaAMEXNoConsiderada", "Principal"
			' ************* FIN JFMG 20-07-2009 ******************			

			' JFMG 03-12-2008 menú para agregar el voucher de giros en pesos en el AFEXchange
			objMenu.addchild sId, "Voucher Giros en Pesos", "VoucherGiroPeso", "Principal"
			'---------------------- Fin -----------------------------
			
		<% End If %>
		
		<% If InStr(sLista, "20") <> 0 Then %>
			sId = objmenu.addparent("Transacciones")
			<% If InStr(sLista, "21") <> 0 Then %>
				<% If bExtranjero Then %>
						'objMenu.addchild sId, "Enviar un Giro", "../Agente/EnviarGiro.asp", "Principal"
				<% End If %>
			<% End If %>
			<% If InStr(sLista, "22") <> 0 Then %>
				<% If bExtranjero Then %>
						'objMenu.addchild sId, "Enviar una Transferencia", "../Agente/EnviarTransfer.asp", "Principal"
						'objMenu.addchild sId, "Enviar un Cheque", "../Agente/EnviarCheque.asp", "Principal"
				<% End If %>
			<% End If %>
			<% If InStr(sLista, "24") <> 0 Then %>
				'If Not bExtranjero Then 
					objMenu.addchild sId, "Agregar Nuevo Cliente", "../Agente/NuevoCliente.asp", "Principal"
				'End If
			<% End If %>
			
			
			' JFMG 24-05-2010 posibilidad de que un cajero agregue clientes corporativos			
			<% If Session("VerClienteCorporativo") Then %>				
				objMenu.addchild sId, "Cliente Corporativo", "<%=Session("URLIngresarClienteCorporativo")%>" & "?TipoOrigenLlamada=1&SCC=<%=Session("CodigoCliente")%>&AGC=<%=Session("CodigoAgente")%>&NUO=<%=Session("NombreUsuarioOperador")%>", "Principal"
			<% End If %>
			' FIN 24-05-2010 
			
			
			<% If InStr(sLista, "14") <> 0 Then %>
				objMenu.addchild sId, "Comprar y Vender", "../Agente/CompraVentaMoneda.asp?Tipo=<%=afxAgente%>", "Principal"
			<% End If %>
			<% If Trim(Session("CodigoAgente")) = "ZC" Or Trim(Session("CodigoAgente")) = "AX" Or _
				  Trim(Session("CodigoAgente")) = "AG" Or Trim(Session("CodigoAgente")) = "AP" Or _
				  Trim(Session("CodigoAgente")) = "AR" Or Trim(Session("CodigoAgente")) = "AH" Or _
				  Trim(Session("CodigoAgente")) = "AZ" Or Trim(Session("CodigoAgente")) = "AO" Or _
				  Trim(Session("CodigoAgente")) = "AQ" Or Trim(Session("CodigoAgente")) = "AI" Or _
				  Trim(Session("CodigoAgente")) = "AY" Or Trim(Session("CodigoAgente")) = "ZB" Or _
				  Trim(Session("CodigoAgente")) = "AE" Or Trim(Session("CodigoAgente")) = "ZD" Or _
				  Trim(Session("CodigoAgente")) = "AM" Or Trim(Session("CodigoAgente")) = "AC" Or _
				  Trim(Session("CodigoAgente")) = "ZE" Or Trim(Session("CodigoAgente")) = "ZF" Or _
				  Trim(Session("CodigoAgente")) = "AN" Or Trim(Session("CodigoAgente")) = "AA" Or _
				  Trim(Session("CodigoAgente")) = "AV" Or Trim(Session("CodigoAgente")) = "AT" Or _
				  Trim(Session("CodigoAgente")) = "AD" Or Trim(Session("CodigoAgente")) = "AJ" or _
				  Trim(Session("CodigoAgente")) = "ZH" Or Trim(Session("CodigoAgente")) = "ZL" or _
				  Trim(Session("CodigoAgente")) = "ZK" Or Trim(Session("CodigoAgente")) = "ZG" or _
				  Trim(Session("CodigoAgente")) = "ZM" Or Trim(Session("CodigoAgente")) = "ZN" or _
				  Trim(Session("CodigoAgente")) = "ZO" Then %>

					objMenu.addchild sId, "Vender Tarjeta", "../Agente/VenderTarjeta.asp", "Principal"
			<% End If %>
			
			' JFMG 06-03-2012 se agrega venta SOAP
			<% If session("Categoria") <= 2 Then %>
			    objMenu.addchild sId, "Vender SOAP", "../soap/soap/registrarventa.aspx?Agente=<%=Session("CodigoAgente")%>", "Principal"
			<% end if %>
			' ************* FIN JFMG 06-03-2012 ******************	
				<% If session("Categoria") <= 2 Then %>
			    objMenu.addchild sId, "Recarga telefónica", "../recargatelefonica/Recargas.aspx?Agente=<%=Session("CodigoAgente")%>", "Principal"
			<% end if %>
			
		<% End If %>
		'objMenu.addchild sId, "Vender Tarjeta", "../Agente/VenderTarjeta.asp", "Principal"
		
		<% If InStr(sLista, "30") <> 0 Then %>
			sId = objmenu.addparent("Servicios")
			<% If InStr(sLista, "31") <> 0 Then %>
				'objMenu.addchild sId, "Cambio de Clave", "../Compartido/CambioClave.asp", "Principal"
			<% End If %>
			<% If InStr(sLista, "32") <> 0 Then %>
				objMenu.addchild sId, "Actualización de Datos", "../HagaseCliente.asp?Tipo=1", "Principal"
			<% End If %>
			<% If InStr(sLista, "33") <> 0 Then %>
				objMenu.addchild sId, "Alarmas", "../Compartido/ConfiguracionAlarmas.asp", "Principal"
			<% End If %>			
'			objMenu.addchild sId, "Atencion de Clientes", "ATC", "Principal"
'			objMenu.addchild sId, "GPA", "../Agente/ListaGiros.asp?Tipo=<%=afxGirosAviso%>&Pagador=<%=Session("CodigoAgente")%>", "Principal"
			objMenu.addchild sId, "Cuenta Corriente", "../Agente/CuentaCorriente.asp", "Principal"
			<% If Session("ModoPrueba") Then %>
					objMenu.addchild sId, "Cargar Archivo New", "../Agente/CargarArchivoTXT.asp", "Principal"
			<%	ElseIf Session("Categoria")=4 Then %>					
					objMenu.addchild sId, "Cargar Archivo", "../Agente/CargarArchivoTXT.asp", "Principal"
			<% End If %>
		<% End If %>
		<% If InStr(sLista, "40") <> 0 Then %>
					objMenu.addchild sId, "Giros Enviados Pendientes", "../Agente/ListaGiros.asp?Tipo=<%=session("afxTipoGirEnvPen")%>&Agente=<%=Session("CodigoAgente")%>", "Principal"
			<% End If %>
		<% If Session("CodigoAgente") = "AW" Then %>
			objMenu.addchild sId, "Compliance Compra/Venta", "../Compliance/ConsultaCompraVenta.asp", "Principal"
		<% End If %>

		<%	If Not Session("ModoPrueba") Then 
				Select Case Session("CodigoAgente") %>
		<%		Case "AM", "AE", "AG", "AP", "AH", "AO", "AT", "AR", "AI", "AN", "AY", "AC", "AA", "AD", "AV", "AQ" %>
					'window.setInterval "CargarCompliance", 7200000, "vbscript"
		<%			If request("cpl") <> "0" Then %>
					'	CargarCompliance
		<%			End If %>
		<%		End Select 
			End If
			
			if instr("JFMIRANDA;ADMIN;AAGUILAR;CMARTINEZ;NCARVAJAL;CPEÑA;AFERNANDEZ", trim(ucase(Session("NombreUsuarioOperador")))) > 0 then
		%>
				objMenu.addchild sId, "Mantenedor Paridades", "../Compartido/MantenedorParidades.asp", "Principal"
		<%
			end if			
		%>
			objMenu.addchild sId, "Imprimir Paridades", "../Compartido/ImprimirParidades.asp", "Principal"
			
		' JFMG 30-03-2011 para mostrar los tipos de cambio con agentes extranjeros
		<% If session("Categoria") < 4 Then %>
			objMenu.addchild sId, "T/C Agentes Internacionales", "../Compartido/TiposCambioAgentesInternacionales.asp", "Principal"
		<% End If %>
		' FIN JFMG 30-03-2011
				
	End Sub

	Sub CargarCompliance1
		'window.open "..\Compliance\compliance.htm", "" , "width=10, height=10"
	End Sub
	
	Sub window_onunload()
	End Sub

	Sub CargarCompliance()
		Dim rs, rsConsulta, Sql
		Dim sLinea, CodigoAgente, CodigoMoneda
				
		CodigoAgente = "<%=Session("CodigoAgente")%>"
		CodigoMoneda = "<%=Session("MonedaExtranjera")%>"
		
		Sql = "Select Distinct " & _
					  "s.fecha_solicitud, s.codigo_solicitud, " & _
					  "ds.codigo_producto, ds.codigo_moneda, " & _
					  "ds.monto_extranjera, ds.tipo_cambio, " & _
					  "ds.monto_nacional, isnull(ch.numero_cheque, '') as numero_cheque, " & _
					  "isnull(c.nombre_completo_cliente, '') as nombre_cliente, '" & _
					  CodigoAgente & "' as codigo_agente, " & _
					  "ds.tipo_operacion " & _
			   "From " & _
					  "detalle_solicitud ds " & _
			   "inner join solicitud s on ds.codigo_solicitud = s.codigo_solicitud " & _
			   "left outer join cliente c on s.codigo_cliente = c.codigo_cliente " & _
			   "left outer join cheque ch on s.codigo_solicitud = ch.codigo_solicitud " & _
			   "Where " & _
					  "s.fecha_solicitud between convert(char, '" & Date() & "', 112) and convert(char, '" & Date() & "', 112) " & _
			   "and	   isnull(s.estado_solicitud, 0) <> 0 " & _
			   "and	   ds.codigo_moneda = '" & CodigoMoneda & "' " & _
			   "and	   ds.tipo_operacion in (1, 2)"
		
		Set rsConsulta = CreateObject("AFEXSql.Sql") 
		Set rs = rsConsulta.Buscar("DSN=AFEXchange;UID=cambios;PWD=cambios;", Sql)

		If rsConsulta.ErrNumber <> 0 Then
			MsgBox rsConsulta.ErrDescription
		End If

		'Set rs = CreateObject("ADODB.Recordset") 
		'rs.CursorLocation = 3
		'rs.Open Sql, "DSN=AFEXchange;UID=cambios;PWD=cambios;", 3, 1
	
		If Err.Number <> 0 Then
			MsgBox Err.Description
		End If
		
		Do Until rs.eof
			sLinea = sLinea &  rs("fecha_solicitud") & ";" & rs("codigo_solicitud") & ";" & _
							   rs("codigo_producto") & ";" & rs("codigo_moneda") & ";" & _
							   rs("monto_extranjera") & ";" & rs("tipo_cambio") & ";" & _
							   rs("monto_nacional") & ";" & rs("numero_cheque") & ";" & _
							   rs("nombre_cliente") & ";" & CodigoAgente & ";" & _
							   rs("tipo_operacion") & vbCrLf
							   
			rs.movenext
		Loop
	
		Set rs = Nothing
	
		frm.contenido.value = sLinea

		frm.action = "http:../compliance/compliance.asp"
		frm.submit
		frm.action=""		
	End Sub
	
//-->
</script>

<body bgcolor="#336699" text="steelblue" scroll="yes">	

<form id="frm" method="post">
	<input type="hidden" name="contenido">
</form>

<!--<%	Select Case Session("CodigoAgente") %>
<%	Case "AM", "AE", "AG", "AP", "AH", "AO", "AT", "AR", "AI", "AN", "AY", "AC", "AA", "AD", "AV", "AQ"  %>
		<OBJECT id=rsConsulta style="LEFT: 0px; TOP: 0px" 
codeBase=AfexSql.CAB#version=1,0,0,0 
classid=CLSID:8952722C-FFEE-11D7-AF27-00E04C9B1440 VIEWASTEXT><PARAM NAME="_ExtentX" VALUE="0"><PARAM NAME="_ExtentY" VALUE="0"></OBJECT>
<% End Select %>-->

<table>
<tr><td>
	<IMG id=IMG1 style="LEFT: 87px; WIDTH: 100px; POSITION: absolute; TOP: -6px; HEIGHT: 102px" height     =102 hspace=0 src="../images/BordeMenuCliente.jpg" width=235 useMap="" border=0 > <!-- <embed SRC="../images/Logo.swf" STYLE="HEIGHT: 170px; LEFT: -20px; TOP: -20px; WIDTH: 170px" type="application/x-shockwave-flash">&nbsp;--><!--<img SRC="../images/AFEX.jpg" STYLE="LEFT: 10px; TOP: 5px; position: absolute" WIDTH="163" HEIGHT="162">-->
</td></tr>
<tr><td>
      <OBJECT id=objMenu 
      style="LEFT: 0px; WIDTH: 190px; POSITION: absolute; TOP: 0px; HEIGHT: 480px" 
      type=text/x-scriptlet height=461 width=190 align=left border=0 VIEWASTEXT><PARAM NAME="Scrollbar" VALUE="0"><PARAM NAME="URL" VALUE="http:../ScriptLets/Menu.htm"></OBJECT>
</td></tr>
</table>



</body>
<script>

	Sub objmenu_OnScriptletEvent(strEventName, varEventData)

	   Select Case strEventName
	   
			Case "linkClick"
				
				If Right(varEventData, 3) = "ATC" Then
					If Trim("<%=Session("ATCAFEXchange")%>") <> "" Then
						window.open "../Agente/AtencionClientes.asp?Accion=<%=afxAccionBuscar%>&Campo=<%=afxCampoCodigoExchange%>&Argumento=<%=Session("ATCAFEXchange")%>", "Principal"
					Else
						window.open "../Agente/AtencionClientes.asp?Accion=<%=afxAccionBuscar%>&Campo=<%=afxCampoCodigoExpress%>&Argumento=<%=Session("ATCAFEXpress")%>", "Principal"
					End If
				
				ElseIf Right(ucase(varEventData), 7) = "PRECIOS" Then
					'window.open "http://Aplicaciones:83/Precios/PreciosSucursales/PreciosSector.aspx", "_target"
                  'INTERNO 12550 LR 27-02-2018
                   window.open  "<%=Session("PantallaPrecios")%>", "_target"
				
				ElseIf Right(ucase(varEventData), 5) = "VISOR" Then
					window.open "../VisorAP/_visorap.asp?vpr2=1", "VisorAP", "height=224,width=170,top=250, left=350,status=no,toolbar=no,menubar=no,resize=yes,location=no,channelmode=no"
				
				' JFMG 08-08-2008 ingreso de pago amex
				ElseIf Right(varEventData, 15) = "ListadoPagoAMEX" Then					
					window.open "../Agente/ListaPagosAMEX.asp", "Principal"	
				ElseIf Right(varEventData, 8) = "PagoAMEX" Then					
					window.open "../Agente/AgregarPagoAMEX.asp", "Principal"				
				' ************************************ Fin *****************************************************
				
				' JFMG 20-07-2009 lista de tarjetas no consideradas
				ElseIf Right(varEventData, 22) = "ListaAMEXNoConsiderada" Then
					window.open "../Agente/ListaTarjetaAMEXNoConsiderada.asp", "Principal"				
					
				' ******* FIN JFMG 20-07-2009 *****************				


				' JFMG 03-12-2008 ingreso de voucher de giros en pesos en AFEXchange
				ElseIf Right(varEventData, 15) = "VoucherGiroPeso" Then
					window.open "../Agente/VoucherGirosPesos.asp", "Principal"
				' ************************************ Fin *****************************************************				
				
				Else
					window.open varEventData, "Principal"
				End If
								
		End Select
		
	End Sub
	
	
</script>
</html>
