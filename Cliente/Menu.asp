<%@ Language=VBScript %>
<!--#INCLUDE virtual="/cliente/Constantes.asp" -->
<%
	Dim sLista
	
	sLista = Request("Menu")
	If sLista = "*" Then
		sLista = "10;11;12;13;14;15;16;20;21;22;23;30;31;32;33"
	End If
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
			<% If InStr(sLista, "12") <> 0 Then %>
				objMenu.addchild sId, "Giros Pendientes de Pago", "../Agente/ListaGiros.asp?Tipo=<%=afxGirosPendientes%>&Cliente=<%=Session("CodigoCliente")%>", "Principal"
			<% End If %>
			<% If InStr(sLista, "11") <> 0 Or InStr(sLista, "12") <> 0  Then %>
					objMenu.addchild sId, "Cartola en línea", "../Agente/ListaGiros.asp?Tipo=<%=afxGirosCartola%>&Cliente=<%=Session("AFEXpress")%>&TipoLlamada=<%=afxCliente%>", "Principal"
			<% End If %>
			<% If InStr(sLista, "11") <> 0 Then %>
					objMenu.addchild sId, "Giros Enviados", "../Agente/ListaGiros.asp?Tipo=<%=afxGirosEnviados%>&Cliente=<%=Session("AFEXpress")%>&TipoLlamada=<%=afxCliente%>", "Principal"
			<% End If %>
			<% If InStr(sLista, "12") <> 0 Then %>
				'msgbox  tipo & "=<%=afxGirosRecibidos%>, " & cliente & "=<%=Session("AFEXpress")%>, " & TipoLlamada & "=<%=afxCliente%>"
				objMenu.addchild sId, "Giros Recibidos", "../Agente/ListaGiros.asp?Tipo=<%=afxGirosRecibidos%>&Cliente=<%=Session("AFEXpress")%>&TipoLlamada=<%=afxCliente%>", "Principal"
			<% End If %>
			<% If InStr(sLista, "13") <> 0 Then %>
				objMenu.addchild sId, "Transferencias Enviadas", "../Agente/ListaTransfer.asp?Tipo=<%=afxTrfEnviadas%>&Titulo=Transferencias&Cliente=<%=Session("AFEXchange")%>", "Principal"
				'objMenu.addchild sId, "Ultimas 15 Transferencias", "../Compartido/UltimasTransferencias.asp?Tipo=1&CodigoCliente=<%=Session("AFEXchange")%>&NombreCliente=<%=Session("NombreCliente")%>&Fondo=Transfers", "Principal"
			<% End If %>
			<% If InStr(sLista, "14") <> 0 Then %>
				'objMenu.addchild sId, "Compras y Ventas", "../Cliente/ListaGiros.asp?Tipo=5&amp;Titulo=Compras y Ventas", "Principal"
			<% End If %>
			<% If InStr(sLista, "15") <> 0 Then %>
				'objMenu.addchild sId, "Ultimas Operaciones", "../Cliente/ListaGiros.asp?Tipo=6&amp;Titulo=Ultimas Operaciones", "Principal"
			<% End If %>
				objMenu.addchild sId, "Tipos de Cambio", "../Compartido/ConsultaTipoCambio.asp", "Principal"
				objMenu.addchild sId, "Direcciones en Chile", "../Sucursales.asp?Region=13", "Principal"
		<% End If %>		
		<% If InStr(sLista, "20") <> 0 Then %>
			sId = objmenu.addparent("Transacciones")
			<% If InStr(sLista, "21") <> 0 Then %>
				'objMenu.addchild sId, "Enviar un Giro", "../Cliente/EnviarGiro.asp", "Principal"
			<% End If %>
			<% If InStr(sLista, "22") <> 0 Then %>
				'objMenu.addchild sId, "Enviar una Transferencia", "../Cliente/EnviarTransfer.asp", "Principal"
			<% End If %>
			<% If InStr(sLista, "23") <> 0 Then %>
				'objMenu.addchild sId, "Compra/Venta de Monedas", "../Compartido/CompraVentaMoneda.asp", "Principal"
			<% End If %>
		<% End If %>
	
		<% If InStr(sLista, "30") <> 0 Then %>
			sId = objmenu.addparent("Servicios")
			<% If InStr(sLista, "31") <> 0 Then %>
				objMenu.addchild sId, "Cambio de Clave", "../Compartido/CambioClave.asp", "Principal"
			<% End If %>
			<% If InStr(sLista, "32") <> 0 Then %>
				'objMenu.addchild sId, "Actualización de Datos", "../HagaseCliente.asp?Tipo=1", "Principal"
			<% End If %>
			<% If InStr(sLista, "33") <> 0 Then %>
				'objMenu.addchild sId, "Alarmas", "../Compartido/ConfiguracionAlarmas.asp", "Principal"
			<% End If %>
		<% End If %>
	End Sub
	
//-->
</script>

<body bgcolor="#336699" text="steelblue" scroll="no">	
<table>
<tr><td>
	<img border="0" height="102" hspace="0" id="IMG1" src="../images/BordeMenuCliente.jpg" style="HEIGHT: 102px; LEFT: 87px; POSITION: absolute; TOP: -6px; WIDTH: 235px" useMap width="235">
	<embed SRC="../images/Logo.swf" STYLE="HEIGHT: 170px; LEFT: -20px; TOP: -20px; WIDTH: 170px" type="application/x-shockwave-flash">&nbsp;
</td></tr>
<tr><td>
      <object align="left" height="362" id="objMenu" style="HEIGHT: 362px; LEFT: -10px; POSITION: relative; TOP: 0px; WIDTH: 190px" type="text/x-scriptlet" width="174" border="0" VIEWASTEXT><param NAME="Scrollbar" VALUE="0"><param NAME="URL" VALUE="http:../ScriptLets/Menu.htm"></object>
</td></tr>
</table>
</body>
<script>

	Sub objmenu_OnScriptletEvent(strEventName, varEventData)

	   Select Case strEventName
	   
			Case "linkClick"
				window.open varEventData, "Principal"
				
		End Select
		
	End Sub
	
	
</script>
</html>
