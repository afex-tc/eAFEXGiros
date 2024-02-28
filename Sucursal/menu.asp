<%@ Language=VBScript %>
<!--#INCLUDE virtual="/Compartido/Constantes.asp" -->
<!--#INCLUDE virtual="/Compartido/Errores.asp" -->
<%
	Dim sLista, bExtranjero
	   
	If Session("PaisCliente") <> Session("PaisMatriz") Then		
		bExtranjero = True
	Else
		bExtranjero = False
	End If
	
	sLista = Request("Menu")
	If sLista = "*" Then
		sLista = "10;11;12;13;14;15;16;20;21;22;23;30;31;32;33"
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
		<% If InStr(sLista, "10") <> 0  Then %>
				sId = objmenu.addparent("Consultas")
				<% If InStr(sLista, "11") <> 0  Then %>
						objMenu.addchild sId, "Lista Clientes", "../Sucursal/ListaClientes.asp", "Principal"
				<%End If%>
				
				' JFMG 27-11-2008
				objMenu.addchild sId, "Lista Operaciones Cumplimiento", "../Sucursal/ListaOperacionesClienteCumplimiento.asp", "Principal"
				' ***************** FIN **************************
				
				' JFMG 01-02-2011
				objMenu.addchild sId, "WorldCheck", "../Sucursal/ListaCoincidenciasWorldCheck.asp", "Principal"
				' FIN JFMG 01-02-2011
				
		<% End If %>
		<% If InStr(sLista, "20") <> 0  Then %>
				sId = objmenu.addparent("Servicios")
				<% If InStr(sLista, "21") <> 0  Then %>		
						objMenu.addchild sId, "Agregar Nuevo Cliente", "../Sucursal/NuevoCliente.asp", "Principal"
				<% End If %>
				<% If InStr(sLista, "22") <> 0  Then %>		
						objMenu.addchild sId, "Configuracion Ficha Cliente", "../Sucursal/ConfiguracionFicha.asp", "Principal"
						
						'Jonathan Miranda G. 03-04-2007
						objMenu.addchild sId, "Configuracion Ficha Cliente Giro", "../Sucursal/ConfiguracionFichaGiro.asp", "Principal"
						'------------ Fin ------------------------
				<% End If %>
		<% End If %>
		
	End Sub

	Sub CargarCompliance1
		'window.open "..\Compliance\compliance.htm", "" , "width=10, height=10"
	End Sub
	
	Sub window_onunload()
	End Sub
	
//-->
</script>

<body bgcolor="#336699" text="steelblue" scroll="no">	




<form id="frm" method="post">
	<input type="hidden" name="contenido">
</form>

<%	Select Case Session("CodigoAgente") %>
<%	Case "AM", "AE", "AG", "AP", "AH", "AO", "AT", "AR", "AI", "AN", "AY", "AC", "AA", "AD", "AV", "AQ"  %>
		<OBJECT classid="CLSID:8952722C-FFEE-11D7-AF27-00E04C9B1440"  codebase="AfexSql.CAB#version=1,0,0,0" 
		id=rsConsulta style="LEFT: 0px; TOP: 0px;" VIEWASTEXT>
		<PARAM NAME="_ExtentX" VALUE="1">
		<PARAM NAME="_ExtentY" VALUE="1"></OBJECT>
<% End Select %>

<table>
<tr><td>
	<img border="0" height="102" hspace="0" id="IMG1" src="../images/BordeMenuCliente.jpg" style="HEIGHT: 102px; LEFT: 87px; POSITION: absolute; TOP: -6px; WIDTH: 235px" useMap width="235">
	<!-- <embed SRC="../images/Logo.swf" STYLE="HEIGHT: 170px; LEFT: -20px; TOP: -20px; WIDTH: 170px" type="application/x-shockwave-flash">&nbsp;-->
	<img SRC="../images/AFEX.jpg" STYLE="LEFT: 10px; TOP: 10px; position: absolute" WIDTH="163" HEIGHT="162">
</td></tr>
<tr><td>
      <object align="left" height="362" id="objMenu" style="HEIGHT: 362px; LEFT: 0px; POSITION: absolute; TOP: 200px; WIDTH: 190px" type="text/x-scriptlet" width="174" border="0" VIEWASTEXT><param NAME="Scrollbar" VALUE="0"><param NAME="URL" VALUE="http:../ScriptLets/Menu.htm"></object>
</td></tr>
</table>
</body>
<script LANGUAGE="VBScript">

	Sub objmenu_OnScriptletEvent(strEventName, varEventData)

	   Select Case strEventName
	   
			Case "linkClick"
				window.open varEventData, "Principal"
												
		End Select
		
	End Sub
	
	
</script>
</html>
