<%@ Language=VBScript %>
<!--#INCLUDE virtual="/compartido/Constantes.asp" -->
<!--#INCLUDE virtual="/compartido/Rutinas.asp" -->
<!--#INCLUDE virtual="/compartido/Errores.asp" -->
<%
	Dim sURL, nTipo, sErrorUsuario
	
	sURL = ""
	sErrorUsuario = ""
	
	Sub ValidarCliente()
		Dim afxCliente
		Dim rs, sString
		Dim sUsuario, sPassword
		Dim sCodigo, sCodigoSistemas, nCampo
		
		On Error Resume Next
		sUsuario = Request.Form("txtUsuario")
		sPassword = Request.Form("txtPassword")		
		If sUsuario = "" Then Exit Sub
		
		If sUsuario = "aruiz" And sPassWord = "1234" Then
			Session("ModoPrueba") = True
			'sUsuario = "lax"
			'sPassword = "123456789"
			CambiarODBC
		End If
		If sUsuario = "lax" And sPassWord = "1234" Then
			Session("ModoPrueba") = True
			sUsuario = "lax"
			sPassword = "123456789"
			CambiarODBC
		End If
		If sUsuario = "moneda" And sPassWord = "1234" Then
			Session("ModoPrueba") = True
			sUsuario = "afexmoneda"
			sPassword = "1235"
			CambiarODBC
		End If
		If sUsuario = "valpo" And sPassWord = "1234" Then
			Session("ModoPrueba") = True
			sUsuario = "afexvalpo"
			sPassword = "123456789"
			CambiarODBC
		End If
		If sUsuario = "vina" And sPassWord = "1234" Then
			Session("ModoPrueba") = True
			sUsuario = "afexviña"
			sPassword = "123456789"
			CambiarODBC
		End If
		If sUsuario = "ibero" And sPassWord = "4567" Then
			Session("ModoPrueba") = True
			sUsuario = "ibero"
			sPassword = "123456789"
			CambiarODBC
		End If
		If sUsuario = "mx" And sPassWord = "1234" Then
			Session("ModoPrueba") = True
			sUsuario = "multiexpress"
			sPassword = "123456789"
			CambiarODBC
		End If
		If sUsuario = "moon" And sPassWord = "1234" Then
			Session("ModoPrueba") = True
			sUsuario = "mvalley"
			sPassword = "123456789"
			CambiarODBC
		End If		
		If sUsuario = "mesa" And sPassWord = "1234" Then
			Session("ModoPrueba") = True
			sUsuario = "afexmesa"
			sPassword = "123456789"
			CambiarODBC
		End If		
		If sUsuario = "ahumada" And sPassWord = "1234" Then
			Session("ModoPrueba") = True
			sUsuario = "cahumada"
			sPassword = "123456789"
			CambiarODBC
		End If		
		If sUsuario = "moneybroker" And sPassWord = "1234" Then
			Session("ModoPrueba") = True
			'sUsuario = "mvalley"
			sPassword = "123456789"
			CambiarODBC
		End If		
		If sUsuario = "titan" And sPassWord = "1234" Then
			Session("ModoPrueba") = True
			sUsuario = "titanpay"
			sPassword = "123456789"
			CambiarODBC
		End If		
		
		Set afxCliente = Server.CreateObject("AfexCorporativo.Cliente")
		Set rs = Server.CreateObject("ADODB.recordset")
		Set rs = afxCliente.ObtenerCliente(Session("afxCnxCorporativa"), sUsuario, sPassword)
		If Err.Number <> 0 Then			
			sErrorUsuario = Err.description 
			Set rs = Nothing
			Set afxCliente = Nothing
			'MostrarErrorMS afxCliente, "Obtener Cliente 1"
			Exit Sub
		End If
		
		If afxCliente.ErrNumber <> 0 Then			
			sErrorUsuario = replace(afxCliente.ErrDescription, vbCrLf , "")
			Set rs = Nothing
			'MostrarErrorAFEX afxCliente, "Obtener Cliente 2"
			Exit Sub
		End If
		
		If rs("estado") = 0 Then			
			sErrorUsuario = "Usted aún no está habilitado para operar con nuestros servicios AFEXweb. Comuníquese con nuestro departamento de Atención al Cliente al teléfono (562) 6369052 o en nuestra zona de «Contáctenos» de esta página o al correo «atencion.clientes@afex.cl»"
			Set rs = Nothing
			Exit Sub
		End If
		
		sCodigo = rs("codigo")
		nTipo = rs("tipo")		
		
		Session("CodigoCliente") = rs("codigo")
		Session("NombreCliente") = rs("nombre")
		If Not IsNull(rs("codigo_agente")) Then
			Session("CodigoAgente") = uCase(Trim(rs("codigo_agente")))
			Session("Categoria") = rs("categoria_agente")
			Session("CiudadCliente") = Trim(UCase(rs("ciudad_agente")))
			Session("PaisCliente") = Trim(UCase(rs("pais_agente")))
		Else
			Session("Categoria") = 0
			Session("CiudadCliente") = Trim(UCase(rs("codigo_ciudad")))
			Session("PaisCliente") = Trim(UCase(rs("codigo_pais")))
		End If		
		Session("AFEXpress") = EvaluarVar(rs("codigo_afexpress"), "")
		Session("AFEXchange") = EvaluarVar(rs("codigo_afexchange"), "")
		Session("CodigoCaja") = EvaluarVar(rs("codigo_caja"), "")
		Session("FechaApertura") = "01-01-2003"
		Session("NombreUsuario") = EvaluarVar(rs("nombre_usuario"), "")
		Session("NombreUsuarioOperador") = EvaluarVar(rs("nombre_usuario"), "")
		Session("NombreOperador") = rs("Nombre")
		Session("RegionAgente") = rs("region")
		If Not IsNull(rs("email")) Then 
			Session("emailCliente") = rs("email")
		Else
			Session("emailCliente") = ""
		End If
		
		Set afxCliente = Nothing
		Set rs = Nothing
		
		'response.Redirect "compartido/error.asp?description=" & Session("CodigoCliente") & ", " & Session("NombreCiente")		
		Select Case nTipo
		Case 4 
			sURL = "Agente/Default.asp?Codigo=" & sCodigo
			'response.Redirect sURL
			
		Case 5
			sURL = "Agente/Default.asp?Codigo=" & sCodigo
			'response.Redirect sURL
		
		Case 98
			'Session("CodigoAgente") = "WB"
			sURL = "Sucursal/Default.asp?Codigo=" & sCodigo

		Case 99
			Session("CodigoAgente") = "WB"
			sURL = "Sucursal/Default.asp?Codigo=" & sCodigo

		Case Else			
			Session("CodigoAgente") = "WB"
			sURL = "Cliente/Default.asp?Codigo=" & sCodigo
			'Response.Redirect sURL
		End Select
	End Sub

	Sub CambiarODBC()
		'Session("afxCnxCorporativa") = "Provider=SQLOLEDB;Password=afxsqlcor;User ID=corporativa;Initial Catalog=corporativa;Data Source=cipres;"
		'Session("afxCnxAFEXchange") = "Provider=SQLOLEDB;Password=afxsql*cip;User ID=sa;Initial Catalog=cambios;Data Source=cipres;"
		'Session("afxCnxAFEXpress") = "Provider=SQLOLEDB;Password=giros;User ID=giros;Initial Catalog=giros;Data Source=cipres;"
		'Session("afxCnxAFEXweb") = "Provider=SQLOLEDB;Password=cambios;User ID=cambios;Initial Catalog=cambios;Data Source=cipres;"
		
 		Session("afxCnxCorporativa") = "DSN=wAfexCorporativa;UID=corporativa;PWD=afxsqlcor;"
		Session("afxCnxAFEXchange") = "DSN=wAFEXchange;UID=cambios;PWD=cambios;"
		Session("afxCnxAFEXpress") = "DSN=wAFEX_giros;UID=giros;PWD=giros;"	
		Session("afxCnxAFEXweb") = "DSN=wAFEXweb;UID=cambios;PWD=cambios;"	
	End Sub	
	
%>

<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<title>AFEX</title>
<link rel="stylesheet" type="text/css" href="Estilos/MenuPrincipal.CSS">
</head>

<script LANGUAGE="VBScript">
<!--
	Sub ValidaCliente_onClick()
		'window.navigate  "ValidarCliente.asp?NombreUsuario=" & txtUsuario.value & "&Password=" & txtPassword.value
	End Sub
	
	Sub window_onLoad()
		Dim sId
		
		objmenu.bgColor = document.bgColor 
		objmenu.stylesheet = "../Estilos/MenuPrincipal.css"
		objmenu.addpassword "Sucursales/Agentes"
		sId = objmenu.addparent("Información")
		objMenu.addchild sId, "Quienes Somos", "../QuienesSomos.asp", "Principal"
		objMenu.addchild sId, "Direcciones en Chile", "../Sucursales.asp?Region=13", "Principal"
		objMenu.addchild sId, "Cobertura Internacional", "../CoberturaInternacional.asp", "Principal"		
		'objMenu.addchild sId, "Tipos de Cambio", "../Compartido/ConsultaTipoCambio.asp?Tipo=<%=afxPrincipal%>", "Principal"
		sId = objmenu.addparent("Simulaciones")
		objMenu.addchild sId, "Enviar un Giro", "../SimulacionGiro.asp", "Principal"
		objMenu.addchild sId, "Enviar una Transferencia", "../SimulacionTransfer.asp", "Principal"
		<%
		If cInt(0 & request("Accion")) = 1 Then
			ValidarCliente
		End If
		%>
		<%If sURL <> "" Then%>
			<% If Session("Categoria") = 3 Then %>
					'window.showModalDialog "Usuario.asp?tp=<%=nTipo%>", , "center: yes"
					window.open "Usuario.asp?tp=<%=nTipo%>", "", "width=300; height=165; menu=no; top=220; left=250"
			<% ElseIf nTipo = 4 Then %>				
					'window.showModalDialog "Usuario.asp?tp=<%=nTipo%>", , "center: yes"
					window.open "Usuario.asp?tp=<%=nTipo%>", "", "width=300; height=165; menu=no; top=220; left=250"
			<% ElseIf nTipo = 98 Then %>				
					'window.showModalDialog "Usuario.asp?tp=<%=nTipo%>", , "center: yes"
					window.open "Usuario.asp?tp=<%=nTipo%>", "", "width=300; height=165; menu=no; top=220; left=250"
			<% ElseIf nTipo = 99 Then %>				
					'window.showModalDialog "Usuario.asp?tp=<%=nTipo%>", , "center: yes"
					window.open "Usuario.asp?tp=<%=nTipo%>", "", "width=300; height=165; menu=no; top=220; left=250"
			<% Else %>			
					window.open "<%=sURL%>"
			<% End If %>
			
		<% End If %>
		<% If sErrorUsuario <> "" Then %>
			msgbox "<%=sErrorUsuario%>" ,, "AFEX"
		<% End If %>
	End Sub
	
//-->
</script>
<body onKeyPress="KeyEnter" id="bBody" sstyle="background-color: white" bgcolor="#2F4F4F" s="lightseagreen" text="#86d4cd" scroll="no">	
<table>
<tr><td>
	<img border="0" height="102" hspace="0" id="IMG1" src="images/Borde%20Menu.jpg" style="HEIGHT: 102px; LEFT: 87px; POSITION: absolute; TOP: -6px; WIDTH: 235px" useMap width="235">
<!--
	<embed SRC="images/Logo.swf" STYLE="HEIGHT: 170px; LEFT: -20px; TOP: -20px; WIDTH: 170px" type="application/x-shockwave-flash">&nbsp; 
-->
	<img src="images/logo%20afex.jpg" style="position: absolute; top: 10; left:20" WIDTH="136" HEIGHT="139">
</td></tr>
<tr><td>
      <object align="left" height="245" id="objMenu" style="HEIGHT: 245px; LEFT: -14px; POSITION: relative; TOP: 140px; WIDTH: 190px" type="text/x-scriptlet" width="174" border="0" VIEWASTEXT><param NAME="Scrollbar" VALUE="0"><param NAME="URL" VALUE="http:Scriptlets/Menu.htm"></object>
</td></tr>
</table>
<form id="frmCliente" method="post">
	<input type="hidden" name="txtUsuario" value>
	<input type="hidden" name="txtPassword" value>
</form>
<!--<td>	<a href="PrincipalEnglish/Menu.asp">Versión English</a></td>-->
</body>
<script>

	Sub objmenu_OnScriptletEvent(strEventName, varEventData)
	   Select Case strEventName
	   
			Case "linkClick"
				window.open varEventData, "Principal"

			Case "Clave"
				frmCliente.txtPassword.value = objmenu.Password
				frmCliente.txtUsuario.value = objmenu.Usuario
				frmCliente.action = "Menu.asp?Accion=1"
				frmCliente.submit 
				frmCliente.action = ""
				
		End Select
		
	End Sub
					

</script>
</html>
