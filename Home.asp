<%@ Language=VBScript %>

<!--#INCLUDE virtual="/Compartido/Errores.asp" -->
<!--#INCLUDE virtual="/Compartido/Rutinas.asp" -->

<%
	Dim sURL, nTipo, sErrorUsuario
		
	sURL = ""
	sErrorUsuario = ""
	
	Public Function LimpiarCampoTxt(ByVal Cadena)
	        LimpiarCampoTxt = Cadena
	        LimpiarCampoTxt = Replace(LimpiarCampoTxt,"select","")
	        LimpiarCampoTxt = Replace(LimpiarCampoTxt,"drop","")
	        LimpiarCampoTxt = Replace(LimpiarCampoTxt,"delete","")
	        LimpiarCampoTxt = Replace(LimpiarCampoTxt,"update","")
	        LimpiarCampoTxt = Replace(LimpiarCampoTxt,"exec","")
	        LimpiarCampoTxt = Replace(LimpiarCampoTxt,"having","")
	        LimpiarCampoTxt = Replace(LimpiarCampoTxt,"truncate","")
	        LimpiarCampoTxt = Replace(LimpiarCampoTxt,"alter","")
	        LimpiarCampoTxt = Replace(LimpiarCampoTxt,"create","")
	        LimpiarCampoTxt = Replace(LimpiarCampoTxt,"grant","")
	        LimpiarCampoTxt = Replace(LimpiarCampoTxt,"<","")
	        LimpiarCampoTxt = Replace(LimpiarCampoTxt,">","")
	        LimpiarCampoTxt = Replace(LimpiarCampoTxt,"'","")
	        LimpiarCampoTxt = Replace(LimpiarCampoTxt,"--","")	   
	End Function
	
	Sub ValidarCliente()
		Dim afxCliente
		Dim rs, sString
		Dim sUsuario, sPassword
		Dim sCodigo, sCodigoSistemas, nCampo
		
		
		
		On Error Resume Next
				
		sUsuario = LimpiarCampoTxt(Request.Form("txtUsuario"))
		sPassword = LimpiarCampoTxt(Request.Form("txtPassword"))
			
	
	
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
		
		
		'****** Jonathan 11-07-2006 *****		
		bGiros = True
		If bGiros Then
		'******* fin ******
			
		Set afxCliente = Server.CreateObject("AfexCorporativo.Cliente")

		If Err.Number <> 0 Then		
			sErrorUsuario = "error 1." & Err.description 
			Set rs = Nothing
			Set afxCliente = Nothing			
			'MostrarErrorMS afxCliente, "Obtener Cliente 1"
			Exit Sub
		End If


		Set rs = Server.CreateObject("ADODB.recordset")

		If Err.Number <> 0 Then		
			sErrorUsuario = "error 2." & Err.description 
			Set rs = Nothing
			Set afxCliente = Nothing			
			'MostrarErrorMS afxCliente, "Obtener Cliente 1"
			Exit Sub
		End If

		Set rs = afxCliente.ObtenerCliente(Session("afxCnxCorporativa"), sUsuario, sPassword)	
		
		If Err.Number <> 0 Then		
			sErrorUsuario = "error 3." & Err.description 
			Set rs = Nothing
			Set afxCliente = Nothing			
			'MostrarErrorMS afxCliente, "Obtener Cliente 1"
			Exit Sub
		End If


		If Err.Number <> 0 Then		
			sErrorUsuario = "error 4." & Err.description 
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
		
		'****** Jonathan 11-07-2006 ******	
		Else		
			Set rs = EjecutarSQLCliente(Session("afxCnxCorporativa"), " execute buscarclienteglobal '" & sUsuario & "', '" & sPassword & "' ")
			
			If Err.Number <> 0 Then		
				sErrorUsuario = sErrorUsuario & ucase(" execute buscarclienteglobal '" & sUsuario & "', '" & sPassword & "' ") & Err.description & " Error.01"
				Set rs = Nothing				
				Exit Sub
			End If			
		End If
		'****** Fin ******

		If rs("estado") = 0 Then			
			sErrorUsuario = "Usted aún no está habilitado para operar con nuestros servicios AFEXweb. Comuníquese con nuestro departamento de Atención al Cliente al teléfono (562) 6369052 o en nuestra zona de «Contáctenos» de esta página o al correo «atencion.clientes@afex.cl»"
			Set rs = Nothing
			Exit Sub
		End If		
		
		sCodigo = rs("codigo")
		nTipo = rs("tipo")		
	
		' JFMG 04-12-2009 se guardan tambien las contraseñas para luego enviarlas a AFEXchangeWEB
		Session("ContrasenaAgente") = sPassword
		Session("NombreUsuarioAgente") = sUsuario
		Session("EnlaceAFEXChangeWeb") = rs("enlaceafexchangeweb")
		' ********** FIN 04-12-2009 **************
	
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
		
		' JFMG 03-07-2010 consulta a la BBDD para ver si el agente puede adjuntar imagenes a clientes corporativos
		on error resume next
		dim iAdjuntarImagenClienteCorporativo, sMensajeError
		sSQL = " exec MostrarAdjuntarImagenClienteCorporativoAgente " & evaluarstr(Session("CodigoAgente"))
		set rs = ejecutarsqlcliente(Session("afxCnxAFEXpress"), sSQL)
		if err.number <> o then
			sMensajeError = "Error al verificar Adjuntar Imagen Cliente Corporativo. " & err.Description 
			iAdjuntarImagenClienteCorporativo = 0
		else
			iAdjuntarImagenClienteCorporativo = rs("AdjuntarImagenClienteCorporativo")
		end if
		Session("VerClienteCorporativo") = iAdjuntarImagenClienteCorporativo
		set rs = nothing
		' FIN JFMG  03-07-2010
		
		'response.Redirect "compartido/error.asp?description=" & Session("CodigoCliente") & ", " & Session("NombreCiente")		
		Select Case nTipo
		Case 4 
			sURL = "../Agente/Default.asp?Codigo=" & sCodigo
			'response.Redirect sURL
			
		Case 5
			sURL = "../Agente/Default.asp?Codigo=" & sCodigo
			'response.Redirect sURL
		
		Case 98
			'Session("CodigoAgente") = "WB"
			sURL = "../Sucursal/Default.asp?Codigo=" & sCodigo

		Case 99
			Session("CodigoAgente") = "WB"
			sURL = "../Sucursal/Default.asp?Codigo=" & sCodigo

		Case Else			
			Session("CodigoAgente") = "AF"
			sURL = "EnLineaAfex.asp"
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

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>

<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />

<title>.:. AFEX .:.</title>
<script language="JavaScript" type="text/javascript">
function click(){
if(event.button==2){
    alert('No esta permitida esta acción sobre el sitio');
}
}
document.onmousedown = click
function checartecla (evt) 
{if (evt.keyCode == 17) {
    alert('No esta permitida esta acción sobre el sitio');
} 
} 

//-->
</script>

<link href="CSS/Css_Home.css" rel="stylesheet" type="text/css" />
<link href="CSS/Css_Form.css" rel="stylesheet" type="text/css" />
<link href="CSS/Css_Capamedio.css" rel="stylesheet" type="text/css" />

<link href="CSS/Links_T1T2T3.css" rel="stylesheet" type="text/css" />
<link href="CSS/linkcss_3.css" rel="stylesheet" type="text/css" />
<link href="CSS/linkcss_2.css" rel="stylesheet" type="text/css" />
<script src="Scripts/AC_RunActiveContent.js" type="text/javascript"></script>
</head>

<script language="vbscript">
<!--

	On Error Resume Next
	
	Dim nSeg
	
	Sub window_onLoad()
		Dim sId		
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
			<% ElseIf Session("CodigoCliente") = 98 then %>
					window.open "Usuario.asp?tp=<%=nTipo%>", "", "width=300; height=165; menu=no; top=220; left=250"
			<% Else %>			
					window.navigate "<%=sURL%>"
			<% End If %>
			
		<% End If %>
		<% If sErrorUsuario <> "" Then %>
			msgbox "<%=sErrorUsuario%>" ,, "AFEX"
		<% End If %>
		
		frmCliente.txtUsuario.focus 
		
		' JFMG 03-07-2010
		if "<%=sMensajeError%>" <> "" then
			msgbox "<%=sMensajeError%>", , "AFEX"
		end if		
		' FIN JFMG 03-07-2010
		
        ' INTERNO-2912 JFMG 01-12-2014
		'JFMG 17-12-2010 muestra promoción pin telefónico
		'If "<%=Session("EstadoPromocionPinTelefonico")%>" = "1" Then
		'    window.open "popup_promocionpintelefonico.html","","top=50,left=300,width=450,height=645"
		'End If
		'FIN JFMG 17-12-2010
        ' FIN INTERNO-2912 JFMG 01-12-2014

		If "<%=Session("EstadoPromocionNavidad")%>" = "1" Then
		    window.open "popup_navidad.htm","","top=50,left=300,width=620,height=320"
		End If
		
		
	End Sub
	

	Sub ValidarUsuario()		
		If frmCliente.txtUsuario.value = Empty Or frmCliente.txtPassword.value = Empty Then exit sub		
		frmCliente.action = "Home.asp?Accion=1"
		frmCliente.submit		
	End Sub
		
	
	sub ValidarTecla()
		dim iTecla
					
		iTecla = window.event.keyCode 		
		if iTecla = 32 or iTecla = 34 or iTecla = 39 then		
			window.event.keyCode = 0		
		end if
		
	end sub
	
-->
</script>

<body   leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0" marginwidth="0" marginheight="0" onkeydown="checartecla(event)">
    
    <div style=" text-align: center; background-color: #31514A;">
    
    <div style=" padding-top: 200px; width: 800px; height: 800px; background-image: url(images/afex-fondo-web-interna.png); vertical-align: bottom; ">

        <table  border="0" align="center" cellspacing="0" bgcolor="#FFFFFF" class="CapaAfuera" >
        <tr>
            <td valign="bottom" align="center">
                <table width="200" border="0" cellspacing="0" bgcolor="#31514A" class="Css_Capa_medio">
                    <tr>
                        <td width="200" height="0" valign="top" bgcolor="#31514A" align="center">                            
                                <form name="frmCliente" method="post" action="">
                                <table width="200" border="0" cellspacing="0" bgcolor="#31514A" align="center">
                                <tr>
                                    <td width="181">
                                        <div align="left">
                                            <img src="Img/img_paginahome/ingresoclientes.jpg" width="130" height="15" /></div>
                                    </td>
                                    <td width="15" valign="top">
                                        &nbsp;
                                    </td>
                                </tr>
                                <tr bgcolor="#31514A" class="Css_Capa_medio">
                                    <td height="0" colspan="2">
                                        <div align="right">
                                            <div align="left">
                                                &nbsp;<span class="Estilo7" style="color: White; font-family: Tahoma; font-size: 12px;">Usuario:</span>&nbsp;&nbsp;&nbsp;</div>
                                        </div>
                                    </td>
                                </tr>
                                <tr bgcolor="#31514A" class="Css_Capa_medio">
                                    <td height="0" colspan="2" bgcolor="#31514A">
                                        <div align="left">
                                            <input name="txtUsuario" type="text" onkeypress="ValidarTecla()" class="textboxHome"
                                                size="28" />
                                            &nbsp;&nbsp;
                                        </div>
                                    </td>
                                </tr>
                                <tr bgcolor="#31514A" class="Css_Capa_medio">
                                    <td height="0" colspan="2">
                                        <div align="right">
                                            <div align="left">
                                                &nbsp;<span class="Estilo7" style="color: White; font-family: Tahoma; font-size: 12px;">Contrase&ntilde;a:</span>&nbsp;&nbsp;&nbsp;</div>
                                        </div>
                                    </td>
                                </tr>
                                <tr bgcolor="#31514A" class="Css_Capa_medio">
                                    <td height="0" colspan="2">
                                        <div align="left">
                                            <input name="txtPassword" type="password" onkeypress="ValidarTecla()" class="textboxHome"
                                                size="28" />
                                            &nbsp;&nbsp;
                                        </div>
                                    </td>
                                </tr>
                                <tr bgcolor="#31514A" class="Css_Capa_medio">
                                    <td height="0" colspan="2">
                                        <div align="right">
                                            
                                                <img onclick="ValidarUsuario()" name="Home_r7_c41" src="Img/img_paginahome/Home_r7_c4.jpg"
                                                    width="90" height="22" border="0" id="Home_r7_c41" alt="Boton Entrar02" style="cursor: hand" />
                                            
                                        </div>
                                    </td>
                                </tr>
                                </table>
                                </form>
                            
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
    
    
    </div>
    
    
    </div>
</body>
</html>
