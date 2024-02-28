<%@ Language=VBScript %>
<%
	Response.Expires = 0
	If Session("CodigoCliente") = "" Then
		response.Redirect "../Compartido/TimeOut.htm"
		response.end
	End If
%>
<%
	
	On Error Resume Next

	Dim nCodigoCliente, sNombreCliente, sRutCliente
	Dim sTitulo
		
	nCodigoCliente = Request("cc")
	sNombreCliente = Request("nc")
	sRutCliente = Request("rt")
	
	sTitulo = sNombreCliente
	If Trim(sTitulo) = Empty Then
		sTitulo = "Agregar Historia"
	End If
%>

<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<title>Agregar una Nueva Historia</title>
<link rel="stylesheet" type="text/css" href="../Estilos/Cliente.css">
</head>

<script LANGUAGE="VBScript">
<!--
	Const sEncabezadoFondo = "Servicios"
	Const sEncabezadoTitulo = "Agregar Historia"
	Const sClass = "TituloPrincipal"

	Sub imgAceptar_onClick()
		
		' JFMG 16-04-2009 se agregan opciones de autorizacion
		if not frmHistoria.chkTipoAutorizacion.disabled  then			 
			VerificarTipoAutorizacion
			frmHistoria.cbxOpcionAutorizacion.disabled = false
		end if
		' ******* FIN ****************
	
		frmHistoria.action = "GuardarNuevaHistoria.asp?cc=<%=nCodigoCliente%>&nc=<%=sNombreCliente%>&rt=<%=sRutCliente%>"
		frmHistoria.submit
		frmHistoria.action = ""
	End Sub	
	
	
	' JFMG 16-04-2009 se agregan opciones de autorización
	sub chkTipoAutorizacion_onClick()
		
		if frmHistoria.chkTipoAutorizacion.checked then
			frmHistoria.cbxOpcionAutorizacion.disabled = false			
			'frmHistoria.cbxOpcionAutorizacion.focus 
			
		else
			frmHistoria.cbxOpcionAutorizacion.value = ""
			frmHistoria.cbxOpcionAutorizacion.disabled = true
		end if
	end sub
	
	sub cbxOpcionAutorizacion_onChange()
		frmHistoria.txtDescripcion.value = ""
		frmHistoria.cbxValorOpcionAutorizacion.value  = frmHistoria.cbxOpcionAutorizacion.value
		frmHistoria.txtDescripcion.value = frmHistoria.cbxValorOpcionAutorizacion.options(frmHistoria.cbxValorOpcionAutorizacion.selectedindex).text
		frmHistoria.hddTextoAutorizacion.value = frmHistoria.cbxValorOpcionAutorizacion.options(frmHistoria.cbxValorOpcionAutorizacion.selectedindex).text		
	end sub
	
	
	sub cbxOpcionAutorizacion_onBlur()		
		VerificarTipoAutorizacion		
	end sub
	
	sub VerificarTipoAutorizacion()	
		if cint("0" & frmHistoria.cbxOpcionAutorizacion.value) = 0 then			
			frmHistoria.chkTipoAutorizacion.checked = false
			chkTipoAutorizacion_onClick()
		end if
	end sub
	
	sub window_onLoad()
		if frmHistoria.cbxOpcionAutorizacion.value = ""  then
			frmHistoria.chkTipoAutorizacion.disabled = true 
		end if			
	end sub	
	'************* FIN 16-04-2009************************
	
		
//-->
</script>
<!--#INCLUDE virtual="/Compartido/Encabezado.htm" -->

<!-- JFMG 16-04-2009-->
<!--#INCLUDE virtual="/Sucursal/Rutinas.asp" -->
<!--#INCLUDE virtual="/Compartido/Rutinas.asp" -->
<!-- ****************** FIN ********************-->

<body>
<form id="frmHistoria" method="post">
<center>
<table align="center" id="tabNuevaHistoria" class="borde" BORDER="0" cellpadding="4" cellspacing="0" style="WIDTH: 370px">	
<tr><td class="Titulo" colspan="2" style="FONT-SIZE: 10pt; HEIGHT: 5px">
      <p><%=sTitulo%></p>   
     </td>
</tr>
<tr>
	<td>Tipo<br>
		<select name="cbxTipo" style="width: 200px">
			<option value="1">Información</option>
			<option value="2">Advertencia</option>
			<option value="3">Grave</option>
		</select>
	</td>
</tr>

<!-- JFMG 16-04-2009 se agrega para agregar historias de autorización -->
	<tr>
		<td><input type="CheckBox" name="chkTipoAutorizacion">Tipo Autorizaci&oacute;n<br>
			<select name="cbxOpcionAutorizacion" disabled>
				<% CargarComboOpcionAutorizacion %>
			</select>
			
			<select name="cbxValorOpcionAutorizacion" style="display: none">
				<% CargarComboValorOpcionAutorizacion %>
			</select>
		</td>
	</tr>
<!-- **************************** FIN  16-04-2009 **********************************-->

<tr>
	<td>
		Descripción <br><TEXTAREA id=txtDescripcion style="WIDTH: 370px; HEIGHT: 50px; TEXT-ALIGN: left" name=txtDescripcion cols=2 size="89"></TEXTAREA>
		<input name="hddTextoAutorizacion" type="hidden" value="">
	</td>
	<td>
	</td>
</tr>
<tr>
</tr>
<tr align="middle">
	<td colspan="2"><IMG id=imgAceptar style="CURSOR: hand" height =20 src="../images/BotonAceptar.jpg" width=70 ></td>
</tr></table>
</center><!--#INCLUDE virtual="/Compartido/Rutinas.htm" -->
</form>
</body>
</html>