<%@ Language=VBScript %>
<%
	'Se asegura que la página no se almacene en la memoria cache
	Response.expires = 0
	If Session("CodigoCliente") = "" Then
		response.Redirect "../Compartido/TimeOut.htm"
		response.end
	End If
%>
<!-- #INCLUDE virtual="/Compartido/Errores.asp" -->
<!-- #INCLUDE virtual="/Compartido/Rutinas.asp" -->
<!-- #INCLUDE virtual="/agente/Constantes.asp" -->
<%
	Dim sNombres, sRut, sPasaporte
	Dim sDireccion, sCiudad
	Dim sNombreCiudad, sNombrePaisPass
	Dim sPaisFono, sAreaFono, sFono
	Dim sNombres2, sRut2, sPasaporte2
	Dim sDireccion2, sCiudad2
	Dim sNombreCiudad2, sNombrePaisPass2
	Dim sPaisFono2, sAreaFono2, sFono2
	Dim nAccion, sAFEXchange, sAFEXpress, sAFEXpressEliminar
	Dim nTipoCliente, sDisabled
	Dim nCampo, sArgumento
	
	nAccion = cInt(0 & Request("Accion"))
	
	sAFEXpress = Request("Nuevo")
	sAFEXpressEliminar = Request("Eliminar")
	sAFEXchange = Request("AFEXchange")
	'mostrarerrorms sAFEXpressEliminar & ", " & sAFEXpress
	'If sAFEXchange <> "" Then
	'	nCampo = 5
	'	sArgumento = sAFEXchange 
	
	If sAFEXpress <> "" Then
		nCampo = 6
		sArgumento = sAFEXpress	
		CargarCliente 
	End If

	If sAFEXpressEliminar <> "" Then
		nCampo = 6
		sArgumento = sAFEXpressEliminar
		CargarClienteEliminar
	End If
	


	Sub CargarCliente()
		Dim rs			
		Set rs = BuscarCliente(nCampo, sArgumento, "", "")
		'Response.Redirect "../Compartido/Error.asp?Titulo=" & rs.eof
		If Not rs.EOF Then
			nTipoCliente = cInt(0 & rs("tipo"))
			sNombres = MayMin(EvaluarVar(rs("nombre_completo"), ""))
			sRut = FormatoRut(EvaluarVar(rs("rut"), ""))
			sPasaporte = EvaluarVar(rs("pasaporte"), "")
			sDireccion = MayMin(EvaluarVar(rs("direccion"), ""))
			sCiudad = Trim(Ucase(EvaluarVar(rs("codigo_ciudad"), "")))
			sPaisFono = EvaluarVar(rs("ddi_pais"), "")
			sAreaFono = EvaluarVar(rs("ddi_area"), "")
			sFono = EvaluarVar(rs("telefono"), "")
			sNombreCiudad = MayMin(EvaluarVar(rs("ciudad"), ""))
			sNombrePaisPass = MayMin(EvaluarVar(rs("paispas"), ""))
			If Err.number <> 0 Then
				MostrarErrorMS ""
			End If
		End If
		Set rs = Nothing
	End Sub
	
	Sub CargarClienteEliminar()
		Dim rs2
		
		
		Set rs = BuscarCliente(nCampo, sArgumento, "", "")
		'Response.Redirect "../Compartido/Error.asp?Titulo=" & rs.eof
		If Not rs.EOF Then
			nTipoCliente = cInt(0 & rs("tipo"))
			sNombres2 = MayMin(EvaluarVar(rs("nombre_completo"), ""))
			sRut2 = FormatoRut(EvaluarVar(rs("rut"), ""))
			sPasaporte2 = EvaluarVar(rs("pasaporte"), "")
			sDireccion2 = MayMin(EvaluarVar(rs("direccion"), ""))
			sCiudad2 = Trim(Ucase(EvaluarVar(rs("codigo_ciudad"), "")))
			sPaisFono2 = EvaluarVar(rs("ddi_pais"), "")
			sAreaFono2 = EvaluarVar(rs("ddi_area"), "")
			sFono2 = EvaluarVar(rs("telefono"), "")
			sNombreCiudad2 = MayMin(EvaluarVar(rs("ciudad"), ""))
			sNombrePaisPass2 = MayMin(EvaluarVar(rs("paispas"), ""))
			If Err.number <> 0 Then
				MostrarErrorMS ""
			End If
		End If
		Set rs = Nothing
	End Sub

%>
<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link rel="stylesheet" type="text/css" href="../Estilos/Cliente.css">
</head>
<script LANGUAGE="VBScript">
<!--
	'Variables de módulo
	'Variables para encabezado
	Dim sEncabezadoFondo
	Dim sEncabezadoTitulo
	
	Const afxCodigoCliente = 3
	Const afxRut = 1
	Const afxPasaporte = 2
	COnst afxNombres = 4
	
	sEncabezadoFondo = "Principal"
	sEncabezadoTitulo = "Asociar Clientes"


	Sub window_onLoad()
		
		CargarMenuActualizar 
		CargarCliente
		CargarClienteEliminar
		
	End Sub

	Sub optRut_onClick()
		window.frmcliente.optpasaporte.checked=False
		window.frmcliente.txtpasaporte.style.display="none"
		window.frmcliente.cbxPaisPasaporte.style.display="none"
		lblPaisPasaporte.style.display="none"
		window.frmCliente.txtRut.style.display = ""		
		frmCliente.txtPasaporte.value = ""
	End Sub
	
	Sub optPasaporte_onClick()
		window.frmcliente.optRut.checked=False
		window.frmCliente.txtRut.style.display = "none"		
		window.frmcliente.txtpasaporte.style.display=""
		window.frmcliente.cbxPaisPasaporte.style.display=""
		lblPaisPasaporte.style.display=""
		frmCliente.txtRut.value = ""
	End Sub


	Sub CargarCliente()
		frmCliente.txtNombres.value = "<%=sNombres%>"
		frmCliente.txtDireccion.value = "<%=sDireccion%>"
		frmCliente.txtPaisFono.value = "<%=sPaisFono%>" 
		frmCliente.txtAreaFono.value = "<%=sAreaFono%>"
		frmCliente.txtFono.value = "<%=sFono%>"
		frmCliente.txtRut.value = "<%=sRut%>" & "<%=sPasaporte%>"
		frmCliente.txtCiudad.value = "<%=sNombreCiudad%>"
	End Sub

	Sub CargarClienteEliminar()
		frmCliente.txtNombres2.value = "<%=sNombres2%>"
		frmCliente.txtDireccion2.value = "<%=sDireccion2%>"
		frmCliente.txtPaisFono2.value = "<%=sPaisFono2%>" 
		frmCliente.txtAreaFono2.value = "<%=sAreaFono2%>"
		frmCliente.txtFono2.value = "<%=sFono2%>"
		frmCliente.txtCiudad2.value = "<%=sNombreCiudad2%>"
	End Sub

	Sub CargarMenuActualizar()
		Dim sId
		
		frmCliente.objmenu.bgColor = document.bgColor 
		frmCliente.objmenu.stylesheet = "../Estilos/Cliente.css"
		sId = frmCliente.objmenu.addparent("Opciones")
		frmCliente.objMenu.addchild sId, "Asociar", "Asociar", "Principal"
	End Sub
	

-->
</script>

<body >
<!-- #INCLUDE virtual="/Compartido/Encabezado.htm" -->
<form id="frmCliente" method="post">
<input type="hidden" name="txtExchange">
<input type="hidden" name="txtExpress">
<table class="Borde" ID="tabPaso1" CELLSPACING="0" border="0" height="200px" style="position: relative; top: 0px; left: 2px;">
	<tr HEIGHT="15">
		<td colspan="3" class="titulo">Cliente con Identificación</td>
	</tr>
	<tr HEIGHT="2">
		<td></td>
		<td COLSPAN="2"></td>		
	</tr>
	<tr>
	<td width="1px"></td>
	<td colspan="1">
	<table width="100%" border="0" style="HEIGHT: 0px; WIDTH: 500px" cellpadding="0" cellspacing="0">
		<tr HEIGHT="40">
			<td VALIGN="center" colspan=""><br>
				<table border="0" cellpadding="1">
				<tr>
					<td>Identificacion<br>
						<input name="txtRut" style="text-align: right; width: 150px" disabled>
					</td>
				</tr>
				</table>
			</td>

		<td >
			<table border="0" width="280px" height="60px">
			<tr><td colspan="2"><object align="right" id="objMenu" style="HEIGHT: 50px; LEFT: 0px; POSITION: relative; TOP: 0px; WIDTH: 190px" type="text/x-scriptlet" width="190" VIEWASTEXT border="0"><param NAME="Scrollbar" VALUE="0"><param NAME="URL" VALUE="http:../Scriptlets/Menu.htm"></object></td></tr>
			<tr><td></div></td></tr>
			<tr height="100%"><td></td></tr>
			</table>
		</td>
			
		</tr>
		<tr></tr>
		<tr ID="trPersona" HEIGHT="0">
		<td colspan="3">
			<table>
			<tr>
				<td>Nombres<br>
				<input NAME="txtNombres" id="txtNombres" SIZE="25" style="width: 300px" onkeypress="IngresarTexto(2)" onblur="frmCliente.txtNombres.value=MayMin(frmCliente.txtNombres.value)"  disabled>
			</tr>
			<td>Teléfono<br>
				<input disabled name="txtPaisFono" style="width: 40px">
				<input disabled name="txtAreaFono" style="width: 40px">
				<input name="txtFono" style="width: 90px"onkeypress="IngresarTexto(1)" onblur="frmCliente.txtFono.value=MayMin(frmCliente.txtFono.value)" disabled>
			</td>
			</table>
		</td></tr>
		<tr>
			<td COLSPAN="">Dirección<br>
				<input STYLE="HEIGHT: 22px; WIDTH: 300px" SIZE="10" NAME="txtDireccion" onblur="frmCliente.txtDireccion.value=MayMin(frmCliente.txtDireccion.value)" disabled>
			</td>
			<td >Ciudad<br>
				<input name="txtCiudad" style="width: 180px" disabled>
			</td>
		</tr>
		<tr HEIGHT="2">
			<td></td>
		</tr>
	</table>
	
	</td>


	</tr>

	<tr height="20"><td></td></tr>
	<tr HEIGHT="15">
		<td colspan="3" class="titulo">Cliente Seleccionado</td>
	</tr>
	<tr HEIGHT="2">
		<td></td>
		<td COLSPAN="2"></td>		
	</tr>
	<tr>
	<td width="1px"></td>
	<td>
	<table width="100%" border="0" style="HEIGHT: 0px; WIDTH: 500px" cellpadding="0" cellspacing="0">
		<tr></tr>
		<tr ID="trPersona" HEIGHT="0">
		<td colspan="3">
			<table>
			<tr>
				<td>Nombres<br>
				<input disabled NAME="txtNombres2" id="txtNombres2" SIZE="25" style="width: 300px" onkeypress="IngresarTexto(2)" onblur="frmCliente.txtNombres.value=MayMin(frmCliente.txtNombres.value)">
			</tr>
			<td>Teléfono<br>
				<input disabled name="txtPaisFono2" style="width: 40px">
				<input disabled name="txtAreaFono2" style="width: 40px">
				<input disabled name="txtFono2" style="width: 90px"onkeypress="IngresarTexto(1)" onblur="frmCliente.txtFono.value=MayMin(frmCliente.txtFono.value)">
			</td>
			</table>
		</td></tr>
		<tr>
			<td COLSPAN="">Dirección<br>
				<input disabled STYLE="HEIGHT: 22px; WIDTH: 300px" SIZE="10" NAME="txtDireccion2" onblur="frmCliente.txtDireccion.value=MayMin(frmCliente.txtDireccion.value)">
			</td>
			<td >Ciudad<br>
				<input disabled name="txtCiudad2" style="width: 180px">
			</td>
		</tr>
		<tr HEIGHT="2">
			<td></td>
		</tr>
	</table>
	
	</tr>
	<tr HEIGHT="2">
		<td></td>
		<td COLSPAN="2"></td>		
	</tr>
	<tr>
	<td width="1px"></td>
	<td>
<!-- -->
	</td>
	</tr>

</table>	
</form>
</body>
<!--#INCLUDE virtual="/Compartido/Rutinas.htm" -->
<script>

	Sub objMenu_OnScriptletEvent(strEventName, varEventData)
		
	   Select Case strEventName
	   
			Case "linkClick"
				If Right(varEventData, 7) = "Asociar" Then
					frmCliente.action = "GrabarAsociarCliente.asp?Nuevo=<%=sAFEXpress%>&Eliminar=<%=sAFEXpressEliminar%>" & "&Giro=<%=Request("Giro")%>&AFEXchange=<%=Request("AFEXchange")%>&Tipo=<%=Request("Tipo")%>&Accion=<%=Request("Accion")%>"
					frmCliente.submit 
					frmCliente.action = "" 
				
				Else
					window.open varEventData, "Principal"
					
				End If
				
		End Select
		
	End Sub
	
</script>
</html>

