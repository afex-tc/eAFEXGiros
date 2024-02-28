<%@ Language=VBScript %>
<%
	'Se asegura que la p�gina no se almacene en la memoria cache
	Response.Expires = 0
	If Session("CodigoCliente") = "" Then
		response.Redirect "../Compartido/TimeOut.htm"
		response.end
	End If
%>
<!-- #INCLUDE virtual="/Compartido/Errores.asp" -->
<!-- #INCLUDE virtual="/Compartido/Rutinas.asp" -->
<!-- #INCLUDE virtual="/agente/Constantes.asp" -->
<%

%>
<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<title>Datos del Tercero</title>
<link rel="stylesheet" type="text/css" href="../Estilos/Cliente.css">
</head>

<script LANGUAGE="VBScript">
<!--
	window.dialogWidth = 35
	window.dialogHeight = 21
	window.dialogLeft = 100
	window.dialogTop = 220
	window.defaultstatus = ""
	
	Sub window_onLoad()
	End Sub
	
	
	Sub imgAceptar_onClick()
		Dim nTipo, sParentesco, sDescripcion, sNombre
		
		If optAviso1.checked Then
			nTipo = 1
			sDescripcion = Trim(observacion.value)
			
		ElseIf optAviso2.checked Then
			nTipo = 2
			sDescripcion = trim(Observacion_1.value)
			sParentesco = trim(Parentesco.value)
			sNombre = trim(NombreParentesco.value)
		
		ElseIf optAviso3.checked Then
			nTipo = 3
			sDescripcion = trim(Motivo.value)
		
		End If
			
		window.returnvalue = nTipo & ";" & sDescripcion & ";" & _
			sParentesco & ";" & sNombre
			
		window.close
		
	End Sub		

	sub optAviso1_onClick()				' opci�n "Contest� Beneficiario"
		Opcionvalue = 1
		' habilita el campo de observaci�n de esta opci�n
		Observacion.disabled = false
		Observacion.focus

		' deshabilita los dem�s objetos
		optAviso2.checked = false
		Parentesco.disabled = true
		NombreParentesco.disabled = true
		Observacion_1.disabled = true
		optAviso3.checked = false
		Motivo.disabled = true
	end sub
	
	sub optAviso2_onClick()				' opci�n "Contest� otra Persona"
		Opcionvalue = 2
		' habilita el campo de observaci�n de esta opci�n
		Parentesco.disabled = false
		NombreParentesco.disabled = false
		Observacion_1.disabled = false
		Parentesco.focus

		' deshabilita los dem�s objetos
		optAviso1.checked = false
		Observacion.disabled = true
		optAviso3.checked = false
		Motivo.disabled = true
	end sub
	
	sub optAviso3_onClick()				' opci�n "No Avisado"
		Opcionvalue = 3
		' habilita el campo de observaci�n de esta opci�n
		Motivo.disabled = false
		Motivo.focus

		' deshabilita los dem�s objetos
		optAviso1.checked = false
		Observacion.disabled = true
		optAviso2.checked = false
		Parentesco.disabled = true
		NombreParentesco.disabled = true
		Observacion_1.disabled = true
	end sub

//-->
</script>
<body sbgcolor="PowderBlue">
<center>
<input type="hidden" id="txtCodigoCliente">
<table id="tabConsulta" class="border" BORDER="0" cellpadding="4" cellspacing="0" style="position: relative; left: 4px; HEIGHT: 195px" width="350px">	
<tr><td colspan="2"><input TYPE="radio" NAME="optAviso1" checked>Contest� Beneficiario</td></tr>
<tr>
	<td>
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Observaci�n
	</td>
	<td>
		<input type="text" name="Observacion" value size="60" maxlength="40">
	</td>
</tr>
<tr><td colspan="2"><input TYPE="radio" NAME="optAviso2">Contest� otra Persona</td></tr>
<tr>
	<td>
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Parentesco
	</td>
	<td>
		<select name="Parentesco" disabled>
			<option value="ABUELO">Abuelo</option>
			<option value="AMIGO">Amigo</option>
			<option value="CONYUGE">C�nyuge</option>
			<option value="EMPLEADO">Empleado</option>
			<option value="HERMANO">Hermano</option>
			<option value="HIJO">Hijo</option>
			<option value="MADRE">Madre</option>
			<option value="NIETO">Nieto</option>
			<option value="NUERA">Nuera</option>
			<option value="PADRE">Padre</option>
			<option value="PRIMO">Primo</option>								
			<option value="SECRETARIA">Secretaria</option>
			<option value="SOBRINO">Sobrino</option>
			<option value="SUEGRO">Suegro</option>
			<option value="TIO">Tio</option>
			<option value="VECINO">Vecino</option>
			<option value="YERNO">Yerno</option>
		</select>
	</td>
</tr>
<tr>
	<td>
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Nombre
	</td>
	<td>
		<input type="text" name="NombreParentesco" value size="60" maxlength="40">
	</td>
</tr>
<tr>
	<td>
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Observaci�n
	</td>
	<td>
		<input type="text" name="Observacion_1" value size="60" maxlength="40">
	</td>
</tr>
<tr>
	<td>
		<input TYPE="radio" NAME="optAviso3">No Avisado
	</td>
</tr>
<tr>
	<td>
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Motivo
	</td>
	<td>
		<select name="Motivo" disabled>
			<option value="BUZON DE VOZ">Buz�n de voz</option>
			<option value="TELEFONO FUERA DE SERVICIO">Tel�fono fuera de servicio</option>
			<option value="TELEFONO NO CONTESTA">Tel�fono no contesta</option>
			<option value="TELEFONO OCUPADO">Tel�fono ocupado</option>
			<option value="TELEFONO NO EXISTE">Tel�fono no existe</option>								
		</select>
	</td>
</tr>
<tr align="middle">
	<td colspan="2"><img id="imgAceptar" src="../images/BotonAceptar.jpg" style="CURSOR: hand" WIDTH="70" HEIGHT="20"></td>
</tr>
</table>
</center>
<!--#INCLUDE virtual="/Compartido/Rutinas.htm" -->
</body>
</html>