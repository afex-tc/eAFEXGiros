<%@ Language=VBScript %>
<%
	'Se asegura que la página no se almacene en la memoria cache
	Response.Expires = 0
	If Session("CodigoCliente") = "" Then
		response.Redirect "../Compartido/TimeOut.htm"
		response.end
	End If
%>
<!-- #INCLUDE virtual="/Compartido/Errores.asp" -->
<!-- #INCLUDE virtual="/Compartido/Rutinas.asp" -->
<%
	dim sFecha
	
	sFecha = date

	Sub CargarEmpleados(ByVal Seleccionado)
		Dim rs, sMayMin, sSelect, sSQL
			
		sSQL = "select codigo_empleado, nombre_completo from empleado where estado=1 order by nombre_completo "
		set rs = ejecutarsqlcliente(session("afxCnxRemunera"), sSQL)
		if err.number <> 0 then
			Response.Redirect "../Compartido/Error.asp?description=" & err.Description
		end if
		
		If  Not rs.EOF Then
			Response.write "<option value=></option>"
			Do Until rs.eof	
				If UCASE(Trim(rs("codigo_empleado"))) = Ucase(Trim(Seleccionado)) Then
					sSelect = "SELECTED"
				Else
					sSelect = ""
				End If
				sMayMin = trim(rs("nombre_completo"))
				sMayMin = MayMin(sMayMin) 
				Response.write "<option " & sSelect & " value=" & _
				trim(rs("codigo_empleado")) & ">" & _
				sMayMin & " </option> "
				If Err.number <> 0 Then
					Set rs = Nothing
					MostrarErrorMS "Cargar Empleados"
				End If
				rs.MoveNext
			Loop
		End If
		Set rs = Nothing	
	End Sub

	sub CargarPorcentage(Porcentage)
		dim i, sSelected, a
		Response.write "<option value=></option>"
		a = 50
		for i = 1 to 2
			if cint("0" & Porcentage) = cint(a) then
				sSelected = "SELECTED"
			else
				sSelected = ""				
			end if
			Response.write "<option " & sSelected & " value=" & a & ">" & a & " </option> "
			a = 100
		next
	end sub

	Response.Expires = 0

%>
<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<title>Contacto Cliente</title>
<link rel="stylesheet" type="text/css" href="../Estilos/Cliente.css">
</head>

<script LANGUAGE="VBScript">
<!--
	window.dialogWidth = 26
	window.dialogHeight = 16
	window.dialogLeft = 240
	window.dialogTop = 220
	window.defaultstatus = ""
	
	
	sub window_onload()		
		if window.cbxPorcentage1.value <> "" then 
			window.cbxPorcentage1.disabled = false
			window.cbxContacto2.disabled = false
		end if
		if window.cbxContacto2.value <> "" then 
			window.cbxContacto2.disabled = false
		end if
		if window.cbxPorcentage2.value <> "" then 
			window.cbxPorcentage2.disabled = false
		end if			
		
		'msgbox "<%=request("contacto1")%>" & "," & "<%=request("porcentagecontacto1")%>" & "," & _
		'		"<%=request("contacto2")%>" & "," & "<%=request("porcentagecontacto2")%>"
		
		if window.txtfechaactivacion.value = "" then
			window.txtfechaactivacion.value = "<%=date%>"
		end if	

	end sub
	
	sub cbxContacto1_onchange()
		if window.cbxContacto1.value <> "" then 
			window.cbxPorcentage1.disabled = false
			if window.cbxPorcentage2.value <> "" then
				window.cbxPorcentage1.value = 100 - ccur(window.cbxPorcentage2.value)
			else
				window.cbxPorcentage1.value = 100
				window.cbxContacto2.disabled = false
			end if
		else
			window.cbxPorcentage1.value = ""
			window.cbxPorcentage1.disabled = true
			window.cbxContacto2.value = ""
			window.cbxContacto2.disabled = true
			window.cbxPorcentage2.value = ""
			window.cbxPorcentage2.disabled = true
		end if
	end sub
	
	sub cbxPorcentage1_onchange()
		if window.cbxPorcentage1.value <> "" then 			
			window.cbxContacto2.disabled = false
			
			if window.cbxPorcentage2.value <> "" then
				window.cbxPorcentage2.value = 100 - ccur(window.cbxPorcentage1.value)
			end if
		else			
			window.cbxContacto2.value = ""
			window.cbxContacto2.disabled = true
			window.cbxPorcentage2.value = ""
			window.cbxPorcentage2.disabled = true
		end if
	end sub
	
	sub cbxContacto2_onchange()
		if window.cbxContacto2.value <> "" then 			
			window.cbxPorcentage2.disabled = false
			if window.cbxPorcentage1.value <> "" then
				window.cbxPorcentage2.value = 50'100 - ccur(window.cbxPorcentage1.value)
				cbxPorcentage2_onchange()
			end if
		else
			window.cbxPorcentage2.value = ""
			window.cbxPorcentage2.disabled = true
			
			if window.cbxPorcentage1.value <> "" then
				window.cbxPorcentage1.value = 100
			end if
		end if
	end sub
	
	sub cbxPorcentage2_onchange()
		if window.cbxPorcentage2.value <> "" then 			
			if window.cbxPorcentage1.value <> "" then
				window.cbxPorcentage1.value = 100 - ccur(window.cbxPorcentage2.value)
			end if
		else			
			window.cbxPorcentage1.value = 100
		end if
	end sub
	
	Sub imgAceptar_onClick()
		window.cbxContacto1.disabled = false
		window.cbxPorcentage1.disabled = false
		window.cbxContacto2.disabled = false
		window.cbxPorcentage2.disabled = false
		
		
		if window.cbxPorcentage1.value = "" then
			window.cbxContacto1.value = ""			
			window.cbxContacto2.value = ""
			window.cbxPorcentage2.value = ""
		end if
		if window.cbxPorcentage2.value = "" then			
			window.cbxContacto2.value = ""			
		end if
		
		window.returnvalue = window.cbxContacto1.value & ";" & window.cbxPorcentage1.value & ";" & _
							 window.cbxContacto2.value & ";" & window.cbxPorcentage2.value & ";" & _
							 window.txtfechaactivacion.value	
		window.close		
	End Sub		

//-->
</script>
<body>
<center>
	<table border="0" cellpadding="1">
		<tr>
			<td>
				Contacto 1<br>
				<select name="cbxContacto1">
					<%cargarempleados request("contacto1")%>
				</select>
			</td>
			<td><b>%</b><br>
				<select name="cbxPorcentage1" style="widht: 3px" disabled>
					<%cargarporcentage request("porcentagecontacto1")%>
				</select>
			</td>
		</tr>
		<tr>
			<td>
				Contacto 2<br>
				<select name="cbxContacto2" disabled>
					<%cargarempleados request("contacto2")%>
				</select>
			</td>
			<td><b>%</b><br>
				<select name="cbxPorcentage2" style="widht: 3px" disabled>
					<%cargarporcentage request("porcentagecontacto2")%>
				</select>
			</td>
		</tr>
		<tr>
			<td>Fecha Activación<br>
				<input type="text" name="txtFechaActivacion" value="<%=request("fechaactivacion")%>" size="8">&nbsp;(dd-mm-aaaa)
			</td>		
		<tr>
			<td colspan="2" align="right"><img border="0" id="imgAceptar" onclick src="../images/BotonAceptar.jpg" WIDTH="70" HEIGHT="20"></td>
		</tr>		
	</table>

</center>
</body>
</html>