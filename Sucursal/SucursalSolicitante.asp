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

	Sub CargarSucursalesIP(ByVal Seleccionado)
		Dim rs, sMayMin, sSelect, sSQL
			
		sSQL = "select nombre, ip_sucursal from sucursal order by nombre "
		set rs = ejecutarsqlcliente(session("afxCnxAFEXchange"), sSQL)
		if err.number <> 0 then
			Response.Redirect "../Compartido/Error.asp?description=" & err.Description
		end if
		
		If  Not rs.EOF Then
			Response.write "<option value=></option>"
			Do Until rs.eof	
				If UCASE(Trim(rs("ip_sucursal"))) = Ucase(Trim(Seleccionado)) Then
					sSelect = "SELECTED"
				Else
					sSelect = ""
				End If
				sMayMin = trim(rs("nombre"))
				sMayMin = MayMin(sMayMin) 
				Response.write "<option " & sSelect & " value=" & _
				trim(rs("ip_sucursal")) & ">" & _
				sMayMin & " </option> "
				If Err.number <> 0 Then
					Set rs = Nothing
					MostrarErrorMS "Cargar Sucursales IP"
				End If
				rs.MoveNext
			Loop
		End If
		Set rs = Nothing
	End Sub	

	Response.Expires = 0

%>
<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<title>Sucurasl Solicitante</title>
<link rel="stylesheet" type="text/css" href="../Estilos/Cliente.css">
</head>

<script LANGUAGE="VBScript">
<!--
	window.dialogWidth = 26
	window.dialogHeight = 10
	window.dialogLeft = 240
	window.dialogTop = 220
	window.defaultstatus = ""
	
	
	sub window_onload()		
		

	end sub
	
	sub imgAceptar_onClick()
		window.returnvalue = cbxSucursales.value 
		window.close()
	end sub
	
	
//-->
</script>
<body>
<center>
	<table border="0" cellpadding="1">
		<tr>
			<td>Seleccione la Sucursal que solicitó la habilitación<br>
				<select name="cbxSucursales">
					<%cargarsucursalesip request("sucursal")%>
				</select>
			</td>			
		</tr>
		<tr>
			<td colspan="2" align="right"><img border="0" id="imgAceptar" onclick src="../images/BotonAceptar.jpg" WIDTH="70" HEIGHT="20"></td>
		</tr>		
	</table>

</center>
</body>
</html>