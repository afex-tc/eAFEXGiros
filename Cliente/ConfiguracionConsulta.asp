<%@ Language=VBScript %>
<%


%>
<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<title>Configuración Consulta de Giros</title>
<link rel="stylesheet" type="text/css" href="../Estilos/Cliente.css">
</head>

<script LANGUAGE="VBScript">
<!--
	On Error Resume Next

	Sub public_put_Tipo(iTipoConsulta)
		Dim objTemp
		Dim sTr
		Dim i
		
		Select Case iTipoConsulta
	
			Case 1, 2, 3	'Consulta de Giros
				sTr = "Giro"
			
			Case 4	'Consulta de Transferencias
				sTr = "Transfer"
			
			Case 9
				sTr = ""
			
		End Select		
		Set objTemp = document.all.tags("TR")
		For i = 0 To objTemp.length - 1
			If objTemp(i).title = sTr Then objTemp(i).style.display = ""
		Next 

	End Sub
	
	Sub imgAceptar_onClick()
		'window.navigate "ListaGiros.asp?Tipo=1&Titulo=" + "<%=sTitulo%>"
      window.event.returnValue = False
	   window.external.raiseEvent "Aceptar", "Aceptar"
	End Sub		

	Function public_put_CodigoCliente(sCliente)
		txtCodigoCliente.value =  sCliente
	End Function
	
	Function public_get_CodigoCliente()
		public_get_CodigoCliente = txtCodigoCliente.value 
	End Function

	Function public_put_NombreCliente(sCliente)
		txtNombres.value =  sCliente
	End Function
	
	Function public_get_NombreCliente()
		public_get_NombreCliente = txtNombres.value 
	End Function

	Function public_put_ApellidoCliente(sCliente)
		txtApellidos.value =  sCliente
	End Function
	
	Function public_get_ApellidoCliente()
		public_get_ApellidoCliente = txtApellidos.value 
	End Function

	Function public_put_Desde(sDesde)
		txtDesde.value =  sDesde
	End Function
	
	Function public_get_Desde()
		public_get_Desde = txtDesde.value 
	End Function

	Function public_put_Hasta(sHasta)
		txtHasta.value =  sHasta
	End Function
	
	Function public_get_Hasta()
		public_get_Hasta = txtHasta.value 
	End Function

	Sub txtDesde_onBlur()
		Dim sFecha
		sFecha = txtDesde.value
		If ValidarFecha(sFecha) Then
			txtDesde.value = sFecha
		Else
			txtDesde.focus 
			txtDesde.select 
		End If
	End Sub

	Sub txtHasta_onBlur()
		Dim sFecha

		sFecha = txtHasta.value
		If ValidarFecha(sFecha) Then
			txtHasta.value = sFecha
		Else
			txtHasta.focus 
			txtHasta.select 
		End If
	End Sub
	
//-->
</script>
<body>
<center>
<table id="tabConsulta" class="borde" BORDER="0" cellpadding="4" cellspacing="0" style="HEIGHT: 195px; width:300px">	
<tr><td class="Titulo" colspan="2" style="FONT-SIZE: 10pt; HEIGHT: 5px">Datos de la consulta</td></tr>
<tr>
	<td align="center">	
		<table id="tabPeriodo" class="bordeinactivo" cellspacing="0" cellpadding="3">
        <tbody>
		<tr>
			<td colspan="2" class="tituloinactivo" >Periodo</td>
		</tr>
		<tr>
			<td>Desde el</td> 
			<td><input SIZE="8" VALUE="01-01-2002" id="txtDesde" ></td>
		</tr>
		<tr>
			<td>Hasta el</td>
			<td><input SIZE="8" VALUE="01-01-2002" id="txtHasta"></td>
			</td>
		</tr>
		</table></td></tr>
<tr align="middle">
	<td colspan="2"><img height="25" id="imgAceptar" src="../images/BotonAceptar.jpg" style="CURSOR: hand" width="80"></td>
</tr></tbody></table>
</center>
<!--#INCLUDE virtual="/Compartido/Rutinas.htm" -->
</body>
</html>