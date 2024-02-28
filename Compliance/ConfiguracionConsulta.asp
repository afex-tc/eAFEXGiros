<%@ Language=VBScript %>
<%
	Dim rsAgente, Sql
	
	Sql = "Select " & _
				  "nombre, codigo_agente " & _
		  "From   " & _
				  "cliente " & _
		  "Where  " & _
				  "codigo_agente is not null " & _
		  "and	  tipo = 4" & _
		  "Order by " & _
				  "nombre_usuario"
		  
	Set rsAgente = CreateObject("ADODB.Recordset")
	rsAgente.Open Sql, Session("afxCnxCorporativa")
%>

<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<title>Configuración Consulta de Compras y Ventas</title>
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

	Function public_put_Agente(sAgente)
		Dim i
		
		If sAgente <> Empty Then
			For i = 0 To cbxAgente.length - 1
				cbxAgente.selectedIndex = i
				If Ucase(cbxAgente.value) = Ucase(sAgente) Then
					Exit For	
				End If
			Next
		End If
	End Function
	
	Function public_get_Agente()
		public_get_Agente = cbxAgente.value 
	End Function

	Function public_put_Porcentaje(sPorcentaje)
		txtPorcentaje.value =  sPorcentaje
	End Function
	
	Function public_get_Porcentaje()
		public_get_Porcentaje = txtPorcentaje.value 
	End Function

	Function public_put_optAFEX(sAFEX)
		If sAFEX Then
			optAFEX.checked = True
			optCliente.checked = False
		Else
			optAFEX.checked = False
			optCliente.checked = True
		End If
	End Function
	
	Function public_get_optAFEX()
		public_get_optAFEX = optAFEX.checked
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
	
	Sub txtPorcentaje_onBlur()
		txtPorcentaje.value = replace(txtPorcentaje.value, ".", ",")
	End Sub
	
	Sub optAFEX_onClick()
		optAFEX.checked = True
		optCliente.checked = False
	End Sub
	
	Sub optCliente_onClick()
		optCliente.checked = True
		optAFEX.checked  = False
	End Sub
	
//-->
</script>
<body>
<center>
<table align="center" id="tabConsulta" class="borde" BORDER="0" cellpadding="4" cellspacing="0" style="WIDTH: 300px">	
<tr><td class="Titulo" colspan="2" style="FONT-SIZE: 10pt; HEIGHT: 5px">
      <p>Datos de la consulta</p>   </td></tr>
<tr>
	<td>
		Agente<br>
		<select id="cbxAgente" name="select1" style="HEIGHT: 22px; WIDTH: 280px"> 
			<%Do Until rsAgente.EOF%>
				  <option value="<%=rsAgente("codigo_agente")%>"><%=rsAgente("nombre")%></option>
				  <%rsAgente.Movenext%>
			<%Loop%> 
			<%Set rsAgente = Nothing%>
		</select><br><br>
		Recargo (T/C)<br>
		<input STYLE="HEIGHT: 22px; TEXT-ALIGN: right; WIDTH: 55px" onkeypress="IngresarTexto(1)" value="0" id="txtPorcentaje">%
		<input type="radio" name="optAFEX">Sobre AFEX
		<input type="radio" name="optCliente">Sobre el cliente
	</td>
	<td>
		<table id="tabPeriodo" class="bordeinactivo" cellspacing="0" cellpadding="3" width="150">
			<tbody>
				<tr>
					<td colspan="2" class="tituloinactivo">
					<p>Periodo</p></td>
				</tr>
				<tr>
					<td>Desde el</td> 
					<td><input SIZE="8" VALUE="01-01-2002" id="txtDesde"></td>
				</tr>
				<tr>
					<td>Hasta el</td>
					<td><input SIZE="8" VALUE="01-01-2002" id="txtHasta"></td>
					</td>
				</tr>
		</table>
		<br>
	</td>
</tr>
<tr>
	<td>
	</td>
</tr>
<tr align="middle">
	<td colspan="2"><img id="imgAceptar" src="../images/BotonAceptar.jpg" style="CURSOR: hand" WIDTH="70" HEIGHT="20"></td>
</tr></tbody></table>
</center><!--#INCLUDE virtual="/Compartido/Rutinas.htm" -->
</body>
</html>