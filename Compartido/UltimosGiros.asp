<%@ LANGUAGE = VBScript %>
<!--#INCLUDE virtual="/Compartido/Errores.asp" -->
<%
	Response.Expires = 0
	If Session("CodigoCliente") = "" Then
	   response.Redirect "../compartido/TimeOut.htm"
	   response.end
	End If

	Dim sCliente
	Dim sNombre, sApellidos, nTipoGiro
	Dim rsGiros
	Dim afxGiros
	Dim i, sMoneda, nTipo
	
	sNombre = Request("NombreCliente")
	sCliente = trim(trim(sNombre) & " " & trim(sApellidos))
	nTipoGiro = cInt(0 & Request("TipoGiro"))
	'MostrarErrorMS nTipoGiro & ", " & sCliente & ", " & Request("CodigoCliente")
	Set afxGiros = Server.CreateObject("AfexGiroXP.Giro")
   
	On Error Resume Next
	If Err.number <> 0 Then
		Set afxGiros = Nothing
		Set rsGiros = Nothing	
		MostrarErrorMS "Ultimos Giros 1"
	End If	
	
	Set rsGiros = afxGiros.Lista(Session("afxCnxAFEXpress"), nTipoGiro, true, _
								 ,, Request("CodigoCliente"),,,,,, 10)
	If rsGiros.EOF Then 
		Set afxGiros = Nothing
		Set rsGiros = Nothing	
		Response.Redirect "../compartido/error.asp?Titulo=Ultimos Envios 2&description=No se encontraron giros del cliente " & sCliente
	End If
	If afxGiros.ErrNumber <> 0 Then
		Set rsGiros = Nothing
		MostrarErrorAFEX afxGiros, "Ultimos Envios 2"
	End If
	
	Set afxGiros = Nothing
	'Set rsTransfer = Nothing
	
'*************************************************************************
' Funciones Y Procedimientos
'*************************************************************************
	
	Dim sEncabezadoFondo
	Dim sEncabezadoTitulo
	
	sEncabezadoFondo = Request("Fondo")
	If sEncabezadoFondo = "" Then
		sEncabezadoFondo = "Giros"
	End If
	sEncabezadoTitulo = sNombre
	nTipo = cInt(0 & Request("Tipo"))
		
%>
<!--#INCLUDE virtual="/Compartido/Rutinas.asp" -->
<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link rel="stylesheet" type="text/css" href="../Estilos/Cliente.css">
<title><%=sNombre%> - Ultimos Envios</title>
<style>
	INPUT
	{
		FONT-FAMILY: Verdana;
		FONT-SIZE: 9pt;
		TEXT-TRANSFORM: none
	}
	TABLE
	{
		FONT-FAMILY: MS Sans Serif;
		FONT-SIZE: 8pt;
		FONT-LINE: 4px;
		TEXT-TRANSFORM: none;
		COLOR: Black
	}
	TR
	{
		HEIGHT: 5px;
	}
</style>
<script LANGUAGE="VBScript">
<!--
	Dim sNombreB, sApellidoB	
	Dim sCodigoPaisB, sCodigoCiudadB
	Dim	iPaisFonoB, iAreaFonoB, cFonoB							 
	Dim sAgentePagador, sMonedaPago
	Dim sCodigoBeneficiario , sRutB, sPasaporteB, sPaisPasapB
		
	<% If nTipo = 0 Then %>
		window.dialogWidth = 29
		window.dialogHeight = 30
		window.dialogLeft = 300
		window.dialogTop = 70
		window.defaultstatus = "Listo"	
	<% End If %>	
	Sub cbxGiros_onKeyUp()
		cbxGiros_onClick()
	End Sub

	Sub cbxGiros_onKeyDown()
		cbxGiros_onClick()
	End Sub
	
	Sub cbxGiros_ondblClick()
		Copiar_onClick
	End Sub
	
	Sub cbxGiros_onClick()		
		<% 
			rsGiros.movefirst
			For i = 1 To rsGiros.RecordCount								
		%>
			If "<%=trim(rsGiros("codigo_giro"))%>" = trim(cbxGiros.value) Then
			<% 
				Select Case nTipoGiro
				Case 	afxGirosEnviados
			%>
					sNombreB = "<%=MayMin(trim(rsGiros("nombre_beneficiario")))%>"
					sApellidoB = "<%=MayMin(trim(rsGiros("apellido_beneficiario")))%>"
				
					sCodigoPaisB = "<%=trim(rsGiros("pais_beneficiario"))%>"
					sCodigoCiudadB = "<%=trim(rsGiros("ciudad_beneficiario"))%>"
				
					iPaisFonoB = "<%=trim(rsGiros("codpais_beneficiario"))%>"
					iAreaFonoB = "<%=trim(rsGiros("codarea_beneficiario"))%>"
					cFonoB = "<%=trim(rsGiros("fono_beneficiario"))%>"
				
					sAgentePagador = "<%=trim(rsGiros("agente_pagador"))%>"
					sMonedaPago = "<%=trim(rsGiros("codigo_moneda"))%>"
					
					sCodigoBeneficiario = "<%=trim(rsGiros("Codigo_beneficiario"))%>"
					sRutB = "<%=trim(rsGiros("Rut_beneficiario"))%>"
					sPasaporteB = "<%=trim(rsGiros("Pasaporte_beneficiario"))%>"
					sPaisPasapB = "<%=trim(rsGiros("paispasap_beneficiario"))%>"
				
					document.all.item("TD1").innerText = "<%=rsGiros("fecha_captacion")%>" 
					document.all.item("TD2").innerText = "<%=MayMin(rsGiros("remitente"))%>"
					document.all.item("TD3").innerText = trim(trim(sNombreB) & " " & trim(sApellidoB))
					document.all.item("TD4").innerText = "<%=MayMin(rsGiros("direccion_beneficiario"))%>"
					document.all.item("TD5").innerText = "<%=MayMin(rsGiros("nombre_pais_beneficiario"))%>"
					document.all.item("TD6").innerText = "<%=MayMin(rsGiros("nombre_ciudad_beneficiario"))%>"
					document.all.item("TD7").innerText = "<%=rsGiros("codigo_giro")%>"
					document.all.item("TD8").innerText = "<%=rsGiros("invoice")%>"
					document.all.item("TD9").innerText = "<%=MayMin(rsGiros("estado"))%>"
					document.all.item("TD10").innerText = "<%=FormatNumber(rsGiros("monto_giro"), 2) %>"

			<%	Case afxGirosRecibidos %>
					sNombreB = "<%=MayMin(trim(rsGiros("nombre_remitente")))%>"
					sApellidoB = "<%=MayMin(trim(rsGiros("apellido_remitente")))%>"
				
					sCodigoPaisB = "<%=trim(rsGiros("pais_remitente"))%>"
					sCodigoCiudadB = "<%=trim(rsGiros("ciudad_remitente"))%>"
				
					iPaisFonoB = "<%=trim(rsGiros("codpais_remitente"))%>"
					iAreaFonoB = "<%=trim(rsGiros("codarea_remitente"))%>"
					cFonoB = "<%=trim(rsGiros("fono_remitente"))%>"
				
					sAgentePagador = "<%=trim(rsGiros("agente_pagador"))%>"
					sMonedaPago = "<%=trim(rsGiros("codigo_moneda"))%>"
				
					document.all.item("TD1").innerText = "<%=rsGiros("fecha_captacion")%>" 
					document.all.item("TD2").innerText = "<%=MayMin(rsGiros("remitente"))%>"
					document.all.item("TD3").innerText = trim("<%=MayMin(trim(rsGiros("nombre_beneficiario")))%>" & " " & "<%=MayMin(trim(rsGiros("apellido_beneficiario")))%>")
					document.all.item("TD4").innerText = "<%=MayMin(rsGiros("direccion_beneficiario"))%>"
					document.all.item("TD5").innerText = "<%=MayMin(rsGiros("nombre_pais_beneficiario"))%>"
					document.all.item("TD6").innerText = "<%=MayMin(rsGiros("nombre_ciudad_beneficiario"))%>"
					document.all.item("TD7").innerText = "<%=rsGiros("codigo_giro")%>"
					document.all.item("TD8").innerText = "<%=rsGiros("invoice")%>"
					document.all.item("TD9").innerText = "<%=MayMin(rsGiros("estado"))%>"
					document.all.item("TD10").innerText = "<%=FormatNumber(rsGiros("monto_giro"), 2) %>"
			<%	End Select %>
		<%					
					rsGiros.MoveNext
		%>				
			End If
		<% 
			Next
		%>
	End Sub

	Sub Copiar_onClick()
				
		window.returnvalue = sNombreB & ";" & _
							 sApellidoB & ";" & _
							 document.all.item("TD4").innerText & ";" & _
							 sCodigoPaisB & ";" & _
							 sCodigoCiudadB & ";" & _
							 iPaisFonoB & ";" & _
							 iAreaFonoB & ";" & _
							 cFonoB & ";" & _
							 document.all.item("TD10").innerText & ";" & _
							 sAgentePagador & ";" & _
							 sMonedaPago & ";" & _
							 sCodigobeneficiario & ";" & _
							 sRutB & ";" & _
							 sPasaporteB & ";" & _
							 sPaispasapB 
		window.close
		
	End Sub		
	
	Sub Volver_onClick()
		window.close
	End Sub
-->
</script>
</head>
<body background="../agente/imagenes/Giros_FondoVentana.jpg" bgcolor="#ffffff" stext="#008080" link="#0000ff" vlink="#000080">
<!--#INCLUDE virtual="/Compartido/Encabezado.asp" -->
<!--<font FACE="Verdana" SIZE="5">	<b><div ALIGN="center"><%=sNombre%></div></b></font><font FACE="Verdana" SIZE="2">	<div ALIGN="center">Ultimas Transferencias</div></font>-->
<select NAME="cbxGiros" SIZE="10" WIDTH="150" style="FONT-SIZE: 8pt; LEFT: 3px; WIDTH: 216px; FONT-FAMILY: MS Sans Serif; POSITION: absolute; TOP: 81px; HEIGHT: 268px">
<%	
	rsGiros.moveFirst
	Dim sDescripcion
	If Not rsGiros.EOF Then
		Do Until rsGiros.EOF
			Select Case nTipoGiro
			Case afxGirosEnviados
				sDescripcion = Trim(rsGiros("beneficiario"))
				
			Case afxGirosRecibidos
				sDescripcion = Trim(rsGiros("remitente"))
			End Select
			Response.Write	"<OPTION VALUE=" & rsGiros("codigo_giro") & ">" & _
							sDescripcion & _
							"</OPTION>"
			rsGiros.MoveNext
		Loop
	End If
	'Set rsTransfer = Nothing
%>
</select>
<table id="Giros" style="LEFT: 217px; WIDTH: 238px; POSITION: absolute; TOP: 81px; 
HEIGHT: 265px" cellSpacing="1" cellPadding="1" width="238" border="1">
		  
  <tr><td>
	<table id="tbGiros" style="WIDTH: 226px; POSITION: relative; HEIGHT: 241px" cellSpacing="1" cellPadding="1" width="226" border="0">
		<tr><td WIDTH="26px"><b>Fecha: </b></td><td ID="TD1" WIDTH="200px"></td></tr>
		<tr><td WIDTH="26px"><b>Remitente: </b></td><td ID="TD2" WIDTH="200px"></td></tr>
		<tr><td WIDTH="2px"><b>Beneficiario: </b></td><td ID="TD3" WIDTH="200px"></td></tr>
		<tr><td WIDTH="26px"></td><td WIDTH="200px"></td></tr>
		<tr><td WIDTH="2px"><b>Dirección: </b></td><td ID="TD4" WIDTH="200px"></td></tr>
		<tr><td WIDTH="2px"><b>Pais: </b></td><td ID="TD5" WIDTH="200px"></td></tr>
		<tr><td WIDTH="2px"><b>Ciudad: </b></td><td ID="TD6" WIDTH="200px"></td></tr>
		<tr><td WIDTH="10%"><b>Codigo Giro:</b></td><td ID="TD7" WIDTH="200px"></td></tr>
		<tr><td WIDTH="10%"><b>Invoice:</b></td><td ID="TD8" WIDTH="200px"></td></tr>
		<tr><td WIDTH="10%"><b>Estado:</b></td><td ID="TD9" WIDTH="200px"></td></tr>
		<tr><td WIDTH="10%"><b>Monto:</b></td><td ID="TD10" WIDTH="200px"></td></tr>
	</table>
  </td></tr>
		  
</table>

<!-- Datos del Intermediario -->
<input type="Hidden" name="txtBancoInterm" value>
<input type="Hidden" name="txtCuentaInterm" value>
<input type="Hidden" name="txtCiudadInterm" value>
<input type="Hidden" name="txtDireccionInterm" value>
<!-- Fin -->
<% If nTipo = 0 Then %>
	<center>
		<input TYPE="image" name="Volver" src="../images/BotonVolver.jpg" border="0" alt="Volver" STYLE="POSITION: Relative; TOP: 300px" WIDTH="70" HEIGHT="20">
		<input TYPE="image" name="Copiar" src="../images/BotonCopiar.jpg" border="0" alt="Copiar Giros" STYLE="POSITION: relative; TOP: 300px" WIDTH="70" HEIGHT="20">
	</center>
<% End If %>
</body>
</html>
