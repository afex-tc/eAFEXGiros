<%@ LANGUAGE = VBScript %>
<!--#INCLUDE virtual="/Compartido/Errores.asp" -->
<%
	Response.Expires = 0
	If Session("CodigoCliente") = "" Then
	   response.Redirect "Default.asp"
	   response.end
	End If

	Dim sCliente
	Dim sNombre
	Dim rsTransfer
	Dim aTransfer
	Dim i, sMoneda, nTipo
	
	'------------Jonathan Miranda - 09-08-2002
	'Se agrega la moneda como filtro
	
	' JFMG 25-09-2012
	'sCliente = " codigo_cliente = '" & Request("CodigoCliente") & "' AND " & _
	'			  " estado_transferencia <> 0 " 				  
	sMoneda = Request("CodigoMoneda")
	'If sMoneda <> "" Then
	'	sCliente = sCliente & " AND codigo_moneda = '" & sMoneda & "' "
	'End If
	sNombre = Request("NombreCliente")
	
	'Set afxTransfer = Server.CreateObject("AfexProducto.Transferencia")
    
	On Error Resume Next
	'If Err.number <> 0 Then
	'	Set afxTransfer = Nothing
	'	Set rsTransfer = Nothing	
	'	MostrarErrorMS "Ultimas Transferencias 1"
	'End If	
	''MostrarErrorMS Session("afxCnxAFEXchange") & ", " & sCliente
	'Set rsTransfer = afxTransfer.Buscar(Session("afxCnxAFEXchange"), 6, sCliente, 30,,,2)
	Set rsTransfer = ejecutarsqlcliente(Session("afxCnxAFEXchange"), "exec MostrarUltimasTransferenciasCliente '" & Request("CodigoCliente") & "', '" & sMoneda & "'")
	If rsTransfer.EOF Then 
		'Set afxTransfer = Nothing
		Set rsTransfer = Nothing	
		Response.Redirect "../compartido/error.asp?Titulo=Ultimas Transferencias 2&description=No se encontraron transferencias para el cliente"
	End If
	'If afxtransfer.ErrNumber <> 0 Then
	'	Set rsTransfer = Nothing
	'	MostrarErrorAFEX afxTransfer, "Ultimas Transferencias 3"
	'End If
	
	'Set afxTransfer = Nothing
	''Set rsTransfer = Nothing
	' FIN JFMG 25-09-2012
	
'*************************************************************************
' Funciones Y Procedimientos
'*************************************************************************
	
	Dim sEncabezadoFondo
	Dim sEncabezadoTitulo
	
	sEncabezadoFondo = Request("Fondo")
	If sEncabezadoFondo = "" Then
		sEncabezadoFondo = "Transfers"
	End If
	sEncabezadoTitulo = sNombre
	nTipo = cInt(0 & Request("Tipo"))
		
%>
<!--#INCLUDE virtual="/Compartido/Rutinas.asp" -->
<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link rel="stylesheet" type="text/css" href="../Estilos/Cliente.css">
<title><%=sNombre%> - Ultimas Transferencias</title>
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
	<% If nTipo = 0 Then %>
		window.dialogWidth = 29
		window.dialogHeight = 30
		window.dialogLeft = 300
		window.dialogTop = 70
		window.defaultstatus = "Listo"	
	<% End If %>	
	Sub cbxTransfer_onKeyUp()
		cbxTransfer_onClick()
	End Sub

	Sub cbxTransfer_onKeyDown()
		cbxTransfer_onClick()
	End Sub
	
	Sub cbxTransfer_ondblClick()
		Copiar_onClick
	End Sub
	
	Sub cbxTransfer_onClick()		
		<% 
			rsTransfer.movefirst
			For i = 1 To rsTransfer.RecordCount								
		%>
			If clng(0 & <%=rsTransfer("correlativo_transferencia")%>) = clng(0 & cbxTransfer.value) Then
				document.all.item("TD1").innerText = "<%=rsTransfer("nombre_banco_origen")%>" 
				document.all.item("TD2").innerText = "<%=rsTransfer("cuenta_corriente_origen")%>"
				document.all.item("TD3").innerText = "<%=rsTransfer("nombre_banco_destino")%>"
				document.all.item("TD4").innerText = "<%=rsTransfer("cuenta_corriente_destino")%>"
				document.all.item("TD5").innerText = "<%=rsTransfer("nombre_titular_destino")%>"
				document.all.item("TD6").innerText = "<%=rsTransfer("nombre_ciudad_destino")%>"
				document.all.item("TD7").innerText = "<%=rsTransfer("numero_aba")%>"
				document.all.item("TD8").innerText = "<%=rsTransfer("Invoice")%>"
				document.all.item("TD20").innerText = "<%=rsTransfer("further_credit")%>"
				document.all.item("TD9").innerText = "<%=rsTransfer("numero_transferencia")%>"
				document.all.item("TD10").innerText = "<%=rsTransfer("monto_transferencia")%>"
				document.all.item("TD11").innerText = "<%=rsTransfer("fecha_transferencia")%>"
				document.all.item("TD12").innerText = "<%=rsTransfer("nombre_estado")%>"
				document.all.item("TD13").innerText = "<%=rsTransfer("direccion_destino")%>"				
				document.all.item("TD14").innerText = "<%=rsTransfer("direccionbeneficiario")%>"								
				window.txtBancoInterm.value = "<%=rsTransfer("banco_intermediario")%>"
				window.txtCuentaInterm.value = "<%=rsTransfer("cuenta_intermediario")%>"
				window.txtCiudadInterm.value = "<%=rsTransfer("nombre_ciudad_intermediario")%>"
				window.txtDireccionInterm.value = "<%=rsTransfer("direccion_intermediario")%>"
				<%					
					rsTransfer.MoveNext
				%>				
			End If
		<%
			Next
		%>
	End Sub

	Sub Copiar_onClick()
				
		window.returnvalue = document.all.item("TD1").innerText & ";" & _
							 document.all.item("TD2").innerText & ";" & _
							 document.all.item("TD3").innerText & ";" & _
							 document.all.item("TD4").innerText & ";" & _
							 document.all.item("TD5").innerText & ";" & _
							 document.all.item("TD6").innerText & ";" & _
							 document.all.item("TD7").innerText & ";" & _
							 document.all.item("TD8").innerText & ";" & _
							 document.all.item("TD13").innerText & ";" & _
							 window.txtBancoInterm.value & ";" & _
							 window.txtCuentaInterm.value & ";" & _
							 window.txtCiudadInterm.value & ";" & _
							 window.txtDireccionInterm.value & ";" & _
							 document.all.item("TD20").innerText & ";" & _
							 document.all.item("TD14").innerText
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
<select NAME="cbxTransfer" SIZE="10" WIDTH="150" style="FONT-SIZE: 8pt; LEFT: 3px; WIDTH: 216px; FONT-FAMILY: MS Sans Serif; POSITION: absolute; TOP: 81px; HEIGHT: 268px">
<%	
	rsTransfer.moveFirst
	If Not rsTransfer.EOF Then
		Do Until rsTransfer.EOF
			Response.Write	"<OPTION VALUE=" & rsTransfer("correlativo_transferencia") & ">" & _
							Trim(rsTransfer("nombre_titular_destino")) & _
							", " & _
							Trim(rsTransfer("nombre_banco_destino")) & _
							"</OPTION>"
			rsTransfer.MoveNext
		Loop
	End If
	'Set rsTransfer = Nothing
%>
</select>
<table id="Transferencia" style="LEFT: 217px; WIDTH: 238px; POSITION: absolute; TOP: 81px; 
HEIGHT: 265px" cellSpacing="1" cellPadding="1" width="238" border="1">
		  
  <tr><td>
	<table id="tbTransfer" style="WIDTH: 226px; POSITION: relative; HEIGHT: 241px" cellSpacing="1" cellPadding="1" width="226" border="0">
		<tr><td WIDTH="26px"><b>Origen: </b></td><td ID="TD1" WIDTH="200px"></td></tr>
		<tr><td WIDTH="26px"><b>Cuenta: </b></td><td ID="TD2" WIDTH="200px"></td></tr>
		<tr><td WIDTH="26px"></td><td WIDTH="200px"></td></tr>
		<tr><td WIDTH="2px"><b>Destino: </b></td><td ID="TD3" WIDTH="200px"></td></tr>
		<tr><td WIDTH="10%"><b>Cuenta:</b></td><td ID="TD4" WIDTH="200px"></td></tr>
		<tr><td WIDTH="10%"><b>Nombre:</b></td><td ID="TD5" WIDTH="200px"></td></tr>
		<tr><td WIDTH="10%"><b>Direccion:</b></td><td ID="TD13" WIDTH="200px"></td></tr>
		<tr><td WIDTH="10%"><b>Ciudad:</b></td><td ID="TD6" WIDTH="200px"></td></tr>
		<tr><td WIDTH="10%"><b>Dirección Beneficiario:</b></td><td ID="TD14" WIDTH="200px"></td></tr>
		<tr><td WIDTH="10%"></td><td></td></tr>
		<tr><td WIDTH="10%"><b>ABA:</b></td><td ID="TD7" WIDTH="200px"></td></tr>
		<tr><td WIDTH="10%"><b>Invoice:</b></td><td ID="TD8" WIDTH="200px"></td></tr>
		<tr><td WIDTH="10%"><b>Crédito a:</b></td><td ID="TD20" WIDTH="200px"></td></tr>
		<tr><td WIDTH="10%"><b>Número:</b></td><td ID="TD9" WIDTH="200px"></td></tr>
		<tr><td WIDTH="10%"><b>Monto:</b></td><td ID="TD10" WIDTH="200px"></td></tr>
		<tr><td WIDTH="10%"><b>Fecha:</b></td><td ID="TD11" WIDTH="200px"></td></tr>
		<tr><td WIDTH="10%"><b>Estado:</b></td><td ID="TD12" WIDTH="200px"></td></tr>
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
		<input TYPE="image" name="Copiar" src="../images/BotonCopiar.jpg" border="0" alt="Copiar Transferencia" STYLE="POSITION: relative; TOP: 300px" WIDTH="70" HEIGHT="20">
	</center>
<% End If %>
</body>
</html>