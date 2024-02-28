<%@ Language=VBScript %>
<%
	'Se asegura que la página no se almacene en la memoria cache
	Response.expires = 0
	'Response.expiresabsolute = Now() - 1
	'Response.addHeader "pragma", "no-cache"
	'Response.addHeader "cache-control", "private"
	'Response.CacheControl = "no-cache"

	If Session("CodigoCliente") = "" Then
		response.Redirect "../Compartido/TimeOut.htm"
		response.end
	End If
%>
<!--#INCLUDE virtual="/Compartido/Constantes.asp" -->
<%

	'Variables de módulo
	'Variables para encabezado
	Dim sEncabezadoFondo
	Dim sEncabezadoTitulo
	Dim sSucursal
	
	Dim sTitulo, nTipo
	Dim nCampo, sArgumento, sArgumento2, sArgumento3, rs, nAccion
	Dim sCliente
	
	sTitulo = Request("Titulo")
	nAccion = cInt(0 & Request("Accion"))
	nTipo = cInt(0 & Request("Tipo"))
	sCliente = Trim(Request("Cl"))
	
	If Trim(sTitulo) = "" Then sTitulo = "Lista de Clientes"
	sEncabezadoFondo = "Consultas"
	sEncabezadoTitulo = sTitulo 

	nCampo = cInt(0 & Request("Campo"))
	sArgumento = request("Argumento")
	sArgumento2 = request("Argumento2")
	sArgumento3 = request("Argumento3")		
	sSucursal = Request("sc")
	sw = request("sw")
	
	If sSucursal <> "" Then
		Set rs = ObtenerCCP(sSucursal, sCliente,sw, request("cn"))	
		If rs.EOF Then
			Set rs = Nothing
		End If
	ElseIf sCliente <> "" Then
		Set rs = ObtenerCCP(sSucursal, sCliente,sw, request("cn"))	
		If rs.EOF Then
			Set rs = Nothing
		End If
	Else
		Set rs = Nothing
	End If
	
	Function ObtenerCCP(ByVal Sucursal, ByVal Cliente, ByVal sw, ByVal cn)
	   Dim rsATC
	   Dim sSQL
	   Dim Condicion
	   Dim Rut
	   Dim Raya
	   Dim sComa
		Dim sRt
	   Set ObtenerCCP = Nothing

	'   On Error Resume Next		

	   sSQL = "SELECT * FROM Cliente " & _
				 "WHERE 1 = 1 "
	   
		if Sucursal <> Empty Then
			if Sucursal <> "XX" then
				sSQL = sSQL & " AND sucursal_origen = '" & Sucursal & "' "
			end if
		end if

	   If Cliente <> Empty Then
	   		if sw = 0 then
				sSQL = sSQL & " AND nombre like '%" & Cliente & "%' "
	   		elseif sw = 1 then
				sRt = ValorRut(cliente)
				sSQL = sSQL & " AND rut = '" & sRt & "' "			
	   		end if
	   end if
	   
		if cn <> empty then sSQl = sSQL & cn

	   sSQL = sSQL & " ORDER BY nombre"
	   sComa = ""

'		response.write ssql & request("cn")
'		response.end

	   Set rsATC = EjecutarSQLCliente(Session("afxCnxCorporativa"), sSQL)

	   If Err.Number <> 0 Then 
			Set rsATC = Nothing
			MostrarErrorMS "Obtener CCP"
		End If
	   
	   Set rsATC.ActiveConnection = Nothing
	   Set ObtenerCCP = rsATC

	   Set rsATC = Nothing 
	End Function

	'Response.Redirect "../Compartido/Error.asp?Titulo=" & rs.eof	

%>
<!--#INCLUDE virtual="/Compartido/Rutinas.asp" -->
<!--#INCLUDE virtual="/sucursal/Rutinas.asp" -->
<!--#INCLUDE virtual="/Compartido/Errores.asp" -->

<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<title>AFEX</title>
<link rel="stylesheet" type="text/css" href="../Estilos/Cliente.css">
<link rel="stylesheet" type="text/css" href="../Estilos/AFX2004.css">
</head>
<script LANGUAGE="VBScript">
<!--
	Dim sOldClass

	Sub window_onload()
		' Jonathan Miranda G. 21-03-2007
		frmCli.cbxRiesgo.value = 0
		'------------------- Fin ------------------------
	End Sub		
	
	Sub txtCliente_onblur()
	Dim sRut
		if frmCli.optRut.checked=True then 
			If txtcliente.value = Empty Then Exit Sub
		
			sRut = ValidarRut(txtcliente.value)			
		
			If sRut <> "" Then
				txtcliente.value = sRut
			Else
				msgbox "Debe ingresar un número de rut valido ", vbOKOnly + vbInformation, "Ingreso de Rut"
				txtcliente.focus
				Exit Sub
			End If
		end if
	End Sub
	
	sub optRut_onclick()
		frmCli.optApel.Checked=false
		frmCli.optPO.Checked=false
		frmCli.optDH.Checked=false

		tblPerfil.style.display = "none"
	end sub
	sub optApel_onclick()
		frmCli.optRut.Checked=false
		frmCli.optPO.Checked=false
		frmCli.optDH.Checked=false

		tblPerfil.style.display = "none"
	end sub
	sub optPO_onclick()
		frmCli.optApel.Checked=false
		frmCli.optRut.Checked=false
		frmCli.optDH.Checked=false

		tblPerfil.style.display = ""
	end sub
	sub optDH_onclick()
		frmCli.optRut.Checked=false
		frmCli.optApel.Checked=false
		frmCli.optPO.Checked=false		

		tblPerfil.style.display = "none"
	end sub

	Sub Seleccionar()

		'msgbox window.event.srcElement.tagname
		sOldClass = window.event.srcElement.className 
		window.event.srcElement.className  = "Seleccionado"

	End Sub
	
	Sub QuitarSeleccion()
	
		'msgbox window.event.srcElement.tagname
		window.event.srcElement.className  = sOldClass
		
	End Sub

	Sub cmdAceptar_onClick()
		dim sw
		dim sWhere		

		If frmCli.optRut.checked = True Then sw = 1 Else sw = 0

		' Jonathan Miranda G. 20-03-2007
		If frmCli.optPO.checked = True Then 
			sw = 3
			
			if frmCli.cbxRiesgo.value > 0 then	
				sWhere = " and nivelriesgo = " & frmCli.cbxRiesgo.value 
			end if
			if frmCli.cbxPerfilPEP.value > 0 then	
				sWhere = sWhere & " and ppep = " & frmCli.cbxPerfilPEP.value 
			end if
			if frmCli.cbxPerfilZona.value > 0 then	
				sWhere = sWhere & " and pzona = " & frmCli.cbxPerfilZona.value 
			end if
			if frmCli.cbxPerfilRS.value > 0 then	
				sWhere = sWhere & " and presidencia = " & frmCli.cbxPerfilRS.value
			end if
			if frmCli.cbxPerfilACT.value > 0 then	
				sWhere = sWhere & " and pactividad = " & frmCli.cbxPerfilACT.value 
			end if
			if frmCli.cbxPerfilIndustria.value > 0 then	
				sWhere = sWhere & " and pindustria = " & frmCli.cbxPerfilIndustria.value 
			end if
            if frmCli.cbxPerfilCliente.value > 0 then
                sWhere = sWhere & " and Perfil_Cliente = " & frmCli.cbxPerfilCliente.value 
            end if
		
		elseIf frmCli.optDH.checked = True Then 
			sw = 3

			sWhere = " and estado = 0 "
		end if
		if cbxSucursal.value = "XX" and txtCliente.value = empty and _
			sWhere = empty then exit sub

		'------------------- Fin ----------------------------------

		window.navigate "ListaClientes.asp?sc=" & cbxSucursal.value & "&Cl=" & txtCliente.value & "&sw=" & sw & "&cn=" & sWhere
	End Sub
		
//-->
</script>
<body id="bb" border="0" style="margin: 2 2 2 2" >
	<table class="Borde" id="" BORDER="0" cellpadding="0" cellspacing="0" style="HEIGHT: 150px; width:100%; background-color: #f4f4f4">	
	<tr height="40" style="background-color: #ffeeaa; #ffdd77; #e1e1e1"><td colspan="3" style="font-size: 16pt">&nbsp;&nbsp;Lista Clientes</td></tr>
	<tr height="1" style="background-color: silver "><td colspan="3" ></td></tr>
	<tr height="4"><td colspan="3" ></td></tr>
	<tr>
		<td colspan="3">
		<table width="100%">
		<tr>
			<td align="center">
				<br>
				<table>
				<tr>
					<td colspan="2">
						<!--#INCLUDE virtual="/sucursal/MenuListaClientes.asp" -->
					</td>
					</div>
				</tr>
				</table>
			</td>
		</tr>
		</table><br>
		</td>
	</tr>

	<tr height="10"><td></td></tr>

<!--<table border="0" cellspacing="0" cellpadding="0" style="LEFT: 0px; POSITION: relative; TOP: 0px">-->

<tr height="10"><td>
	<table cellspacing="1" cellpadding="1" ID="tbReporte" border="0" ALIGN="center" STYLE="background-color: silver; COLOR: #505050; FONT-FAMILY: Verdana; FONT-SIZE: 10px; POSITION: relative; TOP: 0px">
	<tr CLASS="Encabezado" style="background-color: #e1e1e1; height: 25px" align="center">
		<td WIDTH="280">
			<b>Nombre</b>
		</td>
		<td WIDTH="100">
			<b>Rut</b>
		</td>
		<td WIDTH="100">
			<b>Pasaporte</b>
		</td>
		<td align="center" WIDTH="">
			<b>Ficha</b>
		</td>
		<td WIDTH="">
			<b>Docs</b>
		</td>
	</tr>
	<%
		Dim i, nCodigo, sDetalle, sNegocio, sPagina, sRut
				
		Select Case nAccion
		Case afxAccionIngresarMG
		Case Else
		End Select
		'mostrarerrorms nAccion & ", " & sPagina
		i = 0
		If rs Is Nothing Then
		Else		
			Do Until rs.EOF 
					i = ccur(i) + 1
				%>
					<!--<a href="<%=sPagina%>?Accion=<%=afxAccionBuscar%>&Campo=<%=nCodigo%>&Argumento=<%=sCodigoCliente%>">-->
					<!--<tr CLASS="<%=sDetalle%>" sonmouseover="Seleccionar()" sonmouseout="QuitarSeleccion()" >-->
					<tr style="HEIGHT: 25px" language="javascript" onmouseover="javascript:this.bgColor='#f1f1f1'; window.status='<%=rs("nombre")%>'" onmouseout="javascript:this.bgColor='white'; window.status=''" bgColor="white" style="cursor: hand">
					<% If rs("nivelriesgo") = 2 Then %>
                        <a href="http:DetalleClienteCautela.asp?cc=<%=rs("codigo")%>" starget="_blank">
                            <td><%=rs("nombre")%></td>
							<td><%=FormatoRut(rs("rut"))%></td>
							<td><%=rs("pasaporte")%></td>
						</a>
                    <% Else %>
                        <a href="http:DetalleCliente.asp?cc=<%=rs("codigo")%>&rt=<%=rs("rut")%>&nc=<%=rs("nombre")%>" starget="_blank">
                            <td><%=rs("nombre")%></td>
							<td><%=FormatoRut(rs("rut"))%></td>
							<td><%=rs("pasaporte")%></td>
						</a>
                    <% End If %>
						
						<td align="center"><a href="http:ListaImagen.asp?cc=<%=rs("codigo")%>&td=1,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40" target="_blank"><img src="http:../images/transferencia.jpg" border="0"></a></td>
						<td align="center"><a href="http:ListaImagen.asp?cc=<%=rs("codigo")%>&td=2,3,4,5,6,7,8,9,10"  target="_blank"><img src="http:../images/cheque.jpg" border="0"></a></td>
					</tr>
					<!--</a>-->
				<%
				rs.MoveNext
			Loop
		End If
		%>
	</table>
	<br>
</td></tr>
</table>
</body>
<script language="VBScript">

	Sub objConsulta_OnScriptletEvent(strEventName, varEventData)
		
	   Select Case strEventName
	   
			Case "Aceptar"
				msgbox strEventName & ", " & objConsulta.Desde
				
		End Select
		
	End Sub
	
	
</script>

</html>
<%
	Set rs = Nothing
%>