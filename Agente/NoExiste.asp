<%@ Language=VBScript %>
<%
	'Se asegura que la página no se almacene en la memoria cache
	Response.Expires = 0
	If Session("CodigoAgente") = "" Then
		response.Redirect "../Compartido/TimeOut.htm"
		response.end
	End If
%>
<%
sw=request.QueryString ("sw")
nmn=request.QueryString ("nmn")
mn=request.QueryString ("mn")
mto=request.QueryString ("mto")
bs= request.QueryString ("bs")
'response.Write mto
'response.Write bs
'response.Write mn
'response.Write nmn
%>
<!--#INCLUDE virtual="/agente/Constantes.asp" -->
<!--#INCLUDE virtual="/Compartido/Rutinas.htm" -->
<!--#INCLUDE virtual="/Compartido/Transferencias.asp" -->
<!--#INCLUDE virtual="/Compartido/Rutinas.asp" -->
<!--#INCLUDE virtual="/Compartido/Errores.asp" -->

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link rel="stylesheet" type="text/css" href="../Estilos/Principal.css">
<link rel="stylesheet" type="text/css" href="../Estilos/Cliente.css">
<link rel="stylesheet" type="text/css" href="../Estilos/Reportes.css">
</HEAD>
<script LANGUAGE="VBScript">
<!--
	'Variables de módulo
	'Variables para encabezado
	Dim sEncabezadoFondo, sEncabezadoTitulo
	
	sEncabezadoFondo = "Transacciones"
	sEncabezadoTitulo = "Venta Tarjeta Telefónica"
	sub window_onload()	
		If Not CajaPregunta("AFEX En Linea", "Está seguro que desea vender la tarjeta?") Then
		'exit sub		
			frmNoExiste.action = "VenderTarjeta.asp" 
			frmNoExiste.submit
	 		frmNoExiste.action = ""	 
	 	else	
	 	    frmNoExiste.action = "GrabarVentaTarjeta.asp?nmn= " & frmNoExiste.cbxMoneda1.value 
			frmNoExiste.submit
	 		frmNoExiste.action = ""	
	 	end if	
	end sub 	
-->
</script>
<BODY>
<!-- #INCLUDE virtual="/Compartido/Encabezado.htm" -->
<form id="frmNoExiste" method="post">
<input type="hidden" id="cbxMoneda1" name="cbxmoneda1" value="<%=nmn%>">
<input type="hidden" id="cbxMoneda" name="cbxMoneda" value="<%=mn%>">
<input type="hidden" id="cbxMonto" name="cbxMonto" value="<%=mto%>">
<input type="hidden" id="txtNumeroBoleta" NAME="txtNumeroBoleta" value="<%=bs%>">
<table ID="tabPaso2" CELLSPACING="0" CELLPADDING="0" BORDER="0" HEIGHT="100" WIDTH="560" STYLE="BORDER-BOTTOM: steelblue 2px solid; BORDER-LEFT: steelblue 2px solid; BORDER-RIGHT: steelblue 2px solid; BORDER-TOP: steelblue 2px solid;  HEIGHT: 80px; LEFT: 4px; POSITION: relative; TOP: 0px; WIDTH: 466px">
	<tr><td HEIGHT="1"></td></tr>
	<tr><td class="titulo">Valores</td></tr>
	<tr><td>
		<table ID="tabInformacion1">
			<tr>
				<td width="135">Moneda<br>
					<input NAME="txtMoneda1" SIZE="20" style="HEIGHT: 22px; WIDTH: 110px" value="<%=nmn%>" disabled>
				</td>
				<td>Monto<br>
					<input NAME="txtMonto1" SIZE="25"  style="HEIGHT: 22px; WIDTH: 140px" value="<%=mto%>" disabled>
					
				</td>
			</tr>
		</table>
		<table>
			<tr>
			</tr>
		</table>
	</td></tr>
	<tr><td class="titulo">Información</td></tr>
	<tr><td>
		<table ID="tabInformacion">
			<tr>
				<td width="54">Nº Boleta<br>
					<input id="txtNumeroBoleta" NAME="txtNumeroBoleta1"  style="HEIGHT: 22px; WIDTH: 130px" value="<%=bs%>" disabled>
					
				</td>			
				<td width="160">
				</td>
			</tr>			
		</table>		
		<table>
			<tr>
			</tr>
		</table>
	</td></tr>
</table>		
			

</form>
  

</BODY>
</HTML>
