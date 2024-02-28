<%@ LANGUAGE = VBScript %>
<%
		'Se asegura que la página no se almacene en la memoria cache
		Response.Expires = 0
		Response.Buffer = True
		If session("afxCnxAFEXchange") = "" Then
			response.Redirect "Compartido/TimeOut.htm"
			response.end
		End If
		
		Session("AutorizaVisorAP") = 1
		
%>
<!-- #INCLUDE virtual="/compartido/Constantes.asp" -->
<!-- #INCLUDE virtual="/compartido/Rutinas.asp" -->
<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<title>AFEX</title>
<link rel="stylesheet" type="text/css" href="Estilos/Principal.css">
<style TYPE="text/css">
</style> 
</head>
<body >
<!-- <BGSOUND SRC="afexweb.wav"> -->
<a href id="Inicio" STYLE="display: none"></a>
<table border="0" cellspacing="0" cellpadding="0" bordercolordark="white" style="LEFT: 0px">
<tr>
<td width="410">
	<table border="0" cellspacing="0" cellpadding="0" bordercolordark="white" STYLE="BORDER-RIGHT: lightgrey 1px solid; LEFT: -1px; TOP: 0px" align="left">
<!--	<tr><td><img src="images/FondoPrincipal.jpg" WIDTH="404" HEIGHT="93"></td></tr>-->
	
	<tr><td><!-- #INCLUDE virtual="Complemento/_transferencias.htm" --></td></tr>
	<tr><td><!-- #INCLUDE virtual="Complemento/_giros.htm" --></td></tr>
	<tr><td><!-- #INCLUDE virtual="Complemento/_cambios.htm" --></td></tr>
	<tr><td><!-- #INCLUDE virtual="Complemento/_sucursales.htm" --></td></tr>

<!--	<tr><td>		<table border="0" cellspacing="2" cellpadding="0" bordercolordark="white" bordercolorlight="lightgrey">		<tr>		<td width="133">			<table STYLE="BORDER-RIGHT: lightgrey 1px solid">         <TBODY>			<tr><td><IMG border=0 src="images/AfexWireTransferChico.png"></td></tr>			<tr height="50"><td></td></tr>			<tr><td>Requiere hacer pagos en el extranjero? Hágalo con nuestros <A href="Cheques.asp" target=Principal>Cheques en Moneda Extranjera</A> 				o nuestro servicio de <A href="Transferencia.asp" target=Principal>Transferencias Bancarias</A></td></tr>			</tr>			</table>		</td>		<td width="133">			<table STYLE="BORDER-RIGHT: lightgrey 1px solid">			<tr><td><IMG border=0 src="images/AfexMoneyExpressChico.png"></td></tr>			<tr height="10"><td></td></tr>			<tr><td><A href="Giros.asp" target=Principal>Quiere enviar dinero a Chile o el extranjero?</A><br>				AFEX hace posible el envío y recepción de dinero desde Chile a cualquier				lugar del mundo y también desde el extranjero a Chile.</td></tr>			</table>		</td>		<td width="133">			<table>			<tr><td><img border="0" src=""></td></tr>			<tr><td>3 Información para esta columna</td></tr>			</table>		</td>				</tr>		</table>	</td></tr>-->
	</table>
</td>
<td width="168">
	<table border="0" cellspacing="2" cellpadding="2" height="100%" bordercolor="white" bordercolordark="lightgrey" width="100%">
	
		<tr><td><!-- #INCLUDE virtual="Complemento/_visorap.htm" --></td></tr>
		<tr><td><!-- #INCLUDE virtual="Complemento/_atencionclientes.htm" --></td></tr>
		<tr><td><!-- #INCLUDE virtual="Complemento/_noticias.htm" --></td></tr>
		<tr><td><!-- #INCLUDE virtual="Complemento/_hagasecliente.htm" --></td></tr>
		
		<tr height="100%"><td>	
		
	</td></tr>
	</table>
</td>
</tr>
</table>
<table style="BORDER-TOP: lightgrey 1px solid; MARGIN-BOTTOM: 0px" width="588">
	<tr align="middle" height="40"><td style="FONT-WEIGHT: bold">AFEX Ltda. Todos los derechos reservados</td></tr>
</table>
<script language="VBScript">
<!--
	Sub Promocion_onfocus()
		exit sub
		Inicio.focus 
		window.open "GirosCuba.htm", "", "dialogHeight= 600pxl; dialogWidth= 740pxl; " & _
					"dialogTop= 0; dialogLeft= 10; resizable=no; " & _
					"status=no; scrollbars=no"
	End Sub
	
	Sub HagaseCliente_onfocus()
		Inicio.focus 
		navigate "HagaseClienteNew.asp"
	End Sub

	Sub sCambios_OnClick()
		window.open "Cabecera.asp?Opcion=3", "Cabecera"
	End Sub
	
	Sub window_onLoad()
		window.open "Cabecera.asp?Opcion=1", "Cabecera" 
	End Sub
	
	Sub AbrirVisorAP()
		window.open "VisorAP/_visorap.asp", "VisorAP", "height=180,width=170,top=200, left=200,status=no,toolbar=no,menubar=no,location=no,channelmode=no"
	End Sub

	
-->
</script>
</body>
</html>
