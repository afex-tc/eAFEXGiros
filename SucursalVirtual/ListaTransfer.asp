<%@ Language=VBScript %>
<!--#include virtual="/compartido/rutinas.asp"-->
<%
	dim sSQL
	dim rs	
	
	' saca la lista de las transfer enviadas por el cliente	
	sSQL = "select * " & _
		   " from vtransferencia t " & _
				" inner join cliente c on c.codigo_cliente = t.codigo_cliente and " & _
										 " c.codigo_corporativa = " & Session("CodigoCliente") & _
		   " where estado_transferencia <> 0 and " & _
				" (sw_visible <> 0 or sw_visible is null) and " & _
				" fecha_transferencia between " & evaluarstr(request("FechaDesde")) & " and " & evaluarstr(request("FechaHasta")) & _
		   " order by fecha_transferencia "			   
	set rs = ejecutarsqlcliente(session("afxCnxAFEXchange"), sSQL)
	if err.number <> 0 then
		Response.Write "Ocurrió un error al buscar las Transferencias del Cliente. " & err.Description 
		Response.End
	end if
'	if rs.eof then
'		set rs = nothing
'		Response.Write "No se encontraron las Transferencias del Cliente. " & err.Description 
'		Response.End
'	end if	
	
%>

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link href="estilosucursalvirtual.css" rel="stylesheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">

<style type="text/css">
<!--
body {
	margin-top: 1px;
	margin-left: 1px;
}

-->
</style>



</HEAD>
<script language="vbscript">
<!--
	sub cmdBuscar_onclick()
		if txtFechaDesde.value = empty or txtFechaHasta.value = empty then exit sub
		
		window.navigate "listatransfer.asp?FechaDesde=" & txtFechaDesde.value & "&Fechahasta=" & txtFechaHasta.value
	end sub
-->
</script>


<BODY>

	<!-- CALENDARIO -->
	<script type="text/javascript" src="calendar.js"></script>
	<script type="text/javascript" src="calendar-setup.js"></script>
	<script type="text/javascript" src="lang/calendar-en.js"></script>
   <style type="text/css"> @import url("calendar-green.css"); </style>

<style type="text/css">
<!--
.estilonavs {
	font-family: Tahoma, Arial, Helvetica, sans-serif;
	font-size: 10px;
	font-style: normal;
	line-height: 10px;
	color: #FFFFFF;
	text-decoration: none;
	font-weight: bold;
}
a:link {
	font-family: Tahoma, Arial, Helvetica, sans-serif;
	font-size: 10px;
	color: #296129;
	text-decoration: none;
	font-weight: normal;
}
a:hover {
	font-family: Tahoma, Arial, Helvetica, sans-serif;
	font-size: 10px;
	color: #296129;
	text-decoration: none;
	font-weight: normal;
}
a:visited {
	font-family: Tahoma, Arial, Helvetica, sans-serif;
	font-size: 10px;
	color: #296129;
	text-decoration: none;
	font-weight: normal;
}
a:active {
	font-family: Tahoma, Arial, Helvetica, sans-serif;
	font-size: 10px;
	color: #296129;
	text-decoration: none;
	font-weight: normal;
}
-->
<!--
.buscar {
	font-family: Tahoma, Arial, Helvetica, sans-serif;
	font-size: 10px;
	font-style: normal;
	font-weight: normal;
	color: #990000;
	height: 16px;
	width: 75px;
	border: 1px solid #FF9900;
}
text13 {
	font-family: Tahoma, Arial, Helvetica, sans-serif;
	font-size: 12px;
	font-style: normal;
	line-height: normal;
	font-weight: bold;
	font-variant: normal;
	color: #990000;
	text-decoration: none;
}
.text13 {
	font-family: Tahoma, Arial, Helvetica, sans-serif;
	font-size: 13px;
	font-style: normal;
	font-weight: bold;
	font-variant: normal;
	color: #333333;
	text-decoration: none;
	line-height: normal;
}
.text9 {
	font-family: Tahoma, Arial, Helvetica, sans-serif;
	font-size: 9px;
	font-style: normal;
	font-weight: normal;
	font-variant: normal;
	color: #ffffff;
	text-decoration: none;
	line-height: 12px;
}
.text10 {
	font-family: Tahoma, Arial, Helvetica, sans-serif;
	font-size: 10px;
	font-style: normal;
	font-weight: normal;
	font-variant: normal;
	color: #666666;
	text-decoration: none;
	line-height: 12px;
}
.text10bold {
	font-family: Tahoma, Arial, Helvetica, sans-serif;
	font-size: 10px;
	font-style: normal;
	font-weight: bold;
	font-variant: normal;
	color: #0000FF;
	text-decoration: none;
}
.text10bold2 {
	font-family: Tahoma, Arial, Helvetica, sans-serif;
	font-size: 10px;
	font-style: normal;
	font-weight: bold;
	font-variant: normal;
	color: #666666;
	text-decoration: none;
}
.text11 {
	font-family: Tahoma, Arial, Helvetica, sans-serif;
	font-size: 11px;
	font-style: normal;
	font-weight: bold;
	font-variant: normal;
	color: #333333;
	text-decoration: none;
	line-height: 12px;
}

#calendario{
	font-family: Tahoma, Arial, Helvetica, sans-serif;
	font-size: 10px;
	text-align: center;
	font-weight: bold;
	margin-left: auto;
	margin-right: auto;
}
/*#mes para configurar aspectos de la caja que muestra el mes y el año*/
#mes{
	font-weight: bold;
	text-align: center;
	color: #CC6633;
	background-color: #E4CAAF;
}
/*.diaS para configurar aspectos de la caja que muestra los días de la semana*/
.diaS{
	color: #ffffff;
	background-color: #666666;
}
/*.celda para configurar aspectos de la caja que muestra los días del mes*/
.celda {
	background-color: #FFFFFF;
	color: #000000;
	font-weight : normal;
	cursor: default;
}
/*.Hoy para configurar aspectos de la caja que muestra el día actual*/
.Hoy{
	color: #ffffff;
	background-color: #666666;
	font-weight: normal;
	cursor: default;
}
#miCalendario{
	text-align: center;
}
/*.selectores para configurar aspectos de los campos para el mes y el año*/
.selectores{
	font-family: tahoma;
	font-size: 10px;
	color: #990000;
	margin-bottom: 2px;
	margin-top: 2px;
}
-->
</style>
<style type="text/css">
<!--
.dia {color: #C0C0C0; text-decoration:none;}
.dia_nor {color: #666666; text-decoration:none; font-weight:bold;}
.dia_sel {color: #E83700; text-decoration:none; font-weight:bold;}
-->
</style>

<style type="text/css">
<!--
.Estilo4 {font-family: Tahoma; font-size: 12px; color: #FFFFFF; font-weight: bold; }
a:link {
	color: #333333;
	text-decoration: none;
}
a:visited {
	text-decoration: none;
	color: #333333;
}
a:hover {
	text-decoration: underline;
}
a:active {
	text-decoration: none;
}
.Letra8 {
	font-family: Tahoma;
	font-size: 12px;
	color: #000000;
}
.TextBOx {
	font-family: Tahoma;
	font-size: 11px;
	font-weight: normal;
	color: #000000;
	border-top-width: 1px;
	border-right-width: 1px;
	border-bottom-width: 1px;
	border-left-width: 1px;
	border-top-style: solid;
	border-right-style: solid;
	border-bottom-style: solid;
	border-left-style: solid;
	border-top-color: #333333;
	border-right-color: #333333;
	border-bottom-color: #333333;
	border-left-color: #333333;
	text-align: left;
}
.formcelda {
	border: 1px solid #999999;
	cursor: auto;
}
.por {
	font-family: Tahoma;
	font-size: x-small;
	color: #000000;
}
.combobox {
	font-family: Tahoma;
	font-size: 11px;
	color: #000000;
	border: 1px solid #333333;
}
.selectores {	font-family: verdana;
	font-size: 10px;
	color: #990000;
	margin-bottom: 2px;
	margin-top: 2px;
}
.selectores1 {	font-family: verdana;
	font-size: 10px;
	color: #990000;
	margin-bottom: 2px;
	margin-top: 2px;
}
.selectores1 {font-family: verdana;
	font-size: 10px;
	color: #990000;
	margin-bottom: 2px;
	margin-top: 2px;
}
.Estilo16 {font-family: Tahoma; font-size: 11px; color: #333333; font-weight: bold; }
.celdafuera {	border-top: 1px none #999999;
	border-right: 1px solid #999999;
	border-bottom: 1px solid #999999;
	border-left: 1px solid #999999;
}
.textarea1 {	font-family: Tahoma;
	font-size: 11px;
	color: #333333;
	background-color: #FFFFFF;
	cursor: auto;
border-bottom-color:#0000FF border: 1px solid #FF3300;
	text-align: justify;
	letter-spacing: normal;
	border: 1px solid #CCCCCC;
}
.Estilo17 {	font-family: Tahoma;
	font-size: 11px;
	color: #333333;
	background-color: #CCCCCC;
	cursor: hand;
    border-bottom-color:#0000FF border: 1px solid #FF3300;
	border-top: 1px solid #CCCCCC;
	border-right: 1px solid #333333;
	border-bottom: 1px solid #333333;
	border-left: 1px solid #CCCCCC;
	text-align: center;
	letter-spacing: normal;
}
.TextBOx1 {
	font-family: Tahoma;
	font-size: 11px;
	font-weight: normal;
	color: #000000;
	text-align: left;
	border: 1px solid #CCCCdd;
}
body {
	margin-top: 1px;
}
-->
</style>



<table width="540" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td bgcolor="#31514A" class="titulos"><img src="../Img/transferenciaenviadas.jpg" width="215" height="16"></td>
  </tr>
</table>
<P>
	<input name="txtFechaDesde" id="txtFechaDesde" type="text" class="Borde_tabla_abajo" value="<%=request("FechaDesde")%>" size="8">
	<input name="Submit" type="submit" id="cal-button-1" class="Estilo17" value=":Fecha:">
        <script type="text/javascript">
            Calendar.setup({
              inputField    : "txtfechadesde",
              button        : "cal-button-1",
              align         : "Tr"
            });
                        </script>
	<input name="txtFechaHasta" id="txtFechaHasta" type="text" class="Borde_tabla_abajo" value="<%=request("FechaHasta")%>" size="8">
    <input name="Submit" type="submit" id="cal-button-2" class="Estilo17" value=":Fecha:">
        <script type="text/javascript">
            Calendar.setup({
              inputField    : "txtfechahasta",
              button        : "cal-button-2",
              align         : "Tr"
            });
                        </script>
&nbsp;&nbsp;
<img src="../Img/botonbuscar.jpg" name="cmdBuscar">  </P>
<table width="100%" border="0" cellpadding="0" cellspacing="1">
  <tr>
    <td width="20%">&nbsp;</td>    
    <td width="33%">&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
</table>
<table width="540" border="0" cellpadding="0" cellspacing="1" bgcolor="#31514A" style="font-family: tahoma; font-size: 12px">
		<tr>
		  <td><div align="center"><b class="tituloslistatransfer1">Fecha</b></div></td>
		  <td><div align="center"><b class="tituloslistatransfer1">Beneficiario </b></div></td>
		  <td class="tituloslistatransfer1"><div align="center"><b>Cuenta </b></div></td>		  
		  <td><div align="center"><b class="tituloslistatransfer1">Moneda</b></div></td>
		  <td><div align="center"><b class="tituloslistatransfer1">Monto</b></div></td>
		  <td><div align="center"><b class="tituloslistatransfer1">Monto </b></div></td>
		  <td><div align="center"><b class="tituloslistatransfer1">Estado</b></div></td>
	  </tr>
		<tr>
			<td>&nbsp;</td>
			<td>&nbsp;</td>
			<td class="tituloslistatransfer1"><div align="center"><b>Destino</b></div></td>
			<td>&nbsp;</td>
			<td>&nbsp;</td>
			<td><div align="center"><b class="tituloslistatransfer1">USD</b></div></td>
			<td>&nbsp;</td>
		</tr>
		<%Do Until rs.EOF%>
			<a href="enviartransfer.asp?transferencia=<%=rs("correlativo_transferencia")%>">	
			<tr style="HEIGHT: 25px" language="javascript" onmouseover="javascript:this.bgColor='#e7e7ef'" onmouseout="javascript:this.bgColor='#ffffff'" bgColor="#ffffff" style="cursor: hand">
				<td width="60" class="textoempresa">&nbsp;<%=rs("fecha_transferencia")%></td>
				<td width="150" class="textoempresa">&nbsp;<%=rs("nombre_titular_destino")%></td>
				<td width="50" class="textoempresa">&nbsp;<%=rs("cuenta_corriente_destino")%></td>				
				<td width="100" class="textoempresa">&nbsp;<%=rs("nombre_moneda")%></td>
				<td width="40" ALIGN="right" class="textoempresa"><%=FormatNumber(rs("monto_transferencia"), 2)%></td>
				<% If cCur(0 & rs("monto_equivalente")) = 0 Then %>
					<td width="50" ALIGN="right" class="textoempresa"><%=FormatNumber(rs("monto_transferencia"), 2)%></td>
				<% Else %>
					<td width="50" ALIGN="right" class="textoempresa"><%=FormatNumber(rs("monto_equivalente"), 2)%></td>
				<% End If %>
				<td width="50" class="textoempresa">&nbsp;<%=rs("nombre_estado")%></td>
				<td width="60">
					<%if rs("numero_transferencia") <> "" then%>
						<img src="../img/botonswift.jpg" onClick="window.open 'J:/Archivos/Transferencias/Swift/<%=TRIM(rs("numero_transferencia"))%>.jpg'">
					<%end if%>
				</td>
			</tr>
			</a>			
		<%			
			rs.MoveNext
		Loop
		%>
</table>
</BODY>
</HTML>
<%set rs = nothing%>