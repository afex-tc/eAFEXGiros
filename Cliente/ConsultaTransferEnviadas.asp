<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<title>AFEX</title>
<link rel="stylesheet" type="text/css" href="../Hoja%20de%20Estilos%201.css">
<style>
	TD
	{ BACK-COLOR: Black
	}
</style>
</head>
<script LANGUAGE="VBScript">
<!--
	Sub imgAceptar_onClick()

		'muestra el reporte		
		'window.open  "../Reportes/transferencia.rpt?" & _
		'				"init=actx&" & _
		'				"prompt0= &prompt1= &prompt2= " & _
		'				"&prompt3=" & "USD" & _
		'				"&prompt4=" & "Destino" & _
		'				"&prompt5=" & "Cuenta" & _
		'				"&prompt6=" & "Beneficiario" & _
		'				"&prompt7=" & "Aba" & _
		'				"&prompt8=" & "Direccion" & _
		'				"&prompt9=" & _
		'				"&prompt10=" & "1.000" & _
		'				"&prompt11= &prompt12= " & _
		'				"&prompt13= &prompt14= " & _
		'				"&prompt15=" & "Invoice" & _
		'				"&prompt16= &prompt17=" & date & _
		'				"&prompt18=" & date + 2 & _
		'				"&prompt19= &prompt20= " & _
		'				"&prompt21=" & "Para" &  _
		'				"&prompt22=AFEX TRANSFERENCIAS", "Principal"
		If window.tbReporte.style.display = "" then
			window.tbReporte.style.display = "none"
		Else 
			window.tbReporte.style.display = ""
		End If
	End Sub		

	Function imgAceptar_onMouseOver()
		window.imgAceptar.style.cursor = "Hand"		
	End Function


//-->
</script>
<body>
<marquee STYLE="HEIGHT: 96px; LEFT: 0px; POSITION: absolute; TOP: -28px; WIDTH: 437px" BEHAVIOR="slide" DIRECTION="right" SCROLLAMOUNT="30" SCROLLDELAY="50">
	<h6 STYLE="FONT-SIZE: 60pt">Consultas</h6>
</marquee>		
<marquee STYLE="HEIGHT: 74px; LEFT: 21px; POSITION: absolute; TOP: 1px; WIDTH: 394px" BEHAVIOR="slide" DIRECTION="down" SCROLLAMOUNT="8" SCROLLDELAY="100">		
	<h1 STYLE="COLOR: #cfcfcf">Transferencias Enviadas</h1>
</marquee>
<marquee STYLE="HEIGHT: 74px; LEFT: 20px; POSITION: absolute; TOP: 0px; WIDTH: 394px" BEHAVIOR="slide" DIRECTION="down" SCROLLAMOUNT="8" SCROLLDELAY="100">		
	<h1 STYLE="COLOR: steelblue">Transferencias Enviadas</h1>
</marquee>
<br><br><br>
<marquee STYLE="HEIGHT: 308px; LEFT: 91px; POSITION: absolute; TOP: 80px; WIDTH: 415px" BEHAVIOR="slide" DIRECTION="up" SCROLLAMOUNT="50" SCROLLDELAY="100">		
<font face="Verdana" size="4" COLOR="white"><br>
	<strong STYLE="BACKGROUND-COLOR: steelblue">&nbsp;Periodo&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</strong>
</font><br>
<div ID="divPeriodo" noWrap style="BORDER-BOTTOM: thin solid; BORDER-LEFT: thin solid; BORDER-RIGHT: thin solid; BORDER-TOP: thin solid; COLOR: steelblue; HEIGHT: 95px; LEFT: 0px; POSITION: relative; WIDTH: 341px">
	<center>
	<br>Desde el&nbsp;&nbsp; 
	<input SIZE="8" VALUE="01-01-2002" id="text2" name="text2">&nbsp;&nbsp;
	hasta el&nbsp;&nbsp;
	<input SIZE="8" VALUE="01-01-2002" id="text1" name="text1">
	<br><br>
	<IMG height=25 id=imgAceptar src="../images/BotonAceptar.jpg" style   ="LEFT: 122px; TOP: 54px" width=80 >
	</center>
</div>
</marquee>		
<table ID="tbReporte" width="100%" border="0" ALIGN="center" STYLE="COLOR: #505050; DISPLAY: none; FONT-FAMILY: Verdana; FONT-SIZE: 10px; POSITION: absolute; TOP: 250px">
	<tr BORDER="1">
		<td WIDTH="25%" ALIGN="middle" bgcolor="skyblue" STYLE="FONT-FAMILY: Verdana; FONT-SIZE: 11px">
			<b>Banco<br>Origen</b>
		</td>
		<td WIDTH="20%" ALIGN="middle" BGCOLOR="skyblue" STYLE="FONT-FAMILY: Verdana; FONT-SIZE: 11px">
			<b>Monto</b>
		</td>
		<td WIDTH="20%" ALIGN="middle" BGCOLOR="skyblue" STYLE="FONT-FAMILY: Verdana; FONT-SIZE: 11px">
			<b>Banco<BR>Destino</b>
		</td>
		<td WIDTH="18%" ALIGN="middle" BGCOLOR="skyblue" STYLE="FONT-FAMILY: Verdana; FONT-SIZE: 11px">
			<b>Fecha<br>Envío</b>
		</td>
		<td WIDTH="17%" ALIGN="middle" BGCOLOR="skyblue" STYLE="FONT-FAMILY: Verdana; FONT-SIZE: 11px">
			<b>Estado</b>
		</td>
	</tr>
	<%
'		i = 0
'		Do Until rMoneda.EOF 
'			i = ccur(i) + 1
	%>
			<tr bgcolor="#daf6ff">
				<td colSpan="0" rowSpan="0" >&nbsp;&nbsp;Bank Of America</td>
				<td ALIGN="right">2.500,00</td>
				<td ALIGN="left">Bank of Tokio</td>
				<td ALIGN="right">02-01-2002</td>
				<td ALIGN="middle">Enviada</td>
			</tr>
			<tr bgcolor="#daf6ff">
				<td colSpan="0" rowSpan="0" >&nbsp;&nbsp;Bank Of America</td>
				<td ALIGN="right">1.650,00</td>
				<td ALIGN="left">Boston Bank</td>
				<td ALIGN="right">05-02-2002</td>
				<td ALIGN="middle">Enviada</td>

			</tr>
		<%
'			rMoneda.MoveNext
'		Loop
		%>
</table>
</body></object></p></marquee></tr></tbody></table>
</body>
</html>
