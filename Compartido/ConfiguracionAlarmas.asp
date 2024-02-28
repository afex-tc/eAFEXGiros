<%@ Language=VBScript %>
<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link rel="stylesheet" type="text/css" href="../Hoja%20de%20Estilos%201.css">
<script ID="clientEventHandlersJS" LANGUAGE="VBScript">
<!--



//-->
</script>
</head>
<script LANGUAGE="VBScript">
<!--

	Sub MostrarPaso(ByVal sPaso)
		window.tabPaso1.style.display = "none"
		window.tabPaso2.style.display = "none"
		window.tabPaso3.style.display = "none"
		'window.tabPaso4.style.display = "none"
		If sPaso="tabPaso1" Then
			window.tabPaso1.style.display = ""
		ElseIf sPaso="tabPaso2" Then
			window.tabPaso2.style.display = ""
		ElseIf sPaso="tabPaso3" Then
			window.tabPaso3.style.display = ""
		'Else
		'	window.tabPaso4.style.display = ""
		End If
	End Sub

	Sub CambiarCursor(Byval sControl)

		document.all.item(sControl).style.cursor = "Hand"
	
	End Sub
	
-->
</script>

<body>
<a NAME="Inicio" STYLE="POSITION: absolute; TOP: -10px"></a>
<marquee STYLE="HEIGHT: 132px; LEFT: 0px; POSITION: absolute; TOP: -28px; WIDTH: 407px" BEHAVIOR="slide" DIRECTION="right" SCROLLAMOUNT="30" SCROLLDELAY="50">
	<h6 STYLE="FONT-SIZE: 60pt">Servicios</h6>
</marquee>		
<marquee STYLE="HEIGHT: 73px; LEFT: 21px; POSITION: absolute; TOP: 1px; WIDTH: 406px" BEHAVIOR="slide" DIRECTION="down" SCROLLAMOUNT="8" SCROLLDELAY="100">		
	<h1 STYLE="COLOR: #cfcfcf">Configuración de Alarmas</h1>
</marquee>
<marquee STYLE="HEIGHT: 73px; LEFT: 20px; POSITION: absolute; TOP: 0px; WIDTH: 407px" BEHAVIOR="slide" DIRECTION="down" SCROLLAMOUNT="8" SCROLLDELAY="100">		
	<h1 STYLE="COLOR: steelblue">Configuración de Alarmas</h1>
</marquee>
<marquee STYLE="HEIGHT: 1090px; LEFT: 4px; POSITION: absolute; TOP: 32px; WIDTH: 573px" BEHAVIOR="slide" DIRECTION="up" SCROLLAMOUNT="120" SCROLLDELAY="1">		
	<!-- Alarma de Precios -->
	<table HEIGHT="300" ID="tabPaso1" CELLSPACING="0" CELLPADDING="1" BORDER="0" WIDTH="560" STYLE="BORDER-BOTTOM: silver 1px solid; BORDER-LEFT: silver 1px solid; BORDER-RIGHT: silver 1px solid; BORDER-TOP: silver 1px solid; LEFT: 2px; POSITION: relative; TOP: 16px">
		<img height="20" id="imgPaso1" name="imgPaso1" onclick="MostrarPaso('tabPaso1')" onmouseover="CambiarCursor('imgPaso1')" src="../images/TabPrecios.jpg" style="LEFT: 0px; POSITION: relative; TOP: 32px" width="97">
		<img border="0" height="20" id="imgPaso2" name="imgPaso2" onclick="MostrarPaso('tabPaso2')" onmouseover="CambiarCursor('imgPaso2')" src="../images/TabStock.jpg" style="LEFT: 0px; POSITION: relative; TOP: 32px" width="97">
		<img height="20" id="imgPaso3" name="imgPaso3" onclick="MostrarPaso('tabPaso3')" onmouseover="CambiarCursor('imgPaso3')" src="../images/TabNoticias.jpg" style="LEFT: 0px; POSITION: relative; TOP: 32px" width="97">
		<tr bgcolor=#dcf9ff>
			<td WIDTH="20"></td>
			<td HEIGHT="80" COLSPAN="6">
				 <h6 STYLE="HEIGHT: 89px; LEFT: 14px; POSITION: absolute; TOP: 38px; WIDTH: 257px; Z-INDEX: -1">Paso 1</h6>
				 Aquí poner indicaciones acerca de las alarmas
			</td>	
		</tr>
		<tr HEIGHT="15">
			<td colspan="7" bgcolor="steelblue" height="15"><font face="Verdana,Helvetica" color="white" size="2"><b>Alarma de Precios</b></font></td>
		</tr>
		<tr HEIGHT="20">
			<td COLSPAN="7"></td>		
		</tr>
		<tr HEIGHT="40">
			<td></td>
			<td><br><br><br>
				<input TYPE="checkbox" ID="chkPrecio1" NAME="chkPrecio1">
			</td>
			<td VALIGN="center" WIDTH="100">Qué operación desea hacer?<br><br>
				<select NAME="cbxMoneda" style="HEIGHT: 22px; WIDTH: 100px">
					<option SELECTED VALUE="1">Comprar</option>
					<option VALUE="2">Vender</option>
				</select>
			</td>
			<td>Seleccione la moneda<br><br><br>
				<select NAME="cbxMoneda" style="HEIGHT: 22px; WIDTH: 145px">
					<option SELECTED VALUE="USD">Dolar Americano</option>
					<option VALUE="EUR">Euro</option>
					<option VALUE="MKD">Marco Alemán</option>
				</select>
			</td>
			<td VALIGN="center" WIDTH="100">Tipo de Cambio<br>Actual<br><br>
				<input STYLE="TEXT-ALIGN: right" SIZE="10" ALIGN="right" ID="txt" NAME="txt" VALUE="660,56" DISABLED>
			</td>
			<td VALIGN="center" ALIGN="right">Digite el tipo de cambio para activar alarma<br><br>
				<input STYLE="TEXT-ALIGN: right" SIZE="10" ALIGN="rigth" ID="txt" NAME="txt" VALUE="660,56">
			</td>
			<td WIDTH="10"></td>
		</tr>
		<tr HEIGHT="50">
			<td></td>
			<td VALIGN="center">
				<input TYPE="checkbox" ID="chkPrecio2" NAME="chkPrecio2">
			</td>
			<td VALIGN="center" WIDTH="100">
				<select NAME="cbxMoneda" style="HEIGHT: 22px; WIDTH: 100px">
					<option SELECTED VALUE="1">Comprar</option>
					<option VALUE="2">Vender</option>
				</select>
			</td>
			<td>
				<select NAME="cbxMoneda" style="HEIGHT: 22px; WIDTH: 145px">
					<option SELECTED VALUE="USD">Dolar Americano</option>
					<option VALUE="EUR">Euro</option>
					<option VALUE="MKD">Marco Alemán</option>
				</select>
			</td>
			<td VALIGN="center">
				<input STYLE="TEXT-ALIGN: right" SIZE="10" ALIGN="right" ID="txt" NAME="txt" VALUE="660,56" DISABLED>
			</td>
			<td VALIGN="center" ALIGN="right">
				<input STYLE="TEXT-ALIGN: right" SIZE="10" ALIGN="rigth" ID="txt" NAME="txt" VALUE="660,56">
			</td>
			<td WIDTH="10"></td>
		</tr>

		<tr >
			<td COLSPAN="7"></td>		
		</tr>
		<img border="0" id="imgAceptarPrecios" onclick onmouseover="CambiarCursor('imgAceptarPrecios')" src="../images/BotonAceptar.jpg" style="LEFT: 227px; POSITION: relative; TOP: 294px" WIDTH="80" HEIGHT="25">
	</table>
	
	<!-- Alarma de Stock -->		
	<table HEIGHT="300" ID="tabPaso2" CELLSPACING="0" CELLPADDING="1" BORDER="0" WIDTH="560" STYLE="BORDER-BOTTOM: silver 1px solid; BORDER-LEFT: silver 1px solid; BORDER-RIGHT: silver 1px solid; BORDER-TOP: silver 1px solid; DISPLAY: none; LEFT: 2px; POSITION: relative; TOP: 16px">
		<img height="20" id="imgPaso21" name="imgPaso21" onclick="MostrarPaso('tabPaso1')" onmouseover="CambiarCursor('imgPaso21')" src="../images/TabPrecios.jpg" style="LEFT: 0px; POSITION: relative; TOP: 32px" width="97">
		<img border="0" height="20" id="imgPaso22" name="imgPaso22" onclick="MostrarPaso('tabPaso2')" onmouseover="CambiarCursor('imgPaso22')" src="../images/TabStock.jpg" style="LEFT: 0px; POSITION: relative; TOP: 32px" width="97">
		<img height="20" id="imgPaso23" name="imgPaso23" onclick="MostrarPaso('tabPaso3')" onmouseover="CambiarCursor('imgPaso23')" src="../images/TabNoticias.jpg" style="LEFT: 0px; POSITION: relative; TOP: 32px" width="97">
		<tr bgcolor="#dcf9ff">
			<td WIDTH="20"></td>
			<td HEIGHT="80" COLSPAN="6">
				 <h6 STYLE="HEIGHT: 89px; LEFT: 14px; POSITION: absolute; TOP: 38px; WIDTH: 257px; Z-INDEX: -1">Paso 1</h6>
				 Aquí poner indicaciones acerca de las alarmas
			</td>	
		</tr>
		<tr HEIGHT="15">
			<td colspan="7" bgcolor="steelblue" height="15"><font face="Verdana,Helvetica" color="white" size="2"><b>Alarma de Stock</b></font></td>
		</tr>
		<tr HEIGHT="20">
			<td COLSPAN="7"></td>		
		</tr>
		<tr HEIGHT="40">
			<td></td>
			<td><br><br>
				<input TYPE="checkbox" ID="chk" NAME="chk">
			</td>
			<td>Seleccione la moneda<br><br>
				<select NAME="cbxMoneda" style="HEIGHT: 22px; WIDTH: 145px">
					<option SELECTED VALUE="USD">Dolar Americano</option>
					<option VALUE="EUR">Euro</option>
					<option VALUE="MKD">Marco Alemán</option>
				</select>
			</td>
			<td VALIGN="center">Stock Actual<br><br>
				<input STYLE="TEXT-ALIGN: right" SIZE="10" ALIGN="right" ID="txt" NAME="txt" VALUE="25.200,56" DISABLED>
			</td>
			<td VALIGN="center" ALIGN="right">Digite el Monto para activar alarma<br><br>
				<input STYLE="TEXT-ALIGN: right" SIZE="10" ALIGN="rigth" ID="txt" NAME="txt" VALUE="25.200,56">
			</td>
			<td WIDTH="10"></td>
		</tr>
		<tr HEIGHT="40">
			<td></td>
			<td><br>
				<input TYPE="checkbox" ID="chk" NAME="chk">
			</td>
			<td><br>
				<select NAME="cbxMoneda" style="HEIGHT: 22px; WIDTH: 145px">
					<option SELECTED VALUE="USD">Dolar Americano</option>
					<option VALUE="EUR">Euro</option>
					<option VALUE="MKD">Marco Alemán</option>
				</select>
			</td>
			<td VALIGN="center"><br>
				<input STYLE="TEXT-ALIGN: right" SIZE="10" ALIGN="right" ID="txt" NAME="txt" VALUE="25.200,56" DISABLED>
			</td>
			<td VALIGN="center" ALIGN="right"><br>
				<input STYLE="TEXT-ALIGN: right" SIZE="10" ALIGN="rigth" ID="txt" NAME="txt" VALUE="25.200,56">
			</td>
			<td WIDTH="10"></td>
		</tr>
		<tr >
			<td COLSPAN="7"></td>		
		</tr>
		<img border="0" height="25" id="imgAceptarStock" onclick onmouseover="CambiarCursor('imgAceptarStock')" src="../images/BotonAceptar.jpg" style="LEFT: 227px; POSITION: relative; TOP: 294px" width="80">
	</table>

	<!-- Alarma de Noticias -->		
	<table HEIGHT="300" ID="tabPaso3" CELLSPACING="0" CELLPADDING="1" BORDER="0" WIDTH="560" STYLE="BORDER-BOTTOM: silver 1px solid; BORDER-LEFT: silver 1px solid; BORDER-RIGHT: silver 1px solid; BORDER-TOP: silver 1px solid; DISPLAY: none; LEFT: 2px; POSITION: relative; TOP: 16px">
		<img height="20" id="imgPaso31" name="imgPaso31" onclick="MostrarPaso('tabPaso1')" onmouseover="CambiarCursor('imgPaso31')" src="../images/TabPrecios.jpg" style="LEFT: 0px; POSITION: relative; TOP: 32px" width="97">
		<img border="0" height="20" id="imgPaso32" name="imgPaso32" onclick="MostrarPaso('tabPaso2')" onmouseover="CambiarCursor('imgPaso32')" src="../images/TabStock.jpg" style="LEFT: 0px; POSITION: relative; TOP: 32px" width="97">
		<img height="20" id="imgPaso33" name="imgPaso33" onclick="MostrarPaso('tabPaso3')" onmouseover="CambiarCursor('imgPaso33')" src="../images/TabNoticias.jpg" style="LEFT: 0px; POSITION: relative; TOP: 32px" width="97">
		<tr bgcolor="#dcf9ff">
			<td WIDTH="20"></td>
			<td HEIGHT="80" COLSPAN="6">
				 <h6 STYLE="HEIGHT: 89px; LEFT: 14px; POSITION: absolute; TOP: 38px; WIDTH: 257px; Z-INDEX: -1">Paso 1</h6>
				 Aquí poner indicaciones acerca de las alarmas
			</td>	
		</tr>
		<tr HEIGHT="15">
			<td colspan="5" bgcolor="steelblue" height="15"><font face="Verdana,Helvetica" color="white" size="2"><b>Alarma de Noticias</b></font></td>
		</tr>
		<tr  HEIGHT="20">
			<td COLSPAN="5"></td>		
		</tr>
		<tr  HEIGHT="20">
			<td></td>
			<td><b>Nombre del Producto</b></td>
			<td><b>Promociones</b></td>
			<td><b>Noticias</b></td>
			<td WIDTH="10"></td>
		</tr>
		<tr >
			<td></td>
			<td><input TYPE="checkbox" ID="chk" NAME="chk">Giros Nacionales</td>
			<td><input TYPE="checkbox" ID="chk" NAME="chk"></td>
			<td COLSPAN="2"><input TYPE="checkbox" ID="chk" NAME="chk"></td>
		</tr>
		<tr >
			<td></td>			
			<td><input TYPE="checkbox" ID="chk" NAME="chk">Giros Internacionales</td>
			<td><input TYPE="checkbox" ID="chk" NAME="chk"></td>
			<td COLSPAN="2"><input TYPE="checkbox" ID="chk" NAME="chk"></td>
		</tr>
		<tr >
			<td></td>
			<td><input TYPE="checkbox" ID="chk" NAME="chk">Transferencias Bancarias</td>
			<td><input TYPE="checkbox" ID="chk" NAME="chk"></td>
			<td COLSPAN="2"><input TYPE="checkbox" ID="chk" NAME="chk"></td>
		</tr>
		<tr >
			<td></td>
			<td><input TYPE="checkbox" ID="chk" NAME="chk">Travelers Cheques</td>
			<td><input TYPE="checkbox" ID="chk" NAME="chk"></td>
			<td COLSPAN="2"><input TYPE="checkbox" ID="chk" NAME="chk"></td>
		</tr>
		<tr HEIGHT="60">
			<td COLSPAN="5"></td>		
		</tr>
		<img border="0" height="25" id="imgAceptarNoticias" onclick onmouseover="CambiarCursor('imgAceptarNoticias')" src="../images/BotonAceptar.jpg" style="LEFT: 227px; POSITION: relative; TOP: 294px" width="80">
	</table>

</marquee>
</body>
</html>
