<%@ Language=VBScript %>
<%
	' variables para descomponer el MoneySaver
	dim sMS1
	dim sMS2
	dim sMS3
	dim sMS4
	dim sMoneySaver

	' verifica si yá había ingresado un MoneySaver
	sMoneySaver = request("MS")
	if sMoneySaver <> empty then 
		sMS1 = left(sMoneySaver, 3)
		sMS2 = mid(sMoneySaver, 4, 3)
		sMS3 = mid(sMoneySaver, 7, 3)
		sMS4 = right(sMoneySaver, 3)
	end if

%>

<html>
<head>
<title>MoneySaver</title>
</head>
	<body background="imagenes/Giros_FondoVentana.jpg" text="#008080" link="#0000FF" vlink="#000080">
	<script language="VBScript">
	<!--
		sub Aceptar_onClick()
			window.returnvalue = trim(ms1.value) & trim(ms2.value) & trim(ms3.value) & trim(ms4.value)
			window.close
		end sub
		sub Cancelar_onClick()
			window.returnvalue = ""
			window.close
		end sub
	-->
	</script>

	<br>
	<center>
		<font FACE="Verdana" SIZE="6"><b><div ALIGN="CENTER"></div></b></font>
		<table border="0" width="20%">
			<tr>
				<td>				
					<input type="text" name="MS1" STYLE="font-family: Verdana; font-size: 12" size="3" value="<%=sMS1%>" maxlength="3"></input>&nbsp;
					<input type="text" name="MS2" STYLE="font-family: Verdana; font-size: 12" size="3" value="<%=sMS2%>" maxlength="3"></input>&nbsp;
					<input type="text" name="MS3" STYLE="font-family: Verdana; font-size: 12" size="3" value="<%=sMS3%>" maxlength="3"></input>&nbsp;
					<input type="text" name="MS4" STYLE="font-family: Verdana; font-size: 12" size="3" value="<%=sMS4%>" maxlength="3"></input>
				</td>
			</tr>
		</table>
		<br>
		<input type="button" name="Aceptar" STYLE="font-family: Verdana; font-size: 12" value="Aceptar"></input>
		<input type="button" name="Cancelar" STYLE="font-family: Verdana; font-size: 12" value="Cancelar"></input>
	</center>
</body>
</html>