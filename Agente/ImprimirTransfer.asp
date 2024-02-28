<%@ Language=VBScript %>
<%
	Dim sMonto
	
	sMonto = Request("mn") & " " & Request("mt")
%>
<HTML>
<HEAD>
<META name="VI60_DefaultClientScript" Content="VBScript">

<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<BODY>
<SCRIPT LANGUAGE=vbscript>
<!--

	Sub imgImprimir_OnClick()
		imgimprimir.style.display = "none"
		window.print()
		imgimprimir.style.display = ""
	End Sub
-->
</SCRIPT>
<table style="font-family: Verdana, Arial; font-size: 10pt" width=640px>
	<tr>
		<td><font style="font-weight: bold; font-size: 18pt">AFEX</font>
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			&nbsp;&nbsp;&nbsp;&nbsp;
			<img Id="imgImprimir" src="../images/BotonImprimir.jpg" style="CURSOR: hand" WIDTH="70" HEIGHT="20"><br>
			<font style="font-family: arial narrow, arial; font-size: 8pt">CASA MATRIZ, MONEDA 1160 P.8 y 9 - SANTIAGO - CHILE<br>FONO(56 - 2)636 9000 - FAX(56 - 2)636 9071</font>
		</td>
	</tr>
	
	<tr>
		<td style="background-color: silver"><font style="font-size: 14pt; font-weight: bold">Transferencia</font></td>
	</tr>
	<tr>
		<td>
			<table width=640px style="border: 1px solid silver; font-family: Verdana, Arial; font-size: 10pt">
				<tr>
					<td colspan=3 style="background-color: silver">
						<font style="font-weight: bold">PARA</font>
					</td>
				</tr>
				<tr>
					<td width=50px></td>
					<td align=right width=100px>PARA:</td>
					<td width=490px><%=Request("pr")%></td>
				</tr>
				<tr>
					<td></td>
					<td align=right>DESDE:</td>
					<td><%=Request("dd")%></td>
				</tr>
				<tr>
					<td></td>
					<td align=right>FECHA:</td>
					<td><%=Request("fc")%></td>
				</tr>
				<tr height=0px>
					<td></td>
					<td></td>
					<td></td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td>
			<table width=640px style="border: 1px solid silver; font-family: Verdana, Arial; font-size: 10pt">
				<tr>
					<td colspan=3 style="background-color: silver">
						<font style="font-weight: bold">CLIENTE</font>
					</td>
				</tr>
				<tr>
					<td width=50px></td>
					<td align=right width=100px>NOMBRE:</td>
					<td width=490px><%=Request("nb")%></td>
				</tr>
				<tr>
					<td></td>
					<td align=right>RUT:</td>
					<td><%=Request("rt")%></td>
				</tr>
				<tr>
					<td></td>
					<td align=right>TELEFONO:</td>
					<td><%=Request("tl")%></td>
				</tr>
				<tr>
					<td></td>
					<td align=right>FAX:</td>
					<td><%=Request("fx")%></td>
				</tr>
				<tr>
					<td></td>
					<td align=right>CONTACTO:</td>
					<td><%=Request("co")%></td>
				</tr>
				<tr height=0px>
					<td></td>
					<td></td>
					<td></td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td>
			<table width=640px style="border: 1px solid silver; font-family: Verdana, Arial; font-size: 10pt">
				<tr>
					<td colspan=3 style="background-color: silver">
						<font style="font-weight: bold">DATOS DE LA TRANSFERENCIA</font>
					</td>
				</tr>
				<tr>
					<td width=50px></td>
					<td align=right width=180px>AMOUNT:</td>
					<td width=410px><%=sMonto%></td>
				</tr>
				<tr>
					<td></td>
					<td align=right>BENEFICIARY'S BANK:</td>
					<td><%=replace(Request("bb")," @ ", "&")%></td>
				</tr>
				<tr>
					<td></td>
					<td align=right>ADRESS BANK:</td>
					<td><%=Request("ab")%></td>
				</tr>
				<tr>
					<td></td>
					<td align=right>FW/ACC:</td>
					<td><%=Request("fa")%></td>
				</tr>
				<tr>
					<td></td>
					<td align=right>RELEASE DATE:</td>
					<td><%=Request("rd")%></td>
				</tr>
				<tr>
					<td></td>
					<td align=right>VALUE DATE:</td>
					<td><%=Request("vd")%></td>
				</tr>
				<tr>
					<td></td>
					<td align=right>S/REF./IND.ID.:</td>
					<td><%=Request("rf")%></td>
				</tr>
				<tr>
					<td></td>
					<td align=right>BENEFICIARY:</td>
					<td><%=Request("bf")%></td>
				</tr>
				<tr>
					<td></td>
					<td align=right>ACCOUNT:</td>
					<td><%=Request("ct")%></td>
				</tr>
				<tr>
					<td></td>
					<td align=right>INTERMEDIARY BANK:</td>
					<td><%=Request("ib")%></td>
				</tr>
				<tr>
					<td></td>
					<td align=right>ACCOUNT:</td>
					<td><%=Request("ca")%></td>
				</tr>
				<tr>
					<td></td>
					<td align=right>INVOICE:</td>
					<td><%=Request("iv")%></td>
				</tr>
				<tr>
					<td></td>
					<td align=right>CRÉDITO A:</td>
					<td><%=Request("ft")%></td>
				</tr>
				<tr>
					<td></td>
					<td align=right>MESSAGE:</td>
					<td><%=Request("ms")%></td>
				</tr>
				<tr>
					<td></td>
					<td align=right>BENEFICIARY'S ADRESS:</td>
					<td><%=Request("db")%></td>
				</tr>
				<tr height=0px>
					<td></td>
					<td></td>
					<td></td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td>
			<table width=640px style="text-align: justify; border: 1px solid silver; font-family: Verdana, Arial; font-size: 10pt">
				<tr>
					<td colspan=3 style="background-color: silver">
						<font style="font-weight: bold">CONDICIONES</font>
					</td>
				</tr>
				<tr>
					<td style="text-align: justify; font-family: Arial; font-size: 8pt">
						<menu>
						<li>El cliente declara que el origen de los fondos no son producto de actividades ilícitas y que
						el destino de la remesa de esta Transferencia, tampoco tienen relación con alguna actividad ilegal.
						</li>
						<li> El cliente manifiesta que los datos proporcionados a AFEX para realizar la transferencia, y que
						consta en este documento, son fidedignos. Queda AFEX por lo tanto, liberado de toda responsabilidad
						en caso de que esta transferencia no sea acreditada en las condiciones y plazos estipulados, cuya
						causa fuese un error u omisión de información por parte del cliente.
						</li>
						<li>En caso de que sea necesario corregir los datos proporcionados por el cliente para efectuar el
							 pago nuevamente, se cobrara un cargo extra de USD 100, lo mismo se cobrara en caso de investigaciones y devolución de fondos. 
						</li>
						</menu>
						Si tiene dudas, reclamos o sugerencias, contáctenos a nuestro Teléfono de Información del Servicio
						de Transferencias <br>(56 - 2) 636 9024, al Fax (56 -2) 636 9071 o a nuestro correo electrónico transferencias@afex.cl
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td>
		<table style="font-family: Verdana, Arial; font-size: 8pt" width=650px>
			<tr height=30px><td></td></tr>
			<tr>
				<td align=center width=50%>
					---------------------------------<br>
					Firma Cliente
				</td>
				<td align=center width=50%>
					---------------------------------<br>
					Firma <%=Request("dd")%>
				</td>
			</tr>
			<tr height=20px><td></td></tr>
			<tr>
				<td>
					Nombre:_________________________________________<br><br>
					Rut &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;:_________________________<br>
				</td>
			</tr>
			<tr></tr>
			<tr></tr>
		</table>
		</td>
	</tr>
</table>
</BODY>
</HTML>
