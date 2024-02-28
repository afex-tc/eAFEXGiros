<%@ Language=VBScript %>
<!--#INCLUDE virtual="/Compartido/Errores.asp" -->
<!--#INCLUDE virtual="/Compartido/Rutinas.asp" -->
<%
	Dim sMonto
	Dim rsCl, sSQL
	Dim sFechaDesde, sFechaHasta, sCl
		
	sMonto = Request("mn") & " " & Request("mt")
	
	sFechaDesde = Request("fd")
	sFechaHasta = Request("fh")
	sCl = Request("cl")
	
	' JFMG 29-01-2010 se agregan los parámetros de rut o pasaporte
	dim sRut, sPasaporte
	sRut = Request("rt")
	sPasaporte = Request("ps")
	if sRut <> "" then sRut = right("000" & replace(replace(sRut, ".", ""), "-", ""), 9)	
	' ************** FIN JFMG 29-01-2010 ************************
	
	Set rsCl = CartolaCliente
	
	If Err.number <> 0 Then
		Set rsGiro = Nothing
		MostrarErrorMS "" 
	End If
	If rsCl.EOF Then
		Set rsCl = Nothing
		Response.Redirect "../Compartido/Error.asp?Description=No se encontró información del cliente"
	End If	

	Function CartolaCliente()
		' JFMG 29-01-2010 se agrega validación de rut o pasaporte
		if sCl <> "" then
			sSQL = "execute obtener_detalle_cartola '" & sFechaDesde & "', '" & sFechaHasta & "', '" & sCl & "'"
		
		elseif sRut <> "" then
			sSQL = "execute obtener_detalle_cartola '" & sFechaDesde & "', '" & sFechaHasta & "', null, '" & sRut & "'"
			
		elseif sPasaporte <> "" then
			sSQL = "execute obtener_detalle_cartola '" & sFechaDesde & "', '" & sFechaHasta & "', null, null, '" & sPasaporte & "'"
			
		else
			err.Raise -1,"CartolaCliente", "No se recibió ningún parámetro."
		end if
		' ********************* FIN JFMG 29-01-2009 **********************
		
		Set CartolaCliente = EjecutarSQLCliente(Session("afxCnxAFEXpress"), sSQL)
		
		'Si se produjeron errores en la consulta
		If Err.Number <> 0 Then
			MostarErrorMS "Lista de Giros"
		End If
	End Function
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
<table style="FONT-SIZE: 10pt; FONT-FAMILY: Verdana, Arial; PAGE-BREAK-BEFORE: auto" width=640>
	<tr>
		<td><font style="FONT-WEIGHT: bold; FONT-SIZE: 18pt">AFEX</font>
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			&nbsp;&nbsp;&nbsp;&nbsp;
			<IMG id=imgImprimir style="CURSOR: hand" height=20 src="../images/BotonImprimir.jpg" width=70 ><br>
			<font style="FONT-SIZE: 8pt; FONT-FAMILY: arial narrow, arial">CASA MATRIZ, MONEDA 1160 P.8 y 9 - SANTIAGO - CHILE<br>FONO(56 - 2)636 9000 - FAX(56 - 2)636 9071</font>
		</td>
	</tr>
<!--</table>-->
<table width=640>
<hr align=left width="100%" color=Silver size=-20px>
</table>
<table style="FONT-SIZE: 10pt; FONT-FAMILY: Verdana, Arial" width=640>
	<tr>
		<td align=center><font style="FONT-WEIGHT: bold; FONT-SIZE: 18pt">CARTOLA DE CLIENTES</font></td>
	</tr>
</table>
<table style="FONT-SIZE: 8pt; FONT-FAMILY: Verdana, Arial" width=640>
	<tr>
		<td colspan=3><font style="FONT-WEIGHT: bold; FONT-SIZE: 8pt">Nombre: </font><%=rsCl("nombre")%></td>
	</tr>
	<tr>
		<td><font style="FONT-WEIGHT: bold; FONT-SIZE: 8pt">Dirección: </font><%=rsCl("direccion")%></td>
		<td><font style="FONT-WEIGHT: bold; FONT-SIZE: 8pt">Ciudad: </font><%=rsCl("ciudad")%></td>
		<td><font style="FONT-WEIGHT: bold; FONT-SIZE: 8pt">Teléfono: </font><%=rsCl("telefono_completo")%></td>
	</tr>
	<tr>
		<td><font style="FONT-WEIGHT: bold; FONT-SIZE: 8pt">Fecha Creación: </font><%=rsCl("fecha_creacion")%></td>
		<td colspan=2><font style="FONT-WEIGHT: bold; FONT-SIZE: 8pt">Ult.Movto: </font><%=rsCl("fecha_ultmovto")%></td>
	</tr>
</table>
<br>
<table bgcolor=Black cellspacing="1" cellpadding="5" style="FONT-SIZE: 10pt; FONT-FAMILY: Verdana, Arial" width=640>
<!--<table cellspacing="1" cellpadding="5" STYLE="FONT-SIZE: 10px; COLOR: #707070; FONT-FAMILY: Verdana; POSITION: relative; TOP: 0px; HEIGHT: 30px;" bgcolor="#e1e1e1">-->
	<tr height="40" bgcolor="white">
		<td width=110 align=center><font style="FONT-WEIGHT: bold; FONT-SIZE: 8pt">Fecha</font></td>
		<td width=60 align=center><font style="FONT-WEIGHT: bold; FONT-SIZE: 8pt">Tipo<br>Giro</font></td>
		<td width=200 align=center><font style="FONT-WEIGHT: bold; FONT-SIZE: 8pt">Agente Captador</font></td>
		<td width=100 align=center><font style="FONT-WEIGHT: bold; FONT-SIZE: 8pt">Código<br>Captador</font></td>
		<td width=150 align=center><font style="FONT-WEIGHT: bold; FONT-SIZE: 8pt">Agente Pagador</font></td>
		<td width=100 align=center><font style="FONT-WEIGHT: bold; FONT-SIZE: 8pt">Código<br>Pagador</font></td>
		<td width=100 align=center><font style="FONT-WEIGHT: bold; FONT-SIZE: 8pt">Monto</font></td>
	</tr>
</table>
<table style="FONT-SIZE: 8pt; FONT-FAMILY: Verdana, Arial" width=640>
	<%	Dim nTotalMovtos, nTotalMonto
	
		nTotalMovtos = 0
		nTotalMonto = 0
		
		Do Until rsCl.EOF %>
			<tr height="20" bgcolor="white">
				<td width=110 align=center><%=rsCl("fecha_captacion")%></td>
				<td width=60><%=rsCl("tipo_giro")%></td>
				<td width=200><%=rsCl("captador")%></td>
				<td width=100><%=rsCl("invoice")%></td>
				<td width=150><%=rsCl("pagador")%></td>
				<td width=100><%=rsCl("codigo_pagador")%></td>
				<td width=100 align=right><%=FormatNumber(rsCl("monto"), 2)%></td>
			</tr>
	<%		
			nTotalMovtos = nTotalMovtos + 1
			nTotalMonto = nTotalMonto + cCur(rsCl("monto"))
			rsCl.movenext
		Loop
		Set rsCl = Nothing
	%>
</table>
<table width=640>
	<hr align=left width="100%" color=Black size=-20px>
</table>
<table border=0 style="FONT-SIZE: 8pt; FONT-FAMILY: Verdana, Arial" width=640>
	<tr>
		<td width=110>Total Movtos:</td>
		<td width=60></td>
		<td width=200><%=nTotalMovtos%></td>
		<td width=100></td>
		<td width=150></td>
		<td width=100></td>
		<td width=100 align=right><%=FormatNumber(nTotalMonto, 2)%></td>
	</tr>
</table>
</table>
	<!--
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
					<td><%=Request("bb")%></td>
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
					<td align=right>MESSAGE:</td>
					<td><%=Request("ms")%></td>
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
						pago nuevamente, se cobrará un cargo extra de USD 50 en caso reenvío y de USD 75 para investigaciones
						y ajustes. 
						</li>
						</menu>
						Si tiene dudas, reclamos o sugerencias, contáctenos a nuestro Teléfono de Información del Servicio
						de Transferencias <br>(56 - 2) 636 9024, al Fax (56 -2) 636 9071 o a nuestro correo electrónico transferencia@afex.cl
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
	
</table>-->
<BR>
</BODY>
</HTML>
