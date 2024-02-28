<HTML>

<HEAD>
	

	<TITLE>Corregir Beneficiario</TITLE>
	<LINK rel="stylesheet" type="text/css" href="../Estilos/Principal.css">
</HEAD>
<BODY>
<script LANGUAGE="VBScript">
<!--
	Const sEncabezadoFondo = "Beneficiario"
	Const sEncabezadoTitulo = "Corregir Información"
	Const sClass = "TituloPrincipal"
      
-->
</script>
<%
dim Cod_Giro,Monto,Detalle
Dim sSQL4,rs4,cnn
dim invoice,codi,titulo

Cod_Giro=request.QueryString("Giro")
Monto=request.QueryString("Monto")
Detalle=request.QueryString ("Detalle")

	On Error Goto 0

		'On Error Resume Next		
			Set Cnn = server.CreateObject("ADODB.Connection")		
			Cnn.CommandTimeout = 60
			Cnn.Open Session("afxCnxAFEXpress")						
			Set rs4 = server.CreateObject("ADODB.Recordset") 
			'response.Write bs
			sSQL4="SELECT * FROM GIRO WHERE codigo_giro= '"& Cod_giro &"' "
			rs4.Open sSQL4,cnn, 3, 1 
			
			invoice=(rs4.Fields("invoice").Value )
				
			if (trim(invoice)<>empty) then
				codi=invoice
				titulo="Invoice"
			else
				codi=cod_giro
				titulo="Codigo Giro"
			end if	
			
			Set rs4 = Nothing
			Cnn.Close
			Set Cnn = Nothing
						
%>

</B>
<!--#INCLUDE virtual="/Compartido/Encabezado.htm" -->
<!--#INCLUDE virtual="/Compartido/Rutinas.htm" -->
<!--
<MARQUEE STYLE="HEIGHT: 396px; LEFT: 0px; POSITION: relative; TOP: 0px; WIDTH: 456px" BEHAVIOR=slide DIRECTION=up SCROLLAMOUNT=50 SCROLLDELAY=100>		
-->
<br>
	<form name="Corregir" method="post" onsubmit="validar(this)">
		<input type="hidden" name="Cod_Giro" value="<%=Cod_Giro%>">
		<input type="hidden" name="Monto" value="<%=Monto%>">
		<input type="hidden" name="Detalle" value="<%=Detalle%>">
	<table border="0">
	<tr>
		<TD WIDTH="10%"></TD>		
		<TD WIDTH="90%">
		<table border="0">
			<tr>
				<td>
				<%=titulo%> :<b> <%=codi%></b>
				</td>
			</tr>
			<tr>
				<td>
				 Monto		: <%=Monto%> Dolares
				</td>
			</tr>	
			<tr>
				<td>
				  <%=Detalle%>
				</td>
			</tr>	
			<tr>
			<td></td>
			</tr>
			<tr>
			<td></td>
			</tr>
			<tr>
					
			<td>
				<font face="verdana" size="2">Nombre Sucursal o Agente</font><br>
				<input type="text" name="txtNombre" MAXLENGTH=50 ID=txtNombre OnBlur="Corregir.txtNombre.value=MayusculaMinuscula(Corregir.txtNombre.value)"></input><br>
				<font face="verdana" size="2">Email de Contacto</font><br>
				<input type="text" name="txtEmail" MAXLENGTH=50 ID=txtEmail></input>
			</td>					
			</tr>	
			<tr>
				<td></td>
			</tr>	
			<tr>
			<td>
			<input TYPE="checkbox" name="chkNombre" value="Nombre">Nombre o Apellido <br>
			<input TYPE="checkbox" name="chkDireccion"  value="Direccion">Direccion<br>			
			<input TYPE="checkbox" name="chkFono" value="Fono">Fono<br>
			</td>
			</tr>
			<tr>			
				<td>
					<font face="verdana" size="2"> Escriba Datos Correctos<br>
					<textarea name="txtCorregir" rows="10" cols="50" ></textarea>				
					</font>
				</td>
			</tr>
				<tr>
					<td>
						<table align="center" border="0">
							<tr >
								<td><INPUT TYPE="button" ID="cmdEnviar" VALUE=" Enviar "></INPUT></td>
								<td><INPUT TYPE="button" ID="cmdLimpiar" VALUE=" Limpiar "></INPUT></td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
		</td>
	</tr>	
	</table>
	</form>
	
</BODY>
<script language="vbScript">
	
		Sub cmdEnviar_onClick()
	
			If Not CajaPregunta("AFEX En Linea", "Está seguro que desea Modificar la información del Beneficiario?") Then
				LimpiarControles
				exit sub					
			else	
			 	Corregir.action = "EnviarCorreccion.asp"
				Corregir.submit
				Corregir.action = ""
			 end if	
		
		End Sub
		
		Sub LimpiarControles			
			Corregir.txtNombre.value = ""
			Corregir.txtEmail.value = ""
			Corregir.txtCorregir.value = ""		
		End Sub
		
		Sub cmdLimpiar_onclick()
			LimpiarControles
		End Sub
</script>
</HTML>
