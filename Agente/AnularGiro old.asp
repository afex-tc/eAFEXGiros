<%@ Language=VBScript %>
<%
dim Cod_Giro,Monto,Detalle
Dim sSQL4,rs4,cnn
dim invoice

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
			
			Set rs4 = Nothing
			Cnn.Close
			Set Cnn = Nothing
						
	%>	
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<TITLE>Anular Giro</TITLE>
<LINK rel="stylesheet" type="text/css" href="../Estilos/Principal.css">
</HEAD>
<BODY>
<script LANGUAGE="VBScript">
<!--
	Const sEncabezadoFondo = "Anular Giro"
	Const sEncabezadoTitulo = "Anular Giro"
	Const sClass = "TituloPrincipal"
	
-->
</script>
<!--#INCLUDE virtual="/Compartido/Encabezado.htm" -->
<!--#INCLUDE virtual="/Compartido/Rutinas.htm" -->

<!--
<MARQUEE STYLE="HEIGHT: 396px; LEFT: 0px; POSITION: relative; TOP: 0px; WIDTH: 456px" BEHAVIOR=slide DIRECTION=up SCROLLAMOUNT=50 SCROLLDELAY=100>		
-->
<br>
	<form name="Anula" method="post">
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
				Invoice :<b> <%=Invoice%></b>
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
				<td >
				<font face="verdana" size="2">Nombre Sucursal o Agente</font><br>
				<input type="text" name="txtNombre" MAXLENGTH=50 ID=txtNombre OnBlur="Anula.txtNombre.value=MayusculaMinuscula(Anula.txtNombre.value)"></input><br>
				<font face="verdana" size="2">Email de Contacto</font><br>
				<input type="text" name="txtEmail" MAXLENGTH=50 ID=txtEmail></input>
				</td>					
			</tr>	
			<tr>
				<td></td>
			</tr>	
			<tr>			
				<td>
					<font face="verdana" size="2"> Motivo de Anulación<br>
					<textarea name="txtAnular" rows="10" cols="50" ></textarea>				
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
	
			If Not CajaPregunta("AFEX En Linea", "Está seguro que desea Anular este Giro?") Then
				LimpiarControles
				exit sub					
			 else	
			 	Anula.action = "EnviarAnulacion.asp"
				Anula.submit
				Anula.action = ""
			 end if	
		
		End Sub
		
		Sub LimpiarControles
			
			Anula.txtNombre.value = ""
			Anula.txtEmail.value = ""
			Anula.txtAnular.value = ""
		
		End Sub
		
		Sub cmdLimpiar_onclick()
			LimpiarControles
		End Sub
</script>
</HTML>
