<%@ Language=VBScript %>
<%
	'Expiracion
	'Se asegura que la página no se almacene en la memoria cache
	Response.Expires = 0
	If Session("CodigoCliente") = "" Then
		response.Redirect "TimeOut.htm"
		response.end
	End If

	'Variables
	Dim sVigencia, rMoneda
		
	'Obtener Paridades
	sSQL = " execute MostrarParidadesMantenedor "
	Set rMoneda = EjecutarSQLCliente(Session("afxCnxAfexchange"), sSQL)
	if err.number <> 0 then
		set rMoneda = nothing
		mostrarerrorms "Cargar Paridades Monedas Mantenedor"
	end if	
	
%>
<!--#INCLUDE virtual="/compartido/Transferencias.asp" -->
<!--#INCLUDE virtual="/compartido/Rutinas.asp" -->
<!--#INCLUDE virtual="/compartido/Errores.asp" -->
<html>
<head>
<title>Paridades</title>
<link rel="stylesheet" type="text/css" href="../Estilos/Cliente.css">
</head>

<script language="vbscript">
<!--
	sub Imprimir()
		OLECMDID_PRINT = 6 
		OLECMDEXECOPT_DONTPROMPTUSER = 2 
		OLECMDEXECOPT_PROMPTUSER = 1 
		'ACA en caso de usar frames, 
		'enfocamos el frame a imprimir: 
		
		'window.printo.value = Cadena
		
		'window.parent.frames.main.document.body.focus() 
		window.document.body.focus() 
		
		'Llamamos al comando de Impresión Print 
		
		
		on error resume next 
		call IEWB.ExecWB (OLECMDID_PRINT, -1) 		
		
		if err.number <> 0 then 
			alert "No se pudo imprimir" & err.Description 
		end if 
	end sub
-->
</script>
<body background="../agente/imagenes/Giros_FondoVentana.jpg" bgcolor="#FFFFFF" text="#008080" link="#0000FF" vlink="#000080">

<object id="IEWB" width="0" height="0" classid="clsid:8856F961-340A-11D0-A96B-00C04FD705A2" VIEWASTEXT></object> 

<script LANGUAGE="VBScript">
<!--
	Const sEncabezadoFondo = "Información"
	Const sEncabezadoTitulo = "Paridades"
	Const sClass = "TituloPrincipal"
-->
</script>
	
<!--#INCLUDE virtual="/compartido/Encabezado.htm" --> 
		<br>
		<table ID="tbParidades" width="100%" border="0.5" ALIGN="CENTER" STYLE="font-family: Verdana; font-size: 13; color:#505050">
			<tr BORDER="1">
				<td WIDTH="25%" ALIGN="CENTER" BGCOLOR="#EAD094" STYLE="font-family: Verdana; font-size: 14">
					<b>Moneda</b>
				</td>
				<td WIDTH="20%" ALIGN="CENTER" BGCOLOR="#FCF6E9" STYLE="font-family: Verdana; font-size: 14">
					<b>Paridad</b>
				</td>				
			</tr>
			<%
				i = 0
				Do Until rMoneda.EOF 
					i = ccur(i) + 1
			%>
					<tr>
						<td colSpan="0" rowSpan="0" BGCOLOR="#F0E4C0">&nbsp;&nbsp;<%=rMoneda("Alias_Moneda")%></td>
						<td ID="<%="PO" & i%>" BGCOLOR="#FBF7F0" ALIGN="RIGHT"><%=rMoneda("Paridad")%>&nbsp;&nbsp;</td>						
					</tr>
				<%
					rMoneda.MoveNext
				Loop
				%>
		</table>
		
				
		<%	
		Set rMoneda = Nothing
		%>
		
		<!-- se muestran los botones -->
		<center>
			<!--<a href="javascript:history.back()"><img src="../agente/imagenes/GIROS_Anterior.jpg" border="0" 			alt="Volver a la página Anterior" WIDTH="99" HEIGHT="80"></img></a>-->
			<input TYPE="image" onClick="Imprimir" SRC="../images/BotonImprimir.jpg" border="0" alt="Imprimir las Paridades" WIDTH="70" HEIGHT="20">
		</center>
	</body>
</html>