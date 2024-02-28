<%@ Language=VBScript %>
<!--#INCLUDE virtual="/Compartido/TimeOut.asp" -->
<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<title>Cargar Archivo TXT</title>
<link rel="stylesheet" type="text/css" href="../Estilos/Principal.css">
</head>

<body>
<script LANGUAGE="VBScript">
<!--
	Const sEncabezadoFondo = "Servicios"
	Const sEncabezadoTitulo = "Cargar Archivo"
	Const sClass = "TituloPrincipal"
	
	
	Sub aceptar_onclick()
		Dim sLinea, fs, f, s
		
		If Not CajaPregunta("AFEX En Linea", "Está seguro que desea cargar el archivo?") Then
			Exit Sub
		End If
				
		Set  fs = CreateObject("Scripting.FileSystemObject")
		Set f = fs.OpenTextFile(Prueba.fileadjunto.value, 1, False)
		sLinea = f.readall		
		Prueba.contenido.value = slinea
		f.Close
		<% If Session("ModoPrueba") Then %>
				Prueba.action = "CargarGiroTXT.asp"
		<% Else %>
				Prueba.action = "CargarGiroTXT.asp"
		<% End If %>
		Prueba.submit
		Prueba.action=""		
	End Sub
	
-->
</script>
<table cellspacing="0" cellpadding="0" border="0" width="100%" height="40px" style="position: relative; left: 0px; top: 0px">
<tr><td valign="top" class="TituloFondo">
		<script language="VBScript">
		<!--
			document.write sEncabezadoFondo
		//-->
		</script>
		<h1 class="sombra" STYLE="LEFT: 21px; POSITION: absolute; TOP: 1px;"><script>document.write sEncabezadoTitulo</script></h1>
		<h1 id="hTituloPrincipal" STYLE="LEFT: 20px; POSITION: absolute; TOP: 0px;">
			<script language="VBScript"> 
			<!--
						hTituloPrincipal.className = sClass
						document.write sEncabezadoTitulo
			//-->
			</script>
	</h1>
</td></tr>
</table>

<form id="Prueba" method="post">
	<table border="0">
	<tr>
		<td style="font-size: 10pt">Ingrese el nombre completo del archivo de texto y haga clic en aceptar.</td>
	</tr>
	<tr height="20"><td></td></tr>
	<tr>		
		<td>Archivo<br><input type="file" name="fileadjunto" id="fileadjunto" size="50"></td>
		<td><br><input type="hidden" name="contenido" id="contenido" value></td>
		<td><br><input type="button" src="nn.jpg" id="aceptar" value="aceptar"></td>
	</tr>
</form>
</body>
</html>
<!--#INCLUDE virtual="/Compartido/Rutinas.htm" -->
