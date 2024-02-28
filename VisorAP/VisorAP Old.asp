<%@ Language=VBScript %>
<%
	'Response.Expires = 0
	'If Session("CodigoCliente") = "" Then
	'	response.Redirect "../Compartido/TimeOut.htm"
	'	response.end
	'End If
%>
<!--#INCLUDE FILE="../Compartido/Errores.asp" -->
<%
   Dim sMoneda, sFecha, sConexion
   
   'Inicio Variables para configurar la vista de VisorAP
   Dim VPr1, VPr2, VPr3				'Vista Principal
   Dim VPo1								'Vista Posicion
   Dim VUt1, VUt2						'Vista Utilidades   
   Dim VCo1, VCo2, VCo3				'Vista Configuracion
   
   'VPr1 = (1 = cInt(0 & Request("vpr1")))
   'VPr2 = (1 = cInt(0 & Request("vpr2")))
   'VPr3 = (1 = cInt(0 & Request("vpr3")))
   'VUt1 = (1 = cInt(0 & Request("vut1")))
   'VUt2 = (1 = cInt(0 & Request("vut2")))
   VPr1 = (1 = 1)
   VPr2 = (1 = 1)
   VPr3 = (1 = 0)
   VUt1 = (1 = 0)
   VUt2 = (1 = 0)
   'Fin Variables para configurar la vista de VisorAP
   
   
   sMoneda = Request("mn")
   If sMoneda = Empty Then sMoneda = "USD"

   sFecha = Request("fch")
   If sFecha = Empty Then sFecha = Date()
   
   sConexion = Request("cnx")
   If sConexion = Empty Then sConexion = "Provider=SQLOLEDB.1;Password=cambios;User ID=cambios;Initial Catalog=cambios;Data Source=canelo;"
   
   
'	If Not Then
'		Response.Redirect "http:../Compartido/Error.asp?Titulo=VisorAP&Description=..."
'		Response.End 
'	End If

%>

<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<title>Visor AP</title>
<link rel="stylesheet" type="text/css" href="../Estilos/Cliente.css">
<link rel="stylesheet" type="text/css" href="VisorAP.css">
</head>

<!--#INCLUDE FILE="Rutinas.htm" -->
<script LANGUAGE="VBScript">

   Dim rs
	On Error Resume Next
		
	Sub window_onload()
		window.setInterval "CargarVisor", 60000, "vbscript"
		CargarVisor
	End Sub
	
	Sub CargarVisor
		<% If VPr1 Or VPr2 Then %>
				Set rs = ObtenerVisorAP("<%=sConexion%>", "<%=sMoneda%>")				
		<% End If %>
		If Err.Number <> 0 Then 
		End If
		CargarTBVisor
		Set rs = Nothing		
	End Sub
		      
	Sub CargarTBVisor
		<% If VPr1 Then %>
				Cargar_tipocambio
		<% End If %>
		<% If VPr2 Then %>
				Cargar_estadistica
		<% End If %>		
		<% If VPr3 Then %>
				Cargar_posicion
		<% End If %>
	End Sub	

</script>

<body border="0" style="margin: 2 2 2 2" >
<form id="frmVisorAP" method="post">
<table id="tbVisor" border=0 cellpadding="0" cellspacing="0" style="border: 0px; margin: 0 0 0 0;">
	<% If VPr1 Then %>
			<tr><td><!--#INCLUDE FILE="_tipocambio.htm" --></td><tr>
	<% End If %>
	<% If VPr2 Then %>
			<tr><td><!-- #INCLUDE FILE="_estadistica.htm" --><td><tr>
	<% End If %>		
	<% If VPr3 Then %>
			<tr><td><!-- #INCLUDE FILE="_posicion.htm" --><td><tr>
	<% End If %>	
	<tr><td><!-- #INCLUDE FILE="_opciones.htm" --></td></tr>
</table>	
</form>
</body>
</html>