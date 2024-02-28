<%@ Language=VBScript %>
<%
	Response.Expires = 0
	Response.buffer = true	
	Response.expires = 0
	Response.expiresabsolute = Now() - 1
	Response.addHeader "pragma", "no-cache"
	Response.addHeader "cache-control", "private"
	Response.CacheControl = "no-cache"

%>
<!--#INCLUDE virtual="/Compartido/Errores.asp" -->
<!--#INCLUDE virtual="/VisorAP/Rutinas.asp" -->
<%
	If Cint(0 & Session("AutorizaVisorAP")) <> 1 Then
		Response.Write "Acceso no autorizado<br>"
		Response.Write "Comun�quese con AFEX o ingrese a www.afex.cl"
		response.End 		
	End If
	
   Dim sMoneda, sFecha, sConexion
   
   'Inicio Variables para configurar la vista de VisorAP
   Dim VPr1, VPr2, VPr3				'Vista Principal
   Dim VPo1								'Vista Posicion
   Dim VUt1, VUt2						'Vista Utilidades   
   Dim VCo1, VCo2, VCo3				'Vista Configuracion
   
   VPr1 = (1 = cInt(0 & Request("vpr1")))
   VPr2 = (1 = cInt(0 & Request("vpr2")))
   'VPr3 = (1 = cInt(0 & Request("vpr3")))
   'VUt1 = (1 = cInt(0 & Request("vut1")))
   'VUt2 = (1 = cInt(0 & Request("vut2")))
   'VPr1 = (1 = 1)
   'VPr2 = (1 = 1)
   VPr3 = (1 = 0)
   VUt1 = (1 = 0)
   VUt2 = (1 = 0)
   'Fin Variables para configurar la vista de VisorAP
   
   
   sMoneda = Request("mn")
   If sMoneda = Empty Then sMoneda = "USD"

   sFecha = Request("fch")
   If sFecha = Empty Then sFecha = Date()
   
   sConexion = Request("cnx")
   If sConexion = Empty Then sConexion = Session("afxCnxAFEXchange")
   
   Dim  rs
   CargarVisor 
   
	Sub CargarVisor
		Set rs = ObtenerVisorAP(sConexion, sMoneda)				
		If Err.Number <> 0 Then 
		End If
		'CargarTBVisor
		'Set rs = Nothing		
	End Sub
	   

%>

<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<title>Visor AP</title>
<link rel="stylesheet" type="text/css" href="../Estilos/Cliente.css">
<link rel="stylesheet" type="text/css" href="VisorAP.css">
</head>

<script LANGUAGE="VBScript">

   Dim rs
	On Error Resume Next
		
	Sub window_onload()
		'window.setInterval "CargarVisor", 10000, "vbscript"
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
			'Cargar_posicion
		<% End If %>
	End Sub	
	
	Sub CargarVisor
		window.navigate "http:VisorAP.asp?mn=<%=sMoneda%>"
	End Sub

</script>
<body border="0" style="margin: 0 0 0 0" >
<form id="frmVisorAP" method="post">
<table id="tbVisor" border=0 cellpadding="0" cellspacing="0" style="border: 0px; margin: 0 0 0 0;">
	<% If VPr1 Then %>
			<tr><td><!--#INCLUDE virtual="/visorap/_tipocambio.htm" --></td><tr>
	<% End If %>
	<% If VPr2 Then %>
			<tr><td><!--#INCLUDE virtual="/visorap/_estadistica.htm" --></td><tr>
	<% End If %>
</table>	
</form>
</body>
</html>