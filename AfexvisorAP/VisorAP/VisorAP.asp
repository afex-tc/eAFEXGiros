<%@ Language=VBScript %>
<!--#INCLUDE virtual="/afexvisorap/Compartido/_timeout.asp" -->
<!--#INCLUDE virtual="/afexvisorap/visorap/_sucursales.asp" -->
<%

	Dim i, sSucursal, sNombreSC
   sSucursal = Request("sc")
   If sSucursal = Empty Then sSucursal = Session("CodigoSC")
   If sSucursal = Empty Then sSucursal = "AW"
	
	For i = 0 To 2
		If aCodigoSC(i) = sSucursal Then
			Session("cnxVisorAP") = aConexionSC(i)
			Session("CodigoSC") = aCodigoSC(i)
			Session("NombreSC") = aNombreSC(i)
			Session("AliasSC") = Trim(Replace(aNombreSC(i), "Afex", ""))
		End If		
	Next
	
	
   Dim sMoneda, sFecha, sConexion, sEjecutivo, sNombreEjecutivo, sCheckMonedaUti
   Dim sCodigoInternacional
   
   'Inicio Variables para configurar la vista de VisorAP
   Dim VLogo							'Logo encabezado
   Dim VPromo, nPromos				'Promociones
   Dim VNtc								'Noticias
   Dim VPr1, VPr2, VPr3				'Vista Principal
   Dim VPo1								'Vista Posicion
   Dim VPo11
   Dim VUt1, VUt2						'Vista Utilidades
   Dim VUt11, VUt12, VUt13			
   Dim VUt121
   Dim VUt21, VUt22, VUt23
   Dim VUt221
   Dim VMn1								'Vista monedas
   Dim VCo1								'Vista Configuracion   
   Dim VCo11, VCo12, VCo13
   Dim VCl1								'Vista Clientes
   Dim VCd1								'Vista Cierre Diario
   Dim nItv								'Intervalo actualizacion Visor. Defecto 10000 (10 segundos)
   Dim nItvUt							'Intervalo actualizacion utilidades. Defecto 6
   Dim nItvLV							'Intervalo limpiar Visor

   sMoneda = Request("mn")
   If sMoneda = Empty Then sMoneda = "USD"

   sFecha = Request("fch")
   If sFecha = Empty Then sFecha = Date()
   
   sConexion = Request("cnx")
   If sConexion = Empty Then sConexion = Session("cnxVisorAP")

   sEjecutivo = Request("ej")
   If sEjecutivo = Empty Then sEjecutivo = "MGOUDIE"

   sNombreEjecutivo = Request("nej")
   If sNombreEjecutivo = Empty Then sNombreEjecutivo = "Mariana Goudie"

   nItv = cCur(0 & Request("itv"))
   If nItv = 0 Then nItv = 20000

   nPromos = cInt(0 & Request("prms"))
   If nPromos = 0 Then nPromos = 2

   nItvUt = cInt(0 & Request("itvut"))
   If nItvut = 0 Then nItvUt = 6

   nItvLV = cInt(0 & Request("itvlv"))
   If nItvLV = 0 Then nItvLV = 60

   If Request("mut")= "on" Then 
		sCheckMonedaUti = "checked"
		sMonedaUti = sMoneda
	End If
   

   VLogo = (Instr(Session("StringOpciones"), ";logo;") <> 0)
   VPromo = (Instr(Session("StringOpciones"), ";prm;") <> 0)
   VNtc = (Instr(Session("StringOpciones"), ";Ntc;") <> 0)
   VPr1 = (Instr(Session("StringOpciones"), ";pr1;") <> 0)
   VPr2 = (Instr(Session("StringOpciones"), ";pr2;") <> 0)
   VPr3 = (Instr(Session("StringOpciones"), ";pr3;") <> 0)
   VPo1 = (Instr(Session("StringOpciones"), ";po1;") <> 0)
   VPo11 = (Instr(Session("StringOpciones"), ";po11;") <> 0)
   VUt1 = (Instr(Session("StringOpciones"), ";ut1;") <> 0)
   VUt11 = (Instr(Session("StringOpciones"), ";ut11;") <> 0)
   VUt12 = (Instr(Session("StringOpciones"), ";ut12;") <> 0)
   VUt121 = (Instr(Session("StringOpciones"), ";ut121;") <> 0)
   VUt13 = (Instr(Session("StringOpciones"), ";ut13;") <> 0)
   VUt2 = (Instr(Session("StringOpciones"), ";ut2;") <> 0)
   VUt21 = (Instr(Session("StringOpciones"), ";ut21;") <> 0)
   VUt22 = (Instr(Session("StringOpciones"), ";ut22;") <> 0)
   VUt221 = (Instr(Session("StringOpciones"), ";ut221;") <> 0)
   VUt23 = (Instr(Session("StringOpciones"), ";ut23;") <> 0)
   VMn1 = (Instr(Session("StringOpciones"), ";mn1;") <> 0)
   VCo1 = (Instr(Session("StringOpciones"), ";co1;") <> 0)
   VCo11 = (Instr(Session("StringOpciones"), ";co11;") <> 0)
   VCo12 = (Instr(Session("StringOpciones"), ";co12;") <> 0)
   VCo13 = (Instr(Session("StringOpciones"), ";co13;") <> 0)
   VCl1 = (Instr(Session("StringOpciones"), ";cl1;") <> 0)
   VCd1 = (Instr(Session("StringOpciones"), ";cd1;") <> 0)
   
   'VLogo = (1 = cInt(0 & Request("vlogo")))
   'VPromo = (1 = cInt(0 & Request("vpromo")))
   'VNtc = (1 = cInt(0 & Request("vntc")))
   'VPr1 = (1 = cInt(0 & Request("vpr1")))
   'VPr2 = (1 = cInt(0 & Request("vpr2")))
   'VPr3 = (1 = cInt(0 & Request("vpr3")))
   'VUt1 = (1 = cInt(0 & Request("vut1")))
   'VUt11 = (1 = cInt(0 & Request("vut11")))
   'VUt12 = (1 = cInt(0 & Request("vut12")))
   'VUt121 = (1 = cInt(0 & Request("vut121")))
   'VUt13 = (1 = cInt(0 & Request("vut13")))
   'VUt2 = (1 = cInt(0 & Request("vut2")))
   'VUt21 = (1 = cInt(0 & Request("vut21")))
   'VUt22 = (1 = cInt(0 & Request("vut22")))
   'VUt221 = (1 = cInt(0 & Request("vut221")))
   'VUt23 = (1 = cInt(0 & Request("vut23")))
   'Fin Variables para configurar la vista de VisorAP
      
   
   
'	If Not Then
'		Response.Redirect "http:../Compartido/Error.asp?Titulo=VisorAP&Description=..."
'		Response.End 
'	End If

%>

<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<title>VisorAP</title>
<link rel="stylesheet" type="text/css" href="../Compartido/VisorAP.css">
</head>
<!--#INCLUDE virtual="/afexvisorap/Compartido/Boton.htm" -->
<!--#INCLUDE virtual="/afexvisorap/visorap/Rutinas.htm" -->
<script LANGUAGE="VBScript">

	'Inicio variables para identificar si vista está activa en pantalla
	Dim bVUt1, bVUt2, bVMn1, bCo1, bVPr3, bVCl1, bVCd1
	bVUt1 = False
	bVUt2 = False
	bVMn1 = False
	bVCo1 = False
	bVPr3 = False
	bVCd1 = False
	bVCl1 = False
	'Fin variables para identificar si vista está activa en pantalla
	
   Dim rs
	Dim nItvLVActual			'Contador Intervalo limpiar Visor
   Dim nUtActual				'Contador intervalo actualizacion utilidad
	nUtActual = <%=nItvUt%>

	On Error Resume Next
		
	Sub window_onload()
		nPromoActual = 1
		nVecesPromo = 2
		nVecesPromoActual = 1
		<% If cDate(Date()) = cDate(sFecha) Then %>
				window.setInterval "CargarVisor", <%=nItv%>, "vbscript"		
		<% End If %>
		CargarVisor
		Set rs = Nothing
		<% If VCo1 Then %>
			Cargar_configuracion
		<% End If %> 
	End Sub
	
	Sub CargarVisor
		On Error Resume Next
		
		<% If VPr1 Or VPr2 Then %>
				Set rs = ObtenerVisorAP("<%=sConexion%>", "<%=sMoneda%>", "<%=sFecha%>")				
		<% End If %>
		If Err.Number <> 0 Then 
			Err.Clear 
		End If
		CargarTBVisor
		Set rs = Nothing
		<% If VPromo Then %>
				Cargar_promociones
		<% End If %>		
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
		<% If VUt1 Or VUt2 Then %>
				nItvLVActual = nItvLVActual + 1
				If nItvLVActual >= <%=nItvLV%> Then
					If bVUt1 Then Mostrar_utilidad
					If bVUt2 Then Mostrar_utilidadmensual
					nItvLVActual = 0
				End If
				<% If VUt1 Then %>
						If bVUt1 Then Cargar_utilidad "<%=sConexion%>", "<%=sMonedaUti%>", "<%=sFecha%>", "<%=Session("NombreUsuario")%>"
				<% End If %>
				<% If VUt2 Then %>
						If bVUt2 Then Cargar_utilidadmensual "<%=sConexion%>", "<%=sMonedaUti%>", "<%=sFecha%>", "<%=Session("NombreUsuario")%>"
				<% End If %>
		<% End If %>
		<% If VMn1 Then %>
				If bVMn1 Then Cargar_monedas "<%=sConexion%>"
		<% End If %>
			
	End Sub	

</script>
<% If VLogo Then %>
<body border="0" style="margin: 0 0 0 0" background="../images/Afex%20VisorAP.gif" style="background-repeat: no-repeat">
<% Else %>
<body border="0" style="margin: 0 0 0 0">
<% End If %>
<form id="frmVisorAP" method="post">
<table id="tbVisor" border="0" cellpadding="0" cellspacing="0" style="border: 0px">

	<% If VLogo Then %>
			<tr height="39px" align="right"><td><img src="../images/Afex.jpg" WIDTH="30" HEIGHT="30">&nbsp;<td></td></tr>
	<% End If %>
	<tr align="left"><td class="tituloXP" nowrap style="color: '#444444'; background-color: '#efefef'; font-family: MS Sans Serif; height: 15; font-size: 8pt">&nbsp;<%=Session("AliasSC")%> - <%=Session("NombreEmpleado")%></td></tr>
	<% If VPromo Then %>
			<tr><td><!--#INCLUDE virtual="/afexvisorap/visorap/_encabezado.htm" --></td></tr>
	<% End If %>
	<% If VNtc Then %>
			<tr><td><!--#INCLUDE virtual="/afexvisorap/visorap/_noticias.htm" --></td></tr>
	<% End If %>
	<% If VPr1 Then %>
			<tr><td><!--#INCLUDE virtual="/afexvisorap/visorap/_tipocambio.htm" --></td><tr>
	<% End If %>
	<% If VPr2 Then %>
			<tr><td><!-- #INCLUDE virtual="/afexvisorap/visorap/_estadistica.htm" --><td><tr>
	<% End If %>		
	<% If VPr3 Then %>
			<tr><td><!-- #INCLUDE virtual="/afexvisorap/visorap/_posicion.htm" --><td><tr>
	<% End If %>	
	<tr><td><!-- #INCLUDE virtual="/afexvisorap/visorap/_opciones.htm" --></td></tr>
	<% If VPr3 Then %>
			<tr><td><!-- #INCLUDE virtual="/afexvisorap/visorap/_detalleposicion.htm" --><td><tr>
	<% End If %>	
	<% If VUt1 Then %>
			<tr><td><!-- #INCLUDE virtual="/afexvisorap/visorap/_utilidad.htm" --><td><tr>
	<% End If %>
	<% If VUt2 Then %>
			<tr><td><!-- #INCLUDE virtual="/afexvisorap/visorap/_utilidadmensual.htm" --><td><tr>
	<% End If %>
	<% If VMn1 Then %>
			<tr><td><!-- #INCLUDE virtual="/afexvisorap/visorap/_monedas.htm" --><td><tr>
	<% End If %>
	<% If VCo1 Then %>
			<tr><td><!-- #INCLUDE virtual="/afexvisorap/visorap/_configuracion.htm" --><td><tr>
	<% End If %>
	<% If VCl1 Then %>
			<tr><td><!-- #INCLUDE virtual="/afexvisorap/visorap/_clientes.htm" --><td><tr>
	<% End If %>
	<% If VCd1 Then %>
			<tr><td><!-- #INCLUDE virtual="/afexvisorap/visorap/_cierrediario.htm" --><td><tr>
	<% End If %>
</table>	
</form>
</body>
</html>