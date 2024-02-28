<%@ Language=VBScript %>
<%
	Response.Expires = 0
	Response.buffer = true	
	If Not Session("SesionActiva") Then
'		response.Redirect "../Compartido/TimeOut.htm"
		response.end
	End If
%>
<html>
<head>
<style>
	TR.titulo
			{
				font-weight: bold; 
				height: 30px; 
				background-color: '#DFDFDF'
			}
			
	TR.fila
			{
				height: 20px; 
				background-color:	'#F6F6F6';
				cursor: hand
			}
</style>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<title>Utilidades <%=Request("nej")%></title>
<link rel="stylesheet" type="text/css" href="../Compartido/VisorAP.css">
</head>
<!--#INCLUDE virtual="/afexvisorap/visorap/Rutinas.htm" -->
<script LANGUAGE="VBScript">
<!--

	Dim rs

	Dim sPeriodoSQL
	Dim aPeriodo(2)
	Dim aTipo(3)
	Dim aGrupo(4)
	Dim sClipBoard

	Const afxPeriodoDia = 1
	Const afxPeriodoMensual = 2

	Const afxTipoPersonal = 1
	Const afxTipoEjecutivo = 2
	Const afxTipoAFEX = 3

	Const afxGrupoTipoCliente = 1
	Const afxGrupoEjecutivo = 2
	Const afxGrupoMoneda = 3
	Const afxGrupoFecha  = 4


	
	Sub ListaTipo()
	   aPeriodo(afxPeriodoDia) = ""            'Dia
	   aPeriodo(afxPeriodoMensual) = "Mensual"     'Mensual
	   
	   aTipo(afxTipoPersonal) = ""               'Personal
	   aTipo(afxTipoEjecutivo) = ""               'Ejecutivo
	   aTipo(afxTipoAFEX) = "AFEX"
	   
	   aGrupo(afxGrupoTipoCliente) = "TipoCliente"
	   aGrupo(afxGrupoEjecutivo) = "Ejecutivo"
	   aGrupo(afxGrupoMoneda) = "Moneda"
	   aGrupo(afxGrupoFecha) = "Fecha"
	End Sub

	Sub CargarForm()
	   Dim sPeriodo
	   Dim sMoneda
	   Dim sEjecutivo
	      
	   Select Case <%=Request("pr")%>
	   Case 2
	      sPeriodo = NombreMes(Month(<%=Request("fch")%>))
	      sPeriodoSQL = "Mensual"
	   Case Else
	      sPeriodo = <%=Request("fch")%>
	      sPeriodoSQL = ""
	   End Select
	   If cInt(0 & "<%=Request("gr")%>") = afxGrupoMoneda Then sMoneda = "<%=Request("mn")%>"
	   If cInt(0 & "<%=Request("tp")%>") = afxTipoPersonal Then sEjecutivo = "<%=Request("nej")%>"
	   
	   Caption = Trim(Trim(Caption & " " & sMoneda) & " " & sEjecutivo) & " - " & sPeriodo
	End Sub

	Sub ObtenerListaSP()
	   Dim sSQL
	   Dim sEjecutivo
	   
	   If <%=Request("tp")%> = afxTipoPersonal Then sEjecutivo = "<%=Request("ej")%>"
	   On Error Resume Next
	   sSQL = "APObtenerUtilidad" & aTipo(<%=Request("tp")%>) & aGrupo(<%=Request("gr")%>) & aPeriodo(<%=Request("pr")%>) & "<%=Request("cmp")%> '" & FormatoFechaSQL("<%=Request("fch")%>") & "', " & EvaluarStr("<%=Request("mn")%>") & ", " & EvaluarStr("<%=Request("ej")%>")
	   Set rs = EjecutarSQLCliente("<%=Session("cnxVisorAP")%>", sSQL)	   
	   If Err.Number <> 0 Then 
			Set rs = Nothing
		   MsgBox Err.Description
		End If
	End Sub

	Sub window_onload()
		ListaTipo
		CargarForm
		ObtenerListaSP
		CargarLista
		Set rs = Nothing
		'window.moveTo window.screenLeft, window.screenTop + clng(0 & "<%=Request("top")%>")
		'window.screenTop=window.screenTop + cInt(0 & "<%=Request("top")%>")
	End Sub
	
	Sub CargarLista
		Dim sHTML, sIngreso, sPorcentaje
		
		sHTML = "<table id=tbLista border=0 cellspacing=1 cellpadding=1 style=""border: 1px solid '#EEEEEE'"">"
		sHTML = sHTML & "<tr align=center class=titulo style="""">" & _
							 "<td width=200px>" & rs.Fields(0).name & "</td>" & _
							 "<td width=80px>" & rs.Fields(1).name & "</td>" & _
							 "<td width=0px style=""display: none"">" & rs.Fields(2).name & "</td>" & _
							 "<td width=100px>" & rs.Fields(3).name & "</td>" & _
							 "</tr>"
		sClipBoard = ""
		Do Until rs.EOF
			If IsNull(rs.Fields(1)) Then
				sIngreso = 0
			Else
				sIngreso = rs.Fields(1)
			End If
			If IsNull(rs.Fields(3)) Then
				sPorcentaje = 0
			Else
				sPorcentaje = rs.Fields(3)
			End If
			sHTML = sHTML & _
								 "<tr style=""height: 22px; cursor: hand; "" color=#BBBBBB bgcolor=#F6F6F6 onMouseOver=""javascript:this.bgColor='#EEEEEE'"" onMouseOut=""javascript:this.bgColor='#F6F6F6'"" onClick=""MostrarUtilidad '" & rs.Fields(0) & "', '" & rs.Fields(2) & "' "">" & _
								 "<td align=left >" & rs.Fields(0) & "</td>" & _
								 "<td align=right>" & FormatNumber(sIngreso, 0) & "</td>" & _
								 "<td style=""display: none"">" & rs.Fields(2) & "</td>" & _
								 "<td align=right>" & FormatNumber(sPorcentaje, 2) & "</td></tr>"
			sClipBoard = sClipBoard & rs.Fields(0) & vbTab & sIngreso & vbCrLf
			rs.MoveNext
		Loop
		sHTML = sHTML & "</table>"
		dvLista.innerHTML = sHTML

	End Sub

	Sub MostrarUtilidad(ByVal Nombre, Byval Codigo)
		Dim sMoneda 
      sMoneda = "<%=Request("mn")%>"
      Select Case <%=Request("gr")%>
      Case 1
				Window.showModelessDialog  "Operaciones.asp?fch=<%=Request("fch")%>&hst=<%=Request("hst")%>&mn=<%=Request("mn")%>&pr=<%=Request("pr")%>&tp=" & Codigo & "&gr=1&ej=<%=Request("ej")%>&nej=<%=Request("nej")%>", , "dialogWidth:50; dialogHeight:20"
				'Window.open  "Operaciones.asp?fch=<%=Request("fch")%>&hst=<%=Request("hst")%>&mn=<%=Request("mn")%>&pr=<%=Request("pr")%>&tp=" & Codigo & "&gr=1&ej=<%=Request("ej")%>&nej=<%=Request("nej")%>"
				
      Case 2
				Window.showModelessDialog  "Utilidad.asp?fch=<%=Request("fch")%>&mn=<%=Request("mn")%>&pr=<%=Request("pr")%>&tp=<%=Request("tp")%>&gr=1&ej=" & Codigo & "&nej=" & Nombre, , "dialogWidth:25; dialogHeight:20"
		End Select
		
	End Sub

	Sub Copiar_onclick
		PortaPapeles.SetText sClipBoard
	End Sub
	
//-->
</script>
<body>
<div id="dvLista">
</div>
<OBJECT id="PortaPapeles" style="LEFT: 0px; TOP: 0px"   codebase="../Compartido/PortaPapeles.CAB#version=1,0,0,0" 
	classid=CLSID:BB88D9B4-BD22-41BE-B64A-C615EA1D19B6 VIEWASTEXT>
	<PARAM NAME="_ExtentX" VALUE="1">
	<PARAM NAME="_ExtentY" VALUE="1">
</OBJECT>
<br><br>
<center><img id="Copiar" style="cursor: hand" src="../images/botoncopiar.jpg" WIDTH="70" HEIGHT="20"></center>
<br>
</body>
<script>
	Set rs = Nothing
</script>
</html>
