<%@ Language=VBScript %>
<%
	Response.Expires = 0
	Response.buffer = true	
	If Not Session("SesionActiva") Then
'		response.Redirect "../Compartido/TimeOut.htm"
		response.end
	End If

%>
<%
	Dim nPaginaActual
	
	nPaginaActual = cInt(0 & Request("pag"))
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
<link rel="stylesheet" type="text/css" href="../Estilos/Reportes.css">
</head>
<!--#INCLUDE virtual="/afexvisorap/visorap/Rutinas.htm" -->
<script LANGUAGE="VBScript">
<!--

	Dim rs
	Dim sClipBoard
	Dim nPaginaActual
	Dim nPageSize
	Dim nPageCount
	
	Sub ObtenerListaSP()
		Dim sSQL
		Dim sClip
		Dim i
		Dim sPeriodo
   
		On Error Resume Next
		Select Case <%=Request("pr")%>
		Case 2
		   sPeriodo = "Mensual"
		Case Else
		   sPeriodo = ""
		End Select
   
		If "<%=Request("rt")%>" = "00000000" Then sRut = ""
		Select Case <%=Request("tp")%>
		Case 0
		   sSQL = "APObtenerSPEjecutivo " & EvaluarStr("<%=Request("mn")%>") & ", " & EvaluarStr(FormatoFechaSQL("<%=Request("fch")%>")) & ", " & EvaluarStr(FormatoFechaSQL("<%=Request("hst")%>")) & ", " & EvaluarVar("<%=Request("ej")%>", "Null") & ", " & EvaluarStr(sRut) & ", " & EvaluarVar("<%=Request("to")%>", "Null")
		Case 1
		   sSQL = "APObtenerSPTipoCliente" & sPeriodo & " " & EvaluarStr("<%=Request("mn")%>") & ", " & EvaluarStr(FormatoFechaSQL("<%=Request("fch")%>")) & ", " & EvaluarVar("<%=Request("ej")%>", "Null")
		Case 2
		   sSQL = "APObtenerSPTipoSucursal" & sPeriodo & " " & EvaluarStr("<%=Request("mn")%>") & ", " & EvaluarStr(FormatoFechaSQL("<%=Request("fch")%>")) & ", " & EvaluarVar("<%=Request("ej")%>", "Null")
		Case 3
		   sSQL = "APObtenerSPTipoAgente" & sPeriodo & " " & EvaluarStr("<%=Request("mn")%>") & ", " & EvaluarStr(FormatoFechaSQL("<%=Request("fch")%>")) & ", " & EvaluarVar("<%=Request("ej")%>", "Null")
		Case 4
		   sSQL = "APObtenerSPEjecutivo " & EvaluarStr("<%=Request("mn")%>") & ", " & EvaluarStr(FormatoFechaSQL("<%=Request("fch")%>")) & ", " & EvaluarStr(FormatoFechaSQL("<%=Request("hst")%>")) & ", " & EvaluarVar("<%=Request("ej")%>", "Null") & ", " & EvaluarStr(sRut)
		End Select
	   Set rs = EjecutarSQLCliente("<%=Session("cnxVisorAP")%>", sSQL)
	   If Err.Number <> 0 Then 
			Set rs = Nothing
		   MsgBox "..." & Err.Description
		End If
	End Sub

	Sub window_onload()
		ObtenerListaSP

		rs.Pagesize = 50
		nPageSize = 50
		nPageCount = rs.PageCount

		CargarLista 1
		nPaginaActual = 1
		'Set rs = Nothing
	End Sub
	
	Sub window_unload()
		Set rs = Nothing
	End Sub
	
	Sub CargarLista(ByVal PaginaActual)
		Dim sHTML, sIngreso, sMonto
		Dim i
		If PaginaActual < 1 Then PaginaActual = 1		
		If PaginaActual > nPageCount Then PaginaActual = nPageCount
		
		On Error Resume Next
		sHTML = "<b style=""cursor: hand "" onClick=""CargarLista(" & PaginaActual - 1 & ") "">«</b>&nbsp;"
		i = 0
		For i = 1 To nPageCount
			sHTML = sHTML & "<b style=""cursor: hand "" onClick=""CargarLista(" & i & ") "">" & i & "</b>&nbsp;"
		Next
		sHTML = sHTML & "<b style=""cursor: hand "" onClick=""CargarLista(" & PaginaActual + 1 & ") "">»</b>&nbsp;"
		i = 0
		sHTML = sHTML & "<br>"
		sHTML = sHTML & "<table id=tbLista border=0 cellspacing=1 cellpadding=1 style=""border: 1px solid '#EEEEEE'"">"
		sHTML = sHTML & "<tr align=center class=titulo >" & _
							 "<td width=100px>" & rs.Fields(0).name & "</td>" & _
							 "<td width=300px>" & rs.Fields(1).name & "</td>" & _
							 "<td width=100px>" & rs.Fields(2).name & "</td>" & _
							 "<td width=80px>" & rs.Fields(3).name & "</td>" & _
							 "<td width=60px>" & rs.Fields(4).name & "</td>" & _
							 "<td width=60px>" & rs.Fields(5).name & "</td>" & _
							 "<td width=60px>" & rs.Fields(6).name & "</td>" & _
							 "<td width=60px>" & rs.Fields(7).name & "</td>" & _
							 "<td width=20px>" & rs.Fields(8).name & "</td>" & _
							 "<td width=20px style=""display: none"">" & rs.Fields(9).name & "</td>" & _
							 "<td width=20px style=""display: none"">" & rs.Fields(10).name & "</td>" & _
							 "<td width=20px style=""display: none"">" & rs.Fields(11).name & "</td>" & _
							 "</tr>"
		Err.Clear 
		sClipBoard = ""
		sClipBoard = rs.Fields(0).name & vbTab & rs.Fields(1).name & vbTab & rs.Fields(2).name & _
						 vbTab & rs.Fields(3).name & vbTab & rs.Fields(4).name & vbTab & rs.Fields(5).name & _
						 vbTab & rs.Fields(6).name & vbTab & rs.Fields(7).name & vbTab & rs.Fields(8).name & _
						 vbTab & rs.Fields(9).name & vbTab & rs.Fields(10).name & vbTab & rs.Fields(11).name & vbCrLf 
		'rs.PageCount = cInt(0 & "<%=nPaginaActual%>")
		rs.AbsolutePage = PaginaActual
		'Do Until rs.EOF
		For i = 1 To nPagesize
			If rs.eof then Exit For
			
			If IsNull(rs.Fields(7)) Then
				sIngreso = 0
			Else
				sIngreso = FormatNumber(rs.Fields(7), 0)
			End If
			If IsNull(rs.Fields(3)) Then
				sMonto = 0
			Else
				sMonto = FormatNumber(rs.Fields(3), 2)
			End If
			sHTML = sHTML & _
								 "<tr style=""height: 22px; cursor: hand; "" color=#BBBBBB bgcolor=#F6F6F6 onMouseOver=""javascript:this.bgColor='#EEEEEE'"" onMouseOut=""javascript:this.bgColor='#F6F6F6'"" onClick=""MostrarUtilidad '" & rs.Fields(0) & "', '" & rs.Fields(2) & "' "" title="" " & "SP:" & rs.Fields(11) & ", Fecha:" & rs.Fields(09) & ", Hora:" & rs.Fields(10) & " "">" & _
								 "<td align=left >" & rs.Fields(0) & "</td>" & _
								 "<td align=left>" & rs.Fields(1) & "</td>" & _
								 "<td align=center>" & rs.Fields(2) & "</td>" & _
								 "<td align=right>" & sMonto & "</td>" & _
								 "<td align=right>" & FormatNumber(rs.Fields(4), 2) & "</td>" & _
								 "<td align=right>" & FormatNumber(rs.Fields(5), 2) & "</td>" & _
								 "<td align=right>" & FormatNumber(rs.Fields(6), 2) & "</td>" & _
								 "<td align=right>" & sIngreso & "</td>" & _
								 "<td align=right>" & FormatNumber(rs.Fields(8), 2) & "</td>" & _
								 "<td align=left style=""display: none"">" & rs.Fields(9) & "</td>" & _
								 "<td align=left style=""display: none"">" & rs.Fields(10) & "</td>" & _
								 "<td align=left style=""display: none"">" & rs.Fields(11) & "</td>" & _
								 "</tr>"
			err.Clear 
			sClipBoard = sClipBoard & _
							 rs.Fields(0) & vbTab & rs.Fields(1) & vbTab & rs.Fields(2) & _
							 vbTab & sMonto & vbTab & rs.Fields(4) & vbTab & rs.Fields(5) & _
							 vbTab & rs.Fields(6) & vbTab & sIngreso & vbTab & rs.Fields(8) & _
							 vbTab & rs.Fields(9) & vbTab & rs.Fields(10) & vbTab & rs.Fields(11) & vbCrLf 
			rs.MoveNext			
		Next
		'Loop
		sHTML = sHTML & "</table>"
		dvLista.innerHTML = sHTML

	End Sub


	Sub MostrarUtilidad(ByVal Nombre, Byval Codigo)
		Dim s<%=Request("mn")%> 
      s<%=Request("mn")%> = "<%=Request("mn")%>"
      Select Case <%=Request("gr")%>
      Case 1
      
      Case 2
				Window.showModelessDialog  "Utilidad.asp?fch=<%=Request("fch")%>&mn=<%=Request("mn")%>&pr=1&tp=1&gr=1&ej=" & Codigo & "&nej=" & Nombre, , "dialogWidth:25; dialogHeight:20"
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
<OBJECT id="PortaPapeles" style="LEFT: 0px; TOP: 0px"    codebase="../../../../Compartido/PortaPapeles.CAB#version=1,0,0,0" 
	classid=CLSID:BB88D9B4-BD22-41BE-B64A-C615EA1D19B6 VIEWASTEXT>
	<PARAM NAME="_ExtentX" VALUE="1">
	<PARAM NAME="_ExtentY" VALUE="1">
</OBJECT>
<br><br>
<center><img id="Copiar" style="cursor: hand" src="../images/botoncopiar.jpg" WIDTH="70" HEIGHT="20"></center>
<br>
</body>
<script>
	'Set rs = Nothing
</script>
</html>
