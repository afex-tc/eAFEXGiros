<%@ Language=VBScript %>
<%


%>
<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<title>Periodo de Consulta</title>
<link rel="stylesheet" type="text/css" href="../Estilos/Cliente.css">
</head>

<script LANGUAGE="VBScript">
<!--
	On Error Resume Next
	Sub imgAceptar_onClick()
      window.event.returnValue = False
	   window.external.raiseEvent "Aceptar", "Aceptar"
	End Sub		

	Function public_put_Desde(sDesde)
		txtDesde.value =  sDesde
	End Function
	
	Function public_get_Desde()
		public_get_Desde = txtDesde.value 
	End Function

	Function public_put_Hasta(sHasta)
		txtHasta.value =  sHasta
	End Function
	
	Function public_get_Hasta()
		public_get_Hasta = txtHasta.value 
	End Function

	Sub txtDesde_onBlur()
		Dim sFecha
		sFecha = txtDesde.value
		If ValidarFechas(sFecha) Then
			txtDesde.value = sFecha
		Else
			txtDesde.focus 
			txtDesde.select 
		End If
	End Sub

	Sub txtHasta_onBlur()
		Dim sFecha

		sFecha = txtHasta.value
		If ValidarFechas(sFecha) Then
			txtHasta.value = sFecha
		Else
			txtHasta.focus 
			txtHasta.select 
		End If
	End Sub
	
	Function ValidarFecha(ByRef Fecha)
		Fecha = UCase(Trim(Fecha))
		Fecha = Replace(Fecha, "-", "")
		Fecha = Replace(Fecha, "/", "")
		ValidarFecha = True		

		Select Case Fecha
		Case "" 
			Fecha = Date()
			Exit Function
		Case "HOY"
			fecha = Date()
			Exit Function
		Case "AYER"
			Fecha = Date() - 1
			Exit Function
		End Select

		ValidarFecha = False
		If Len(Fecha) < 6 Then
			MsgBox "Debe ingresar una fecha válida con uno de estos formatos: " & vbCrLf & vbCrLf & _
					 "dd-mm-aaaa, dd-mm-aa, ddmmaa, HOY o AYER"
			Exit Function
		End If
		If Len(Fecha) = 7 Then
			MsgBox "Debe ingresar una fecha válida con uno de estos formatos: " & vbCrLf & vbCrLf & _
					 "dd-mm-aaaa, dd-mm-aa, ddmmaa, HOY o AYER"
			Exit Function
		End If
		If Len(Fecha) > 8 Then
			MsgBox "Debe ingresar una fecha válida con uno de estos formatos: " & vbCrLf & vbCrLf & _
					 "dd-mm-aaaa, dd-mm-aa, ddmmaa, HOY o AYER"
			Exit Function
		End If
		If Len(Fecha) = 6 Then
			Fecha = Left(Fecha, 2) & "-" & Mid(fecha, 3, 2) & "-"  & Right(Fecha, 2)
			Fecha = cDate(fecha)
		ElseIf Len(Fecha) = 8 Then
			Fecha = Left(Fecha, 2) & "-" & Mid(fecha, 3, 2) & "-" & Right(Fecha, 4)
			Fecha = cDate(fecha)
		End If
		ValidarFecha = True
	End Function
	
	Function ValidarFechas(ByRef Fecha)
		Fecha = UCase(Trim(Fecha))
		Fecha = Replace(Fecha, "-", "")
		Fecha = Replace(Fecha, "/", "")
		ValidarFechas = True		

		Select Case Fecha
		Case "" 
			Fecha = Date()
			Exit Function
		Case "HOY"
			fecha = Date()
			Exit Function
		Case "AYER"
			Fecha = Date() - 1
			Exit Function
		End Select

		ValidarFechas = False
		If Len(Fecha) < 6 Then
			MsgBox "Debes ingresar una fecha válida con uno de estos formatos: " & vbCrLf & vbCrLf & _
					 "dd-mm-aaaa, dd-mm-aa, ddmmaa, HOY o AYER"
			Exit Function
		End If
		If Len(Fecha) = 7 Then
			MsgBox "Debes ingresar una fecha válida con uno de estos formatos: " & vbCrLf & vbCrLf & _
					 "dd-mm-aaaa, dd-mm-aa, ddmmaa, HOY o AYER"
			Exit Function
		End If
		If Len(Fecha) > 8 Then
			MsgBox "Debes ingresar una fecha válida con uno de estos formatos: " & vbCrLf & vbCrLf & _
					 "dd-mm-aaaa, dd-mm-aa, ddmmaa, HOY o AYER"
			Exit Function
		End If
		If Len(Fecha) = 6 Then
			Fecha = Left(Fecha, 2) & "-" & Mid(fecha, 3, 2) & "-"  & Left(Trim(cStr(Year(Date))), 2) & Right(Fecha, 2)
		ElseIf Len(Fecha) = 8 Then
			Fecha = Left(Fecha, 2) & "-" & Mid(fecha, 3, 2) & "-" & Right(Fecha, 4)
		End If
		ValidarFechas = True
	End Function
	
//-->
</script>
<body>
<table id="tabPeriodo"  class="bordeinactivo" cellspacing="0" cellpadding="3">
<tr>
	<td colspan="3" class="tituloinactivo">Periodo&nbsp;&nbsp;<a style="font-size: 8pt">(ddmmyy)</a></td>
</tr>
<tr  bgcolor="white">
	<td>Desde&nbsp;<input SIZE="8" VALUE="<%=Date%>" id="txtDesde">&nbsp;&nbsp;&nbsp;</td>
	<td>Hasta&nbsp;<input SIZE="8" VALUE="<%=Date%>" id="txtHasta">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
	<td><img id="imgAceptar" src="../images/BotonAceptar.jpg" style="CURSOR: hand" WIDTH="70" HEIGHT="20"></td>
</tr>
</table>
<!--#INCLUDE virtual="/Compartido/Rutinas.htm" -->
</body>
</html>