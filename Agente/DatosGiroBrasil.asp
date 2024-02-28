<%@ Language=VBScript %>
<%
	'Se asegura que la página no se almacene en la memoria cache
	Response.Expires = 0
	If Session("CodigoCliente") = "" Then
		response.Redirect "../Compartido/TimeOut.htm"
		response.end
	End If
%>
<!-- #INCLUDE virtual="/Compartido/Errores.asp" -->
<!-- #INCLUDE virtual="/Compartido/Rutinas.asp" -->
<!-- #INCLUDE virtual="/agente/Constantes.asp" -->
<%

Dim cTC
Dim sSQL
Dim rsTC
Dim cReales

sSql = " SELECT  top 1 * " & _
       " FROM    Tipo_Cambio " & _
       " WHERE   fecha_termino IS Null " & _
       "  AND     sw_tipo = 2 " & _
       "  AND     codigo_moneda = " & EvaluarStr("BRR")
Set rsTC = EjecutarSQLCliente(Session("afxCnxAFEXpress"), sSQL)
If Err.Number <> 0 Then
	Set rsTC = Nothing
				
	Response.Redirect "../Compartido/Error.asp?description=" & err.Description
End If

If rsTC.EOF Then
	cTC = 0
Else
	cTC = rsTC("valor")
End If
cReales = ccur(Request("Dolares")) * ccur(cTC)

Set rsTC = Nothing
		
%>
<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<title>Datos Giro a Brasil</title>
<link rel="stylesheet" type="text/css" href="../Estilos/Cliente.css">
</head>

<script LANGUAGE="VBScript">
<!--
	window.dialogWidth = 26
	window.dialogHeight = 16
	window.dialogLeft = 240
	window.dialogTop = 220
	window.defaultstatus = ""
		
	Sub imgAceptar_onClick()
		
		window.returnvalue = window.txtBanco.value & ";" & window.txtAgencia.value & ";" & _
							 window.txtCtaCte.value & ";" & window.txtCPF.value & ";" & _
							 window.txtReales.value
		window.close		
	End Sub		

//-->
</script>
<body>
<center>
	<table border="0" cellpadding="1">
		<tr>
			<td>Banco&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Nº Agencia</td>
		</tr>
		<tr>
			<td>
				<input id="txtBanco" maxlength="30" style="width: 250px;" value="">			
				<input id="txtAgencia" maxlength="10" style="width: 100px;" value="">
			</td>
		</tr>		
		<tr>
			<td>Cta.Cte&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;CPF</td>
		</tr>
		<tr>
			<td>
				<input id="txtCtaCte" maxlength="17" style="width: 175px;" value="">			
				<input id="txtCPF" maxlength="17" style="width: 175px;" value="">
			</td>
		</tr>		
		<tr>
			<td colspan="2">Dolares&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
								&nbsp;&nbsp;&nbsp;
								Rate&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
								&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Reales</td>
		</tr>
		<tr>
			<td colspan="2">
				<input id="txtDolares" style="width: 120px;" disabled value="<%=Request("Dolares")%>">
				<input id="txtRate" style="width: 110px;" disabled value="<%=FormatNumber(cTC, 2)%>">
				<input id="txtReales" style="width: 120px;" disabled value="<%=FormatNumber(cReales, 2)%>">
			</td>
		</tr>
		<tr>
			<td>&nbsp;</td>
		</tr>
		<tr>
			<td align="right"><img border="0" id="imgAceptar" onclick src="../images/BotonAceptar.jpg" WIDTH="70" HEIGHT="20"></td>
		</tr>		
	</table>

</center>
</body>
</html>