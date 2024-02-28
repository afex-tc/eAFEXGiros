<%@ Language=VBScript %>
<%
	Dim rsTipoDocumento, Sql
	Dim nCodigoCliente, sNombreCliente, sRutCliente
	Dim sTitulo
	
	nCodigoCliente = Request("cc")
	sNombreCliente = Request("nc")
	sRutCliente = Request("rt")
	
	sTitulo = sNombreCliente
	If Trim(sTitulo) = Empty Then
		sTitulo = "Agrega un Nuevo Documento"
	End If

	Sql = "Select codigo, nombre From Tipo_Documento Where codigo <> 3 and codigo not in (select codigo from Tipo_Documento where codigo between 22 and 26)" 
 'con esto se omiten los tipos de documentos que ya no se usan, pero no se eliminaran de la tabla para mantener historial
	
	Set rsTipoDocumento = CreateObject("ADODB.Recordset")
	rsTipoDocumento.Open Sql, Session("afxCnxCorporativa")
	
%>

<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<title>Agrega un Nuevo Documento</title>
<link rel="stylesheet" type="text/css" href="../Estilos/Cliente.css">
</head>

<script LANGUAGE="VBScript">
<!--
	Const sEncabezadoFondo = "Servicios"
	Const sEncabezadoTitulo = "Nuevo Documento"
	Const sClass = "TituloPrincipal"

	Sub imgAceptar_onClick()
		frmDocumento.action = "GuardarNuevoDocumento.asp?cc=<%=nCodigoCliente%>&nc=<%=sNombreCliente%>&rt=<%=sRutCliente%>"
		frmDocumento.submit
		frmDocumento.action = ""
    
	End Sub		
//-->
</script>
<!--#INCLUDE virtual="/Compartido/Encabezado.htm" -->
<body>
<form id="frmDocumento" method="post">
<center>
<table align="center" id="tabNuevoDoc" class="borde" border="0" cellpadding="4" cellspacing="0" style="width: 400px">	
<tr><td class="Titulo" colspan="2" style="font-size: 10pt; height: 5px">
      <p><%=sTitulo%></p>   </td></tr>
<tr>
	<td>
		Documento
		<select id="cbxTipoDocumento" name="cbxTipoDocumento" style="height: 22px; width: 280px"> 
			<%Do Until rsTipoDocumento.EOF%>
				  <option value="<%=rsTipoDocumento("codigo")%>"><%=rsTipoDocumento("nombre")%></option>
				  <%rsTipoDocumento.Movenext%>
			<%Loop%> 
			<%Set rsTipoDocumento = Nothing%>
		</select><br><br>
		Número&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		<input style="height: 22px; text-align: right; width: 55px" onKeyPress="IngresarTexto(1)" value="0" id="txtNumero" name="txtNumero" maxlength="5">
		Nombre archivo&nbsp;&nbsp;&nbsp;
		<input style="height: 22px; text-align:left; width: 155px" id="txtNombre" name="txtNombre" >
	</td>
</tr>
<tr>
	<td>
	</td>
</tr>
<tr align="middle">
	<td colspan="2"><img id="imgAceptar" src="../images/BotonAceptar.jpg" style="cursor: hand" width="70" height="20"></td>
</tr></tbody></table>
</center><!--#INCLUDE virtual="/Compartido/Rutinas.htm" -->
</form>
</body>
</html>