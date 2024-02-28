<%@  language="VBScript" %>
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

%>
<html>
<head>
    <meta name="GENERATOR" content="Microsoft Visual Studio 6.0">
    <title>Rechazo/Deshabilitación de documentos</title>
    <link rel="stylesheet" type="text/css" href="../Estilos/Cliente.css">
</head>

<script language="VBScript">
<!--
	'window.dialogWidth = 35
	window.dialogHeight = 10
	window.dialogLeft = 100
	window.dialogTop = 220
	window.defaultstatus = ""
	
	Sub imgAceptar_onClick()
		window.returnvalue = Trim(observacion.value) 	
		window.close
		
	End Sub		

//-->
</script>

<body sbgcolor="PowderBlue">
    <center>
        <input type="hidden" id="txtCodigoCliente">
        <table id="tabConsulta" class="border" border="0" cellpadding="4" cellspacing="0"
            style="position: relative; left: 4px; height: 100px" width="200px">
            <tr>
                <td colspan="2">
                 <center>Motivo de rechazo/deshabilitación</center>
                </td>
            </tr>
            <tr>
                <td>
                                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
                <td>
                    <input type="text" name="Observacion" size="60" maxlength="40" onkeypress="IngresarTexto(3)">
                </td>
            </tr>
            <tr align="middle">
                <td colspan="2">
                    <img id="imgAceptar" src="../images/BotonAceptar.jpg" style="cursor: hand" width="70"
                        height="20">
                </td>
            </tr>
        </table>
    </center>
    <!--#INCLUDE virtual="/Compartido/Rutinas.htm" -->
</body>
</html>
