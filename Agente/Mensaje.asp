<%@  language="VBScript" %>
<%
	'Se asegura que la página no se almacene en la memoria cache
	Response.Expires = 0
	If Session("CodigoCliente") = "" Then
		response.Redirect "../Compartido/TimeOut.htm"
		response.end
	End If
    Dim sRegla
    sRegla =  Request("mensaje")
    
%>
<!-- #INCLUDE virtual="/Compartido/Errores.asp" -->
<!-- #INCLUDE virtual="/Compartido/Rutinas.asp" -->
<!-- #INCLUDE virtual="/agente/Constantes.asp" -->
<%

%>
<html>
<head>
    <meta name="GENERATOR" content="Microsoft Visual Studio 6.0">
    <title>Cumplimiento</title>
    <link rel="stylesheet" type="text/css" href="../Estilos/Cliente.css">
</head>

<script language="VBScript">
<!--
'	window.dialogWidth = 28
    window.dialogHeight = 13
'	window.dialogLeft = 600
'	window.dialogTop =100
'	window.defaultstatus = ""
	
	Sub imgAceptar_onClick()
		'window.returnvalue = Trim(observacion.value) 	
		window.close
		
	End Sub		

//-->
</script>

<body sbgcolor="PowderBlue">
        <table id="tabConsulta" class="border" border="0" cellpadding="4" cellspacing="0"
            style="position: relative; left: 10px;">
            <tr>
                <td> <img id="img1" src="/images/Advetencia.jpg" width="40" height="40"><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /></td>
                <td><b><label style="color:red;">Advertencia</label></b><br />
                    <b>Esta operación debe ser sometida a reglas de cumplimiento</b><br /><br />
                       <b>Detalle:</b><br />
                        <label id="lblMensajeRegla"><%=sRegla %></label><br />
                       Debe solicitar al cliente:<br />
                       - Ficha de Cliente<br />
                       - Copia de C.I.<br />
                       - Ficha de Control de Operaciones<br />
                       Solicitar AUTORIZACIÓN  al Departamento de Atención de Clientes.<br /><br />
                       <b>Origen:</b><br />
                       <b><label style="color:red;">Cumplimiento</label></b>
                </td>
            </tr>
            <tr align="middle">
                <td colspan="2">
                    <img id="imgAceptar" src="../images/BotonAceptar.jpg" style="cursor: hand" width="70"
                        height="20">
                </td>
            </tr>
        </table>
    <!--#INCLUDE virtual="/Compartido/Rutinas.htm" -->
</body>
</html>
