<%@  language="VBScript" %>
<%
	'Se asegura que la página no se almacene en la memoria cache
	Response.Expires = 0
	If Session("CodigoCliente") = "" Then
		response.Redirect "../Compartido/TimeOut.htm"
		response.end
	End If
    Dim sMensaje1, sMensaje2, sMostrarBotones
    sMensaje1 =  Request("mensajeCump")
    sMensaje2 =  Request("pregunta")
    sMostrarBotones =  Request("MostrarBotones")
    'Response.Write "MostrarBotones:" & sMostrarBotones

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
	window.dialogWidth = 25
    window.dialogHeight = 10
'	window.dialogLeft = 600
'	window.dialogTop =100
'	window.defaultstatus = ""
	
	Sub imgAceptar_onClick()
		window.returnvalue = "1" 	
        'window.returnvalue = frmGiroDeposito.cbxBanco.value & ";" & frmGiroDeposito.cbxTipoCta.value & ";" & _
		'					     frmGiroDeposito.txtNumeroCta.value & ";" & frmGiroDeposito.cbxMonedaDeposito.value 
		window.close
		
	End Sub		

    Sub imgCancelar_onClick()
        window.returnvalue = "0"
	    window.close
	End Sub	
//-->
</script>

<body sbgcolor="PowderBlue">
        <table id="tabConsulta" class="border" border="0" cellpadding="4" cellspacing="0"
            style="position: relative; left: 10px;">
            <tr>
                <td> <br /><br /><img id="img1" src="/images/Advetencia.jpg" width="40" height="40"><br /><br /><br /><br /></td>
                <td><!--<b><label style="color:red;">Origen Cumplimiento</label></b><br />-->
                        <label id="lblMensaje"><%=sMensaje1 %></label><br />
                        <label id="lblMensaje2"><%=sMensaje2 %></label>
                        </td>
            </tr>
            <tr align="middle" style="display: <%=sMostrarBotones%>">
                <td colspan="2">
                    <img id="imgAceptar" src="../images/BotonSi.jpg" style="cursor: hand" width="70"
                        height="20" >
                    <img id="imgCancelar" src="../images/BotonNo.jpg" style="cursor: hand" width="70"
                        height="20" />
                </td>
            </tr>
        </table>
    <!--#INCLUDE virtual="/Compartido/Rutinas.htm" -->
</body>
</html>
