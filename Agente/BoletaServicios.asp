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
	Dim sNombres, sApellidos, sDireccion
	Dim sAreaFono, sPaisFono, sFono, sDescripcion
	
	sNombres = Request("Nombres")
	sApellidos = Request("Apellidos")
	sDireccion = Request("Direccion")
	sFono = "(" & Request("PaisFono") & Request("AreaFono") & ") " & Request("Fono")
	sDescripcion = Request("Descripcion")
	
	
%>
<html>
<head>
    <meta name="GENERATOR" content="Microsoft Visual Studio 6.0" />
    <title>Boleta de Servicios</title>
    <link rel="stylesheet" type="text/css" href="../Estilos/Cliente.css" />
</head>

<script language="VBScript">
<!--
	window.dialogWidth = 38
	window.dialogHeight = 28
	window.dialogLeft = 160
	window.dialogTop = 220
	window.defaultstatus = ""
	
	Sub window_onLoad()
	End Sub
	
	
	Sub imgAceptar_onClick()
		
		window.close
		
	End Sub		

//-->
</script>

<body>
    <!--APPL-8183_MS_24-09-2014-->
    <center>
        <input type="hidden" id="txtCodigoCliente" />
        <table id="tabConsulta" class="border" border="0" cellpadding="4" cellspacing="0"
            style="height: 100%; font-size: 14pt" width="100%">
            <!--<tr><td class="Titulo" colspan="3" style="FONT-SIZE: 10pt; HEIGHT: 5px">Datos del Tercero</td></tr>-->
            <tr align="center">
                <td title="Giro" align="left">
                    <table>
                        <tr>
                            <td colspan="3" style="width:60%;" >
                                <b>Datos Remitente</b><br />
                                &nbsp;&nbsp;&nbsp;Nombre:&nbsp;<%=sNombres%>&nbsp;<%=sApellidos%><br />
                                &nbsp;&nbsp;&nbsp;Dirección:&nbsp;<%=sDireccion%><br />
                                &nbsp;&nbsp;&nbsp;Ciudad:&nbsp;<%=Request("Ciudad")%><br />
                                <br />
                            </td>
                            <td colspan="2">
                                <br />
                                &nbsp;&nbsp;&nbsp;Rut:&nbsp;<br />
                                &nbsp;&nbsp;&nbsp;Teléfono:&nbsp;<%=sFono%>
                            </td>
                        </tr>
                        <tr>
                            <td colspan="3">
                                <b>Giro</b><br />
                                &nbsp;&nbsp;&nbsp;Código:&nbsp;<%=Request("Codigo")%>
                            </td>
                            <td colspan="2">
                                <br />
                                <br />
                                &nbsp;&nbsp;&nbsp;PIN:&nbsp;<%=Request("sPIN")%><br />
                                <br />
                            </td>
                        </tr>
                        <tr>
                            <td colspan="3" style="width:60%;">
                                <b>Destino</b><br />
                                &nbsp;&nbsp;&nbsp;Ciudad:&nbsp;<%=Request("CiudadB")%>
                            </td>
                            <td colspan="2">
                                <br />
                                &nbsp;&nbsp;&nbsp;País:&nbsp;<%=Request("PaisB")%><br />
                                <br />
                            </td>
                        </tr>
                        <tr>
                            <td colspan="3" style="width:60%;">
                                <b>Destinatario</b><br />
                                &nbsp;&nbsp;&nbsp;Nombre:&nbsp;<%=Request("NombresB")%>&nbsp;<%=Request("ApellidosB")%><br />
                                &nbsp;&nbsp;&nbsp;Dirección:&nbsp;<%=Request("DireccionB")%><br />
                                &nbsp;&nbsp;&nbsp;Ciudad:&nbsp;<%=Request("CiudadB")%><br />
                                &nbsp;&nbsp;&nbsp;Mensaje:&nbsp;<%=Request("Mensaje")%><br />
                                <br />
                            </td>
                            <td colspan="2">
                                <br />
                                &nbsp;&nbsp;&nbsp;Teléfono:&nbsp;<%=Request("FonoB")%>
                            </td>
                        </tr>
                        <tr>
                            <td colspan="5">
                                <b>Transferencia</b><br />
                            </td>
                        </tr>
                        <tr>
                            <td>
                                Transferencia al Destinatario:
                            </td>
                            <td>
                            </td>
                            <td align="right">
                                <%=Request("Monto")%>
                            </td>
                            <td align="right">
                                <%=Request("MontoEquivalente")%>
                            </td>
                            <td>
                            </td>
                        </tr>
                        <tr>
                            <td>
                                &nbsp;&nbsp;&nbsp;Gastos de Transferencia:
                            </td>
                            <td align="right" style="width:15%;">
                                <%=Request("Gastos")%><br />
                            </td>
                            <td>
                            </td>
                            <td>
                            </td>
                            <td>
                            </td>
                        </tr>
                        <tr>
                            <td>
                                &nbsp;&nbsp;&nbsp;Comisión con IVA:
                            </td>
                            <td align="right">
                                <%=Request("Comision")%>
                            </td>
                            <td>
                            </td>
                            <td align="left">
                                &nbsp;&nbsp;&nbsp;TC&nbsp;<%=Request("TipoCambio")%>
                            </td>
                            <td align="left">
                                &nbsp;&nbsp;&nbsp;<%=Request("TotalNacional")%>
                            </td>
                        </tr>
                        <tr>
                            <td>
                                Tarifa:
                            </td>
                            <td>
                            </td>
                            <td align="right" style="width:15%;">
                                <%=Request("tarifa")%>
                            </td>
                            <td>
                            </td>
                            <td>
                            </td>
                        </tr>
                        <tr>
                            <td>
                                Total:
                            </td>
                            <td>
                            </td>
                            <td align="right">
                                <%=Request("totalCliente")%>
                            </td>
                            <td>
                            </td>
                            <td>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr align="middle">
                <td colspan="2">
                    <img id="imgAceptar" src="../images/BotonCerrar.jpg" style="cursor: hand" width="70"
                        height="20" />
                </td>
            </tr>
        </table>
    </center>
    <!--FIN APPL-8183_MS_24-09-2014-->
    <!--#INCLUDE virtual="/Compartido/Rutinas.htm" -->
</body>
</html>