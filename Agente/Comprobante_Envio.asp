<%@  codepage="65001" %>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Comprobante de Envío</title>
    <style type="text/css">
        .style2 {
            width: 40%;
            height: 16px;
        }
    </style>
</head>
<body onload="Imprimir()">
    <form id="form1" runat="server">
        <div>
            <table style="border-right: 1px solid; border-top: 1px solid; border-left: 1px solid; width: 100%; border-bottom: 1px; text-align: left">
                <tr>
                    <td style="width: 70%; text-align: center">
                        <span style="font-size: 16px; padding-bottom: 9px; margin: 0px 0px 20px; border-bottom: #f2f2f2 1px solid; font-family: Verdana; letter-spacing: -0.03em">Comprobante de Envío
                        </span>
                    </td>
                </tr>
            </table>
            <table style="border-right: 1px solid; border-top: 1px; font-size: 13px; border-left: 1px solid; width: 100%; border-bottom: 1px; font-family: Verdana; text-align: left">
                <tr>
                    <td style="text-align: left">
                        <b>Número de referencia: <% =session("NumeroReferencia") %></b>
                    </td>
                    <td style="text-align: left">
                        <b>Código AFEX: <% =Session("CodigoAfex")%></b>
                    </td>
                    <td style="text-align: right">
                        <% =session("FechaC") %>
                        &nbsp;           
                        <% Session("HoraC") = right("00" & Session("HoraC"),6) %>
                        <%=left(Session("HoraC"), 2) & ":" & Mid(Session("HoraC"), 3, 2) & ":" & Right(Session("HoraC"), 2)%>           
                    </td>
                </tr>
            </table>
            <div style="border-right: 1px; border-top: 1px solid; border-left: 1px; width: 100%; border-bottom: 1px; font-family: Verdana; text-align: left">
                <table style="font-size: 10px; width: 100%">
                    <tr>
                        <td style="width: 60%; text-align: left">
                            <span style="font-weight: normal; font-size: 12px; padding-bottom: 9px; margin: 0px 0px 20px; border-bottom: #f2f2f2 1px solid; font-family: Verdana, Tahoma, arial; letter-spacing: -0.03em">Remitente / Persona que envía
                            </span>
                        </td>
                        <td style="width: 40%"></td>
                        <tr>
                            <td style="width: 60%; text-align: left">Nombre:<b><% =Session("NombreRemitente")%></b>
                            </td>
                            <td style="width: 40%">Nº identificación: <b><% =Session("NumeroIdentificacionRemitente")%></b>
                            </td>
                        </tr>
                </table>
                <table style="font-size: 10px; width: 100%">
                    <tr>
                        <td style="width: 60%; text-align: left">Dirección: <b><% =Session("DireccionRemitente")%></b>
                        </td>
                        <td style="width: 40%">Tipo de identificación: <b><% =Session("TipoIdentificacion")%></b>
                        </td>
                    </tr>
                </table>
                <table style="font-size: 10px; width: 100%">
                    <tr>
                        <td style="width: 60%; text-align: left">Ciudad: <b><% =Session("CiudadRemitente")%></b>
                        </td>
                        <td style="width: 40%">Teléfono: <b><% =Session("TelefonoRemitente")%></b>
                        </td>
                    </tr>
                </table>
                <table style="font-size: 10px; width: 100%">
                    <tr>
                        <td style="width: 60%; text-align: left">Fecha de nacimiento: <b><% =Session("FechaNacimientoRemitente")%></b>
                        </td>
                        <td style="width: 40%">Nacionalidad: <b><% =Session("NacionalidadRemitente")%></b>
                        </td>
                    </tr>
                    <tr>
                        <td style="text-align: left" colspan="2">Ocupación: <b><% =Session("OcupacionRemitente")%></b>
                        </td>
                    </tr>
                </table>
            </div>
            <div style="border-right: 1px; border-top: 1px solid; border-left: 1px; width: 100%; border-bottom: 1px solid; font-family: Verdana; text-align: left">
                <table style="font-size: 10px; width: 100%; height: 71px">
                    <tr>
                        <td colspan="2" style="text-align: left">
                            <span style="font-weight: normal; font-size: 12px; padding-bottom: 9px; margin: 0px 0px 20px; border-bottom: #f2f2f2 1px solid; font-family: Verdana, Tahoma, arial; letter-spacing: -0.03em">Beneficiario / Persona que recibe:
                            </span>
                        </td>
                    </tr>
                    <tr>
                        <td colspan="2" style="text-align: left">Nombre: <b><% =Session("NombreBeneficiario")%></b>
                        </td>
                    </tr>
                    <tr>
                        <td style="width: 60%; text-align: left">Ciudad: <b><% =Session("CiudadBeneficiario")%></b>
                        </td>
                        <td style="width: 40%; text-align: left">País: <b><% =Session("PaisBeneficiario")%></b>
                        </td>
                    </tr>
                    <tr>
                        <td colspan="2" style="text-align: left">Mensaje :<b><% =Session("Mensaje")%></b>
                        </td>
                    </tr>

                    <tr>
                        <td colspan="2" style="text-align: left">&nbsp;
                        </td>
                    </tr>
                </table>
            </div>
            <table style="border-right: 1px solid; border-top: 1px; font-size: 10px; border-left: 1px solid; width: 100%; border-bottom: 1px solid; font-family: Verdana; text-align: left">
                <tr>
                    <td style="width: 60%; vertical-align: top;">
                        <%= Replace(Replace(Replace(Replace(Session("MensajeFirma"),"Ã¡","&aacute;"),"Ã­","&iacute;"),"Ã³","&oacute;"),"Ãº","&uacute;") %>
                    </td>
                    <td style="width: 40%;">
                        <table style="border-right: 1px solid; border-top: 1px; font-size: 10px; border-left: 1px solid; width: 100%; border-bottom: 1px solid; font-family: Verdana; text-align: left">
                            <tr>
                                <td style="text-align: left">
                                    <span style="font-weight: normal; font-size: 12px; padding-bottom: 9px; margin: 0px 0px 20px; border-bottom: #f2f2f2 1px solid; font-family: Verdana, Tahoma, arial; letter-spacing: -0.03em">Detalle Giro
                                    </span>
                                </td>
                            </tr>
                            <tr>
                                <td style="text-align: left">Agencia: <b><% =Session("Agencia")%></b>
                                </td>
                            </tr>
                            <tr>
                                <td style="text-align: left">
                                    <% If session("Monto")= "USD" then %>
				                        Monto: <% =Session("Monto")%> <% =FormatNumber(Session("MonedaMontoEnvio"), 2, -1, 0, -1)%>
                                    <% else %>
				                        Monto: <% =Session("Monto")%> <% =FormatNumber(Session("MonedaMontoEnvio"), 0, -1, 0, -1)%>
                                    <% End If %>
                                </td>
                            </tr>
                            <tr>
                                <td style="text-align: left" class="style2">
                                    <% If session("Monto")= "USD" then %>
	                                    Cargo:  <% =Session("Cargo")%><% =FormatNumber( Session("MonedaEnvioCargo"),  2, -1, 0, -1)%>
                                    <% Else %>
				                        Cargo:  <% =Session("Cargo")%><% =FormatNumber( Session("MonedaEnvioCargo"),  0, -1, 0, -1)%>
                                    <% End If %>
                                </td>
                            </tr>
                            <tr>
                                <td style="text-align: left">
                                    <% If  session("TotalGiro") ="US$" then %>       
                                        Total: <b style="font-size: 14px"><% =Session("TotalGiro")%> <% =FormatNumber(Session("totalRecibir"),2 , -1, 0, -1)%></b>
                                    <%else %>
			                            Total: <b style="font-size: 14px"><% =Session("TotalGiro")%> <% =FormatNumber(Session("totalRecibir"), 0, -1, 0, -1)%></b>
                                    <% End If %>   
                                </td>
                            </tr>
                            <tr>
                                <td style="text-align: left">
                                    <% IF session("MonedaRecibir")= "USD" then %>
				                        Monto a recibir: <% =session("MonedaRecibir") %>&nbsp; <% =FormatNumber(Session("MontoRecibir"), 2, -1, 0, -1)%>
                                    <% else %>
				                        Monto a recibir: <% =session("MonedaRecibir") %>&nbsp; <% =FormatNumber(Session("MontoRecibir"), 0, -1, 0, -1)%>
                                    <% End If %>
                                </td>
                            </tr>
                            <tr>
                                <td style="text-align: left">Atendido por: <% =Session("AtendidoPor")%>
                                </td>
                            </tr>
                        </table>
                    </td>
                </tr>
                <tr>
                    <td style="width: 60%">
                        <table style="width: 260px">
                            <tr>
                                <td style="border-top-width: 1px; border-left-width: 1px; border-left-color: black; width: 100%; border-top-color: black; border-bottom: black 1px solid; border-right-width: 1px; border-right-color: black"></td>
                                <tr>
                                    <td style="text-align: center">Firma Cliente</td>
                                </tr>
                        </table>
                    </td>
                    <td style="width: 40%; text-align: left">Estado Giro: <% =Session("EstadoGiro")%>
                    </td>
                </tr>
            </table>
            <table style="font-size: 10px; left: 10px; width: 100%; TOP: 500px; height: 38px">
                <tr>
                    <td align="left" style="border-top-width: thin; border-left-width: thin; vertical-align: top; border-bottom: thin dashed; height: 27px; border-right-width: thin"
                        valign="top">Copia Cliente
                    </td>
                </tr>
            </table>
            <div style="width: 100%; height: 10px">
            </div>
            <table style="border-right: 1px solid; border-top: 1px solid; border-left: 1px solid; width: 100%; border-bottom: 1px; text-align: left">
                <tr>
                    <td style="width: 70%; text-align: center">
                        <span style="font-size: 16px; padding-bottom: 9px; margin: 0px 0px 20px; border-bottom: #f2f2f2 1px solid; font-family: Verdana; letter-spacing: -0.03em">Comprobante de Envío
                        </span>
                    </td>
                </tr>
            </table>
            <table style="border-right: 1px solid; border-top: 1px; font-size: 13px; border-left: 1px solid; width: 100%; border-bottom: 1px; font-family: Verdana; text-align: left">
                <tr>
                    <td style="text-align: left">
                        <b>Número de referencia: <% =session("NumeroReferencia") %></b>
                    </td>
                    <td style="text-align: left">
                        <b>Código AFEX: <% =Session("CodigoAfex")%></b>
                    </td>
                    <td style="text-align: right">
                        <% =session("FechaC") %>
                        &nbsp; 
                        <%=left(Session("HoraC"), 2) & ":" & left(Mid(Session("HoraC"), 3), 2) & ":" & Right(Session("HoraC"), 2)%>
                    </td>
                </tr>
            </table>
            <div style="border-right: 1px; border-top: 1px solid; border-left: 1px; width: 100%; border-bottom: 1px; font-family: Verdana; text-align: left">
                <table style="font-size: 10px; width: 100%">
                    <tr>
                        <td style="width: 60%; text-align: left">
                            <span style="font-weight: normal; font-size: 12px; padding-bottom: 9px; margin: 0px 0px 20px; border-bottom: #f2f2f2 1px solid; font-family: Verdana, Tahoma, arial; letter-spacing: -0.03em">Remitente / Persona que envía
                            </span>
                        </td>
                        <td style="width: 40%"></td>
                        <tr>
                            <td style="width: 60%; text-align: left">Nombre:<b><% =Session("NombreRemitente")%></b>
                            </td>
                            <td style="width: 40%">Nº identificación: <b><% =Session("NumeroIdentificacionRemitente")%></b>
                            </td>
                        </tr>
                </table>
                <table style="font-size: 10px; width: 100%">
                    <tr>
                        <td style="width: 60%; text-align: left">Dirección: <b><% =Session("DireccionRemitente")%></b>
                        </td>
                        <td style="width: 40%">Tipo de identificación: <b><% =Session("TipoIdentificacion")%></b>
                        </td>
                    </tr>
                </table>
                <table style="font-size: 10px; width: 100%">
                    <tr>
                        <td style="width: 60%; text-align: left">Ciudad: <b><% =Session("CiudadRemitente")%></b>
                        </td>
                        <td style="width: 40%">Teléfono: <b><% =Session("TelefonoRemitente")%></b>
                        </td>
                    </tr>
                </table>
                <table style="font-size: 10px; width: 100%">
                    <tr>
                        <td style="width: 60%; text-align: left">Fecha de nacimiento: <b><% =Session("FechaNacimientoRemitente")%></b>
                        </td>
                        <td style="width: 40%">Nacionalidad: <b><% =Session("NacionalidadRemitente")%></b>
                        </td>
                    </tr>
                    <tr>
                        <td style="text-align: left" colspan="2">Ocupación: <b><% =Session("OcupacionRemitente")%></b>
                        </td>
                    </tr>
                </table>
            </div>
            <div style="border-right: 1px; border-top: 1px solid; border-left: 1px; width: 100%; border-bottom: 1px solid; font-family: Verdana; text-align: left">
                <table style="font-size: 10px; width: 100%; height: 71px">
                    <tr>
                        <td colspan="2" style="text-align: left">
                            <span style="font-weight: normal; font-size: 12px; padding-bottom: 9px; margin: 0px 0px 20px; border-bottom: #f2f2f2 1px solid; font-family: Verdana, Tahoma, arial; letter-spacing: -0.03em">Beneficiario / Persona que recibe:
                            </span>
                        </td>
                    </tr>
                    <tr>
                        <td colspan="2" style="text-align: left">Nombre: <b><% =Session("NombreBeneficiario")%></b>
                        </td>
                    </tr>
                    <tr>
                        <td style="width: 60%; text-align: left">Ciudad: <b><% =Session("CiudadBeneficiario")%></b>
                        </td>
                        <td style="width: 40%; text-align: left">País: <b><% =Session("PaisBeneficiario")%></b>
                        </td>
                    </tr>
                    <tr>
                        <td colspan="2" style="text-align: left">Mensaje :<b><% =Session("Mensaje")%></b>
                        </td>
                    </tr>
                    <tr>
                        <td colspan="2" style="text-align: left">&nbsp;
                        </td>
                    </tr>
                </table>
            </div>
            <table style="border-right: 1px solid; border-top: 1px; font-size: 10px; border-left: 1px solid; width: 100%; border-bottom: 1px solid; font-family: Verdana; text-align: left">
                <tr>
                    <td style="width: 60%; vertical-align: top;">
                        <%= Replace(Replace(Replace(Replace(Session("MensajeFirma"),"Ã¡","&aacute;"),"Ã­","&iacute;"),"Ã³","&oacute;"),"Ãº","&uacute;") %>
                    </td>
                    <td style="width: 40%;">
                        <table style="border-right: 1px solid; border-top: 1px; font-size: 10px; border-left: 1px solid; width: 100%; border-bottom: 1px solid; font-family: Verdana; text-align: left">
                            <tr>
                                <td style="text-align: left">
                                    <span style="font-weight: normal; font-size: 12px; padding-bottom: 9px; margin: 0px 0px 20px; border-bottom: #f2f2f2 1px solid; font-family: Verdana, Tahoma, arial; letter-spacing: -0.03em">Detalle Giro
                                    </span>
                                </td>
                            </tr>
                            <tr>
                                <td style="text-align: left">Agencia: <b><% =Session("Agencia")%></b>
                                </td>
                            </tr>
                            <tr>
                                <td style="text-align: left">
                                    <% If session("Monto")= "USD" then %>
				                        Monto: <% =Session("Monto")%> <% =FormatNumber(Session("MonedaMontoEnvio"), 2, -1, 0, -1)%>
                                    <% else %>
				                        Monto: <% =Session("Monto")%> <% =FormatNumber(Session("MonedaMontoEnvio"), 0, -1, 0, -1)%>
                                    <% End If %>
                                </td>
                            </tr>
                            <tr>
                                <td style="text-align: left" class="style2">
                                    <% If session("Monto")= "USD" then %>
	                                    Cargo:  <% =Session("Cargo")%><% =FormatNumber( Session("MonedaEnvioCargo"),  2, -1, 0, -1)%>
                                    <% Else %>
				                        Cargo:  <% =Session("Cargo")%><% =FormatNumber( Session("MonedaEnvioCargo"),  0, -1, 0, -1)%>
                                    <% End If %>
                                </td>
                            </tr>
                            <tr>
                                <td style="text-align: left">
                                    <% If  session("TotalGiro") ="US$" then %>       
                                        Total: <b style="font-size: 14px"><% =Session("TotalGiro")%> <% =FormatNumber(Session("totalRecibir"),2 , -1, 0, -1)%></b>
                                    <%else %>
			                            Total: <b style="font-size: 14px"><% =Session("TotalGiro")%> <% =FormatNumber(Session("totalRecibir"), 0, -1, 0, -1)%></b>
                                    <% End If %>   
                                </td>
                            </tr>
                            <tr>
                                <td style="text-align: left">
                                    <% IF session("MonedaRecibir")= "USD" then %>
				                        Monto a recibir: <% =session("MonedaRecibir") %>&nbsp; <% =FormatNumber(Session("MontoRecibir"), 2, -1, 0, -1)%>
                                    <% else %>
				                        Monto a recibir: <% =session("MonedaRecibir") %>&nbsp; <% =FormatNumber(Session("MontoRecibir"), 0, -1, 0, -1)%>
                                    <% End If %>
                                </td>
                            </tr>
                            <tr>
                                <td style="text-align: left">Atendido por: <% =Session("AtendidoPor")%>
                                </td>
                            </tr>
                        </table>
                    </td>
                </tr>
                <tr>
                    <td style="width: 60%">
                        <table style="width: 260px">
                            <tr>
                                <td style="border-top-width: 1px; border-left-width: 1px; border-left-color: black; width: 100%; border-top-color: black; border-bottom: black 1px solid; border-right-width: 1px; border-right-color: black"></td>
                                <tr>
                                    <td style="font-size: 9px; text-align: center">Firma Cliente</td>
                                </tr>
                        </table>

                    </td>
                    <td style="width: 40%; text-align: left">Estado Giro: <% =Session("EstadoGiro")%>
                    </td>
                </tr>
            </table>
            <table style="width: 100%">
                <tr>
                    <td align="left" style="font-size: 10px; vertical-align: top; border-bottom-style: none"
                        valign="top">Copia Agente
                    </td>
                </tr>
            </table>
        </div>
        <script type="text/javascript">
            function Imprimir() {
                if (confirm("Prepare la Impresora. Presione Aceptar para continuar")) Print();
            }
            function Print() {
                window.print()
            }
        </script>
    </form>

    <script type="text/javascript">

        function Imprimir() 
		{
            window.print();
            window.close();
        }

    </script>
</body>
</html>
<%
session.contents.remove("NumeroReferencia")
session.contents.remove("CodigoAfex")
session.contents.remove("Fecha")
session.contents.remove("Hora")
session.contents.remove("NombreRemitente")
session.contents.remove("NumeroIdentificacionRemitente")
session.contents.remove("DireccionRemitente")
session.contents.remove("TipoIdentificacion")
session.contents.remove("CiudadRemitente")
session.contents.remove("TelefonoRemitente")
session.contents.remove("FechaNacimientoRemitente")
session.contents.remove("NacionalidadRemitente")
session.contents.remove("OcupacionRemitente")
session.contents.remove("NombreBeneficiario")
session.contents.remove("CiudadBenficiario")
session.contents.remove("PaisBeneficiario")
session.contents.remove("Mensaje1")
session.contents.remove("Mensaje2")
session.contents.remove("Agencia")
session.contents.remove("Monto")
session.contents.remove("MonedaMontoEnvio")
session.contents.remove("Cargo")
session.contents.remove("MonedaEnvioCargo")
session.contents.remove("TotalGiro")
session.contents.remove("TotalRecibir")
session.contents.remove("MonedaRecibir")
session.contents.remove("MontoRecibir")
session.contents.remove("AtendidoPor")
session.contents.remove("EstadoGiro")
session.Contents.Remove("MensajeFirma")
%>