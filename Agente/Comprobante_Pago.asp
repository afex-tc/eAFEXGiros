<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Comprobante de Pago</title>    
</head>
<body onload="Imprimir()">
    <form id="form1" runat="server">
    <div style="height:13px;" >
    &nbsp;
    </div>
    <div style="font-family:Verdana">    
   <table style="border-style: solid solid none solid; border-width: 1px; width: 100%; text-align:left;">
    <tr>
        <td style="text-align: center;">                    
            <span style=" margin: 0; font-size: 16px; font-family:verdana;
margin-bottom: 20px; padding-bottom: 9px; border-bottom: 1px solid #F2F2F2;
letter-spacing: -0.035em;">Comprobante de Pago</span>
        </td>
    </tr>
    </table>
    <table style="font-size:13px; border-style: none solid none solid; border-width: 1px; width: 100%; text-align:left;">
    <tr>
        <td style="text-align: left;">
            <b>Número de referencia: <% =session("NumeroReferencia") %></b>
        </td>
        <td  style="text-align: left;">
            <b>Código AFEX: 
            <% =session("CodigoAfex") %></b>
        </td>                                   
        <td style="text-align:right">
            <% =session("Fecha") %>
            &nbsp;
            <% Session("Hora") = right("00" & Session("Hora"),6) %>
            <%=Left(Session("Hora"), 2) & ":" & Mid(Session("Hora"), 3, 2) & ":" & Right(Session("Hora"), 2)%>
        </td> 
    </tr>
</table>                        
<div style="border-style: solid none none none; border-width: 1px; width: 100%; text-align: left;">
    
    <table style="width: 100%; font-size:10px;">
        <tr>
            <td style="text-align: left; width: 60%;">
                <span style=" margin: 0; font-family: Verdana, Tahoma, arial; font-size: 12px; font-weight: bold; 
                margin-bottom: 20px; padding-bottom: 9px; border-bottom: 1px solid #F2F2F2;
                letter-spacing: -0.035em; font-weight: normal;">Beneficiario / Persona que recibe:
                </span>    
            </td>                                   
            <td style="width: 40%;">
             
            </td>   
        </tr>
        <tr>
            <td style="text-align: left; width: 60%;">
                Nombre: 
                <b>
                <% =Session("NombreBeneficiario")%>
                </b>
            </td>                                   
            <td style="width: 40%;">
                Nº identificación:
                <b>
                <% =session("NumeroIdentificacionBeneficiario") %>
                </b>
            </td>   
        </tr>
    </table>
    <table style="width: 100%; font-size:10px;"> 
        <tr>
            <td style="text-align: left; width: 60%;">
                Dirección: 
                <b>
                <% =Session("DireccionBeneficiario")%>
                </b>
            </td>
            <td style="width: 40%;">
                Tipo de identificación: 
                <b>
                <% =Session("TipoIdentificacionBeneficiario")%>
                </b>
            </td>
        </tr>
    </table>
    <table style="width: 100%; font-size:10px;"> 
        <tr>
            <td style="text-align: left; width: 60%;">
                Ciudad:
                <b>
                <% =session("CiudadBeneficiario") %>
                </b>
            </td>
            <td style="width: 40%;">
                Teléfono:
                <b>
                <% =Session("TelefonoBeneficiario")%>
                </b>
            </td>
        </tr>
    </table>                   
    <table style="width: 100%; font-size:10px;"> 
        <tr>
            <td style="text-align: left; width: 60%;">
                Fecha de nacimiento:
                <b>
                <% =Session("FechaNacimientoBeneficiario")%>
                </b>
            </td>
            <td style="width: 40%;">
                Nacionalidad:
                <b>
                <% =session("NacionalidadBeneficiario") %>
                </b>
            </td>                                
        </tr>
        <tr>
            <td style="text-align: left; " colspan="2">
                Ocupación:
                <b>
                <% =Session("OcupacionBeneficiario")%>
                </b>
            </td>
        </tr>
    </table>
</div>                      
<div style="border-style: solid none solid none; border-width: 1px; width: 100%; text-align: left;">
    
    <table style="width: 100%; height: 71px; font-size:10px;">
        <tr>
            <td colspan="2" style="text-align: left;">
                <span style=" margin: 0; font-family: Verdana, Tahoma, arial; font-size: 12px; font-weight: bold; 
                margin-bottom: 20px; padding-bottom: 9px; border-bottom: 1px solid #F2F2F2;
                letter-spacing: -0.035em; font-weight: normal;">Remitente / Persona que envía:</span>                                                           
            </td>
        </tr>        
        <tr>
            <td colspan="2" style="text-align: left;">
                Nombre: 
                <b>
                <% =session("NombreRemitente") %>
                </b>
            </td>
        </tr>
        <tr>
            <td style="text-align: left; width: 60%;">
                Ciudad:
                <b>
                <% =session("CiudadRemitente") %>
                </b>
            </td>
            <td style="text-align: left; width: 40%;">
                País:
                <b>
                    <% =Session("PaisRemitente")%>
                </b>
            </td>
        </tr>
        <tr>
            <td colspan="2" style="text-align: left;">
                Mensaje:
                <b><% =session("Mensaje") %></b>
            </td>
        </tr>
        <tr>
            <td colspan="2" style="text-align: left;">
                &nbsp;</td>
        </tr>
        <tr>
            <td colspan="2" style="text-align: left;">
                &nbsp;</td>
        </tr>
        <tr>
            <td colspan="2" style="text-align: left;">
                &nbsp;</td>
        </tr>
    </table>                        
</div>    
<div style="border-width:1px;">
<table style=" border-style: none solid solid solid; border-width: 1px; width: 100%; text-align: left; font-family:Verdana; font-size:10px;">
    <tr>
        <td style=" width: 60%;"/>
        <td style=" text-align: left; width: 40%; ">
            <span style=" margin: 0; font-family: Verdana, Tahoma, arial; font-size: 12px; font-weight: bold; 
margin-bottom: 20px; padding-bottom: 9px; border-bottom: 1px solid #F2F2F2;
letter-spacing: -0.035em; font-weight: normal;">Detalle del pago:</span>    
        </td>
    </tr>
    <tr>
        <td/>
            <td style=" text-align: left; " >
            Agencia:
            <b><% =session("Agencia") %></b>
        </td>
    </tr>
    <tr>
        <td style=" width: 60%; text-align:left;" valign="bottom"/>
        <td style=" text-align: left; width: 40%; ">
            Monto a recibir:
            <b style="font-size:14px;">
            <% if Session("MonedaRecibir")="USD" then %>
                <%=Session("MonedaRecibir")%>&nbsp;
                <% =FormatNumber(Session("MontoRecibir"), 2, -1, 0, -1)%>
            <% else %>
                <%=Session("MonedaRecibir")%>&nbsp;
                <% =FormatNumber(Session("MontoRecibir"), 0, -1, 0, -1)%>
              <% End IF %>
            </b>            
        </td>
    </tr>                       
    <tr>
        <td style=" width: 60%; text-align:left;" valign="bottom" />
        <td style=" text-align: left; width: 40%; ">
              Atendido por:
            <b>
                <% =session("AtendidoPor") %>
            </b>
        </td>
    </tr>                       
    <tr>
        <td style=" width: 60%; text-align:left;" valign="bottom" />
        <td style=" text-align: left; width: 40%; " />
    </tr>                       
    <tr>
        <td style=" width: 60%; text-align:left;" valign="bottom" />
            &nbsp;<td style=" text-align: left; width: 40%; " />
            &nbsp;</tr>                       
    <tr>
        <td style=" width: 60%; text-align:left;" valign="bottom" />
            &nbsp;<td style=" text-align: left; width: 40%; " />
            &nbsp;</tr>                       
    <tr>
        <td style=" width: 60%; text-align:Left">        
        <table style="width: 260px;">
            <tr>
                <td style=" border-width: 1px; border-color:Black; border-bottom-style:solid; width:100%"/>
            </tr>
            <tr>
                <td style="text-align:center;">
                    Firma Cliente</td>
            </tr>            
            </table>
        
        </td>
        <td style=" text-align: left; width: 40%;" />
    </tr>
</table>
</div>
 <table style="width:100%;">
            <tr>
                <td align="left" style="vertical-align:top; border-bottom-style: dashed; border-width: thin" 
                    valign="top">
                    <table>
                    <tr>
                       <td valign="top" style="vertical-align:top; font-size:10px;" >
                           Copia Cliente
                       </td>
                    </tr>
                    </table>
                </td>
        </tr>
        </table>
    <span style="font-size:9px;">&nbsp;</span> 
   <table style="border-style: solid solid none solid; border-width: 1px; width: 100%; text-align:left;">
    <tr>
        <td style="text-align: center;">                    
            <span style=" margin: 0; font-size: 16px; font-family:verdana;
margin-bottom: 20px; padding-bottom: 9px; border-bottom: 1px solid #F2F2F2;
letter-spacing: -0.035em;">Comprobante de Pago</span>
        </td>
    </tr>
    </table>
    <table style="font-size:13px; border-style: none solid none solid; border-width: 1px; width: 100%; text-align:left;">
    <tr>
        <td style="text-align: left;">
            <b>Número de referencia: <% =session("NumeroReferencia") %></b>
        </td>
        <td  style="text-align: left;">
            <b>Código AFEX: 
            <% =session("CodigoAfex") %></b>
        </td>                                   
        <td style="text-align:right">
            <% =session("Fecha") %>
            &nbsp;
            <% Session("Hora") = right("00" & Session("Hora"),6) %>
            <%=Left(Session("Hora"), 2) & ":" & Mid(Session("Hora"), 3, 2) & ":" & Right(Session("Hora"), 2)%>           
        </td> 
    </tr>
</table>                        
<div style="border-style: solid none none none; border-width: 1px; width: 100%; text-align: left;">
    
    <table style="width: 100%; font-size:10px;">
        <tr>
            <td style="text-align: left; width: 60%;">
                <span style=" margin: 0; font-family: Verdana, Tahoma, arial; font-size: 12px; font-weight: bold; 
                margin-bottom: 20px; padding-bottom: 9px; border-bottom: 1px solid #F2F2F2;
                letter-spacing: -0.035em; font-weight: normal;">Beneficiario / Persona que recibe:
                </span>    
            </td>                                   
            <td style="width: 40%;">
             
            </td>   
        </tr>
        <tr>
            <td style="text-align: left; width: 60%;">
                Nombre: 
                <b>
                <% =Session("NombreBeneficiario")%>
                </b>
            </td>                                   
            <td style="width: 40%;">
                Nº identificación:
                <b>
                <% =session("NumeroIdentificacionBeneficiario") %>
                </b>
            </td>   
        </tr>
    </table>
    <table style="width: 100%; font-size:10px;"> 
        <tr>
            <td style="text-align: left; width: 60%;">
                Dirección: 
                <b>
                <% =Session("DireccionBeneficiario")%>
                </b>
            </td>
            <td style="width: 40%;">
                Tipo de identificación: 
                <b>
                <% =Session("TipoIdentificacionBeneficiario")%>
                </b>
            </td>
        </tr>
    </table>
    <table style="width: 100%; font-size:10px;"> 
        <tr>
            <td style="text-align: left; width: 60%;">
                Ciudad:
                <b>
                <% =session("CiudadBeneficiario") %>
                </b>
            </td>
            <td style="width: 40%;">
                Teléfono:
                <b>
                <% =Session("TelefonoBeneficiario")%>
                </b>
            </td>
        </tr>
    </table>                   
    <table style="width: 100%; font-size:10px;"> 
        <tr>
            <td style="text-align: left; width: 60%;">
                Fecha de nacimiento:
                <b>
                <% =Session("FechaNacimientoBeneficiario")%>
                </b>
            </td>
            <td style="width: 40%;">
                Nacionalidad:
                <b>
                <% =session("NacionalidadBeneficiario") %>
                </b>
            </td>                                
        </tr>
        <tr>
            <td style="text-align: left; " colspan="2">
                Ocupación:
                <b>
                <% =Session("OcupacionBeneficiario")%>
                </b>
            </td>
        </tr>
    </table>
</div>                      
<div style="border-style: solid none solid none; border-width: 1px; width: 100%; text-align: left;">
    
    <table style="width: 100%; height: 71px; font-size:10px;">
        <tr>
            <td colspan="2" style="text-align: left;">
                <span style=" margin: 0; font-family: Verdana, Tahoma, arial; font-size: 12px; font-weight: bold; 
                margin-bottom: 20px; padding-bottom: 9px; border-bottom: 1px solid #F2F2F2;
                letter-spacing: -0.035em; font-weight: normal;">Remitente / Persona que envía:</span>                                                           
            </td>
        </tr>        
        <tr>
            <td colspan="2" style="text-align: left;">
                Nombre: 
                <b>
                <% =session("NombreRemitente") %>
                </b>
            </td>
        </tr>
        <tr>
            <td style="text-align: left; width: 60%;">
                Ciudad:
                <b>
                <% =session("CiudadRemitente") %>
                </b>
            </td>
            <td style="text-align: left; width: 40%;">
                País:
                <b>
                    <% =Session("PaisRemitente")%>
                </b>
            </td>
        </tr>
        <tr>
            <td colspan="2" style="text-align: left;">
                Mensaje:
                <b><% =session("Mensaje") %></b>
            </td>
        </tr>
        <tr>
            <td colspan="2" style="text-align: left;"/>
        </tr>
        <tr>
            <td colspan="2" style="text-align: left;"/>
        </tr>
        <tr>
            <td colspan="2" style="text-align: left;"/>
        </tr>
    </table>                        
</div>    
<div style="border-width:1px;">
<table style=" border-style: none solid solid solid; border-width: 1px; width: 100%; text-align: left; font-family:Verdana; font-size:10px;">
    <tr>
        <td style=" width: 60%;"/>
        <td style=" text-align: left; width: 40%; ">
            <span style=" margin: 0; font-family: Verdana, Tahoma, arial; font-size: 12px; font-weight: bold; 
margin-bottom: 20px; padding-bottom: 9px; border-bottom: 1px solid #F2F2F2;
letter-spacing: -0.035em; font-weight: normal;">Detalle del pago:</span>    
        </td>
    </tr>
    <tr>
        <td/>
            <td style=" text-align: left; " >
            Agencia:
            <b><% =session("Agencia") %></b>
        </td>
    </tr>
    <tr>
        <td style=" width: 60%; text-align:left;" valign="bottom"/>
        <td style=" text-align: left; width: 40%; ">
            Monto a recibir:
            <b style="font-size:14px;">
            <% If Session("MonedaRecibir")="USD" then %>
                <%=Session("MonedaRecibir")%>&nbsp;
                <% =FormatNumber(Session("MontoRecibir"), 2, -1, 0, -1)%>
            <% else %>
                <%=Session("MonedaRecibir")%>&nbsp;
                <% =FormatNumber(Session("MontoRecibir"), 0, -1, 0, -1)%>
              <% End IF %>
            </b>            
        </td>
    </tr>                       
    <tr>
        <td style=" width: 60%; text-align:left;" valign="bottom" />
        <td style=" text-align: left; width: 40%; ">
              Atendido por:
            <b>
                <% =session("AtendidoPor") %>
            </b>
        </td>
    </tr>                       
    <tr>
        <td style=" width: 60%; text-align:left;" valign="bottom" />
        <td style=" text-align: left; width: 40%; " />
    </tr>                       
    <tr>
        <td style=" width: 60%; text-align:Left">        
        <table style="width: 260px;">
            <tr>
                <td style=" border-width: 1px; border-color:Black; border-bottom-style:solid; width:100%"/>
            </tr>
            <tr>
                <td style="text-align:center;">
                    Firma Cliente</td>
            </tr>            
            </table>
        
        </td>
        <td style=" text-align: left; width: 40%;" />
    </tr>
</table>
</div>
<table>
<tr style=" width: 100%">
<td style=" text-align: left; vertical-align:top; width: 100%; font-size:10px;">
Copia Agente
</td>
</tr>
</table>
	<script type="text/javascript"  language="javascript" >
    function Imprimir()
    {        		
        if (confirm("Prepare la Impresora. Presione Aceptar para continuar")) Print();        			    		        
    }        
       function Print()
       {
            window.print()
       }
	</script></div>
    </form>
</body>


<script type="text/javascript">

function Imprimir()
{
    window.print();
    window.close();
}

</script>
<%
    Session.Contents.Remove("NumeroReferencia")
    Session.Contents.Remove("Fecha")
    Session.Contents.Remove("Hora")
    Session.Contents.Remove("NumeroAutorizacion")
    Session.Contents.Remove("NumeroReferencia")
    Session.Contents.Remove("CodigoAfex")
    Session.Contents.Remove("NombreBeneficiario")
    Session.Contents.Remove("NumeroIdentificacionBeneficiario")
    Session.Contents.Remove("DireccionBeneficiario")
    Session.Contents.Remove("TipoIdentificacionBeneficiario")
    Session.Contents.Remove("CiudadBeneficiario")
    Session.Contents.Remove("TelefonoBeneficiario")
    Session.Contents.Remove("FechaNacimientoBeneficiario")
    Session.Contents.Remove("NacionalidadBeneficiario")
    Session.Contents.Remove("OcupacionBeneficiario")
    Session.Contents.Remove("NombreRemitente")
    Session.Contents.Remove("CiudadRemitente")
    Session.Contents.Remove("PaisRemitente")
    Session.Contents.Remove("Agencia")
    
%>

</html>