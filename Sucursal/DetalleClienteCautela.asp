<%@  language="VBScript" %>
<%
	Response.Expires = 0
%>
<!--#INCLUDE virtual="/Compartido/Rutinas.asp" -->
<!--#INCLUDE virtual="/Compartido/Errores.asp" -->
<!--#INCLUDE virtual="/sucursal/Rutinas.asp" -->
<!--#INCLUDE virtual="/Compartido/RutinasEncriptar.asp" -->
<%
	nCodigoCliente = Request("cc")
	
	Dim sRut
	Dim sNumeroSerieID
	Dim sPasaporte
	Dim sEstado
	Dim sNombres
	Dim sNombreCompleto
	Dim sApellidoP
	Dim sApellidoM
	Dim nTipoCliente
	Dim nDiasRetencion
	Dim iRiesgo
	Dim sFechaCreacion
    Dim sNombrePaisPasaporte

    Dim bMostrarMenu
	bMostrarMenu = True
	If Session("NombreUsuarioOperador") = "" Then
	    bMostrarMenu = False
	End If
    Dim rsHistoria
    dim nHistoriaCompleta
	
	On Error Resume Next
	If Not IsNull(nCodigoCliente) And nCodigoCliente <> "" then
	    CargarCliente
	End If
    Set rsHistoria = ObtenerHistoria(1)
		
	Sub CargarCliente()
	
	    sSQL = "EXECUTE [ObtenerCLienteCautela] "
        sSQL = sSQL & "@Codigo = '" &nCodigoCliente &"'"
        
        Dim rsCliente		
	    Set rsCliente = EjecutarSQLCliente(Session("afxCnxCorporativa"), sSQL)

	    If rsCliente.Fields.Count > 0 Then
            If Not IsNull(rsCliente("rut")) And rsCliente("rut") <> "" then
                sRut = ValorRut(rsCliente("rut"))
            Else
                sRut = Empty
            End If
	        sNumeroSerieID = EvaluarVar(rsCliente("serieidentificacion"), "")
		    sPasaporte= EvaluarVar(rsCliente("pasaporte"), "")
            sNombrePaisPasaporte = EvaluarVar(rsCliente("nombrepaispasaporte"), "")
		    sEstado= EvaluarVar(rsCliente("estado"), "")	
		    Session("sEstadoOriginal") = EvaluarVar(sEstado, "")		
		    sNombres = EvaluarVar(rsCliente("Nombres"), "")
		    sNombreCompleto = EvaluarVar(rsCliente("Nombre"), "")	
		    sApellidoP = EvaluarVar(rsCliente("apellido_paterno"), "")
		    sApellidoM = EvaluarVar(rsCliente("apellido_materno"), "")
		    nTipoCliente = EvaluarVar(rsCliente("tipo"), 0)
		    nDiasRetencion = cInt(0 & EvaluarVar(rsCliente("dias_retencion"), 0))
		    iRiesgo = rsCliente("nivelriesgo")
		    If IsNull(rsCliente("fecha_creacion")) Then
			    sFechaCreacion = Empty
		    Else
			    sFechaCreacion = FormatDateTime(rsCliente("fecha_creacion"), 2)

		    End If
		    
        Else
            sCodigo = 0
        End If
        
        If Err.number <> 0 Then
		    response.Redirect "../Compartido/Error.asp?Titulo=Error en Agregar Cliente&Number=" & Err.number   & "&Source=" & Err.Source & "&Description=" & Err.description
	    End If
		
	End Sub

    Function ObtenerHistoria(HistoriaCompleta)
	   Dim rsATC
	   Dim sSQL
	   Dim Completa

	   Set ObtenerHistoria = Nothing
	   On Error Resume Next
       sSQL = "EXEC ObtenerHistoriaCliente "  & nCodigoCliente & ", " & HistoriaCompleta
	   Set rsATC = EjecutarSQLCliente(Session("afxCnxCorporativa"), sSQL)

	   If Err.Number <> 0 Then 
			Set rsATC = Nothing
			MostrarErrorMS "Obtener Historia 1"
		End If
	   
	   Set rsATC.ActiveConnection = Nothing
	   Set ObtenerHistoria = rsATC

	   Set rsATC = Nothing	
	End Function
	
	

%>
<html>
<head>
    <style>
        a:hover
        {
            color: blue;
        }
        INPUT.dINPUT
        {
            border-right: gray 1px solid;
            border-top: silver 1px solid;
            border-left: silver 1px solid;
            border-bottom: gray 1px solid;
        }
    </style>
    <meta name="GENERATOR" content="Microsoft Visual Studio 6.0">
    <meta http-equiv="CACHE-CONTROL" content="NO-CACHE, must-revalidate">
    <meta http-equiv="Pragma" content="no-cache">
    <title>Configuración Consulta de Giros</title>
    <link rel="stylesheet" type="text/css" href="../Estilos/Cliente.css">
    <link rel="stylesheet" type="text/css" href="../Estilos/AFX2004.css">
</head>

<script language="VBScript">
    Sub window_onload()
		<%	If sRut = Empty Then %>
				tdRut.style.display = "none"
				frmCliente.cmdValidarRegistro.style.display = "none"
		<%	Else %>
				tdPasaporte.style.display = "none"				
		<%	end If %>
		
	End Sub
	
</script>

<!--INCLUDE virtual="/Compartido/Encabezado.htm" -->
<body id="bb" border="0" style="margin: 2 2 2 2">
    <form id="frmCliente" method="post">
    <input type="hidden" name="txtTipoCliente" value="<%=nTipoCliente%>">
    <!--<input type="hidden" name="txtNombreCompleto" value="<%=sNombreCompleto%>">-->
    <table class="Borde" id="tabConsulta" border="0" cellpadding="0" cellspacing="0"
        style="height: 150px; width: 100%; background-color: #f4f4f4">
        <tr height="40" style="background-color: #ffeeaa; #ffdd77; #e1e1e1">
            <td colspan="3" style="font-size: 16pt">
                &nbsp;&nbsp;<%=sApellidoP%>&nbsp;<%=sApellidoM%>&nbsp;<%=sNombres%>
            </td>
        </tr>
        <tr height="1" style="background-color: silver">
            <td colspan="3">
            </td>
        </tr>
        <tr height="4">
            <td colspan="3">
            </td>
        </tr>
        <tr>
            <td colspan="3">
                <table width="100%">
                    <tr>
                        <td>
                            <table style="font-size: 8pt;">
                                <tr>
                                    <td id="tdRut" colspan="2">
                                        Rut<br>
                                        <input class="dInput" name="txtRut" style="width: 100px" value="<%=FormatoRut(sRut)%>"
                                            disabled />
                                        <input class="dInput" name="txtNumeroSerieRut" style="width: 100px;" value="<%=sNumeroSerieID%>"
                                            disabled />
                                    </td>
                                    <td id="tdPasaporte">
                                        Pasaporte<br>
                                        <input class="dInput" name="txtPasaporte" style="width: 100px" value="<%=sPasaporte%>"
                                            disabled />
                                        <input class="dInput" name="txtNombrePaisPasaporte" value="<%=sNombrePaisPasaporte%>"
                                            style="width: 100px" disabled />
                                    </td>
                                    <td>
                                        &nbsp;<br />
                                        <input type="button" name="cmdValidarRegistro" value="Validar Rut" />
                                    </td>
                                    <td>
                                        Código<br />
                                        <input class="dInput" name="txtCodigoCliente" value="<%=Request("cc")%>" style="width: 60px"
                                            disabled />
                                    </td>
                                    <td>
                                        Fecha Creación<br>
                                        <input class="dInput" name="txtFechaCreacion" value="<%=sFechaCreacion%>" style="width: 90px"
                                            disabled />
                                    </td>
                                    
                                </tr>
                            </table>
                        </td>
                        <td align="right">
                            <table>
                                <tr>
                                    <%If bMostrarMenu Then %>
                                    
                                    <td colspan="2">
                                        <!--#INCLUDE virtual="/Sucursal/MenuDetalleClienteCautela.htm" -->
                                    </td>
                                    <%End If%>
                                </tr>
                            </table>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
        <tr height="10">
            <td>
            </td>
        </tr>
        <tr>
            <td>
                <table cellspacing="1" cellpadding="1" swidth="30%" style="font-family: Verdana;
                    font-size: 10pt; position: relative; top: 0px; sborder: 1; background-color: silver;">
                    <tr height="22">
                        <td id="tdDocumento" style="background-color: #ffffcc; #e1e1e1; cursor: hand">
                            <b>&nbsp;&nbsp;Datos del Cliente en Cautela&nbsp;&nbsp;</b>
                        </td>
                        <td id="tdModoActualizacion" style="background-color: #ccddee; #e1e1e1; display: none">
                            <b>&nbsp;&nbsp;Modo Actualización&nbsp;&nbsp;</b>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
        <tr height="1" style="background-color: silver">
            <td colspan="3">
            </td>
        </tr>
        <tr height="10">
            <td>
            </td>
        </tr>
        <tr>
            <td>
                <table>
                    <tr id="trPersona" height="20">
                        <td>
                            <table>
                                <tr>
                                    <td>
                                        Apellido Paterno<br>
                                        <input name="txtApellidoP" id="txtApellidoP" onkeypress="IngresarTexto(2)" style="width: 180px"
                                            maxlength="20" value="<%=sApellidoP%>" disabled=disabled>
                                    </td>
                                    <td>
                                        Apellido Materno<br>
                                        <input name="txtApellidoM" id="txtApellidoM" onkeypress="IngresarTexto(2)" style="width: 180px"
                                            maxlength="20" value="<%=sApellidoM%>" disabled=disabled>
                                    </td>
                                    <td>
                                        Nombres<br>
                                        <input name="txtNombres" id="txtNombres" style="width: 308px" onkeypress="IngresarTexto(2)"
                                            maxlength="30" value="<%=sNombres%>" disabled=disabled>
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
        <tr>
            <td colspan="3">
                <%
                %>
                <table class="Borde" cellspacing="0" cellpadding="0" style="font-family: Verdana;
                    font-size: 10pt; position: relative; top: 0px; background-color: silver;">
                    <tr height="22">
                        <td id="tdHistoria" style="background-color: #ffffcc; #e1e1e1; cursor: hand">
                            <b>&nbsp;&nbsp;Historia&nbsp;&nbsp;</b>
                        </td>
                    </tr>
                </table>
                <table cellspacing="1" cellpadding="1" width="100%" id="tbHistoria" align="center"
                    style="color: #505050; font-family: Verdana; font-size: 10px; position: relative;
                    top: 0px; border: 1px; background-color: silver; ">
                    <tr style="height: 20px" align="center">
                        <td style="background-color: #e1e1e1" width="10%">
                            <b>Fecha</b>
                        </td>
                        <!--<td style="background-color: #e1e1e1" WIDTH="10%">
								<b>Hora</b>
							</td>-->
                        <td style="background-color: #e1e1e1" width="80%">
                            <b>Detalle</b>
                        </td>
                    </tr>
                    <%
							Dim sDetalle, sHora, sBGColor
							Do Until rsHistoria.EOF
								'sHora = Right("000000" & rsHistoria("hora"), 6)
								'sHora = Left(sHora, 2) & ":" & Mid(sHora, 3, 2) & ":" & Right(sHora, 2)
								Select Case rsHistoria("tipo")
									Case 1	'Informacion
										sBGColor = "#ddeeff"
										sColor = "#ddeeff"
									Case 2	'Advertencia
										sBGColor = "#ffffee"
										sColor = "#ffff00"
									Case 3	'Peligro
										sBGColor = "#ffddee"
										sColor = "#ffddee"
									Case Else
										sBGColor = "white"
										sColor = "#ff0000"
									End Select
									sHora= rsHistoria("hora")
                    %>
                    <tr style="background-color: <%=sBGColor%>; color: black; <%=sColor%>; height: 16px">
                        <td>
                            <%=rsHistoria("fecha") %>
                        </td>
                        <!--<td><%=sHora%></td>-->
                        <td>
                            <%=rsHistoria("descripcion") & " (" & rsHistoria("NombreUsuario") & ")"%>
                        </td>
                    </tr>
                    <% 
								rsHistoria.MoveNext
							Loop 
							Set rsHistoria = Nothing
							Set rsCliente = Nothing
                    %>
                </table>
            </td>
        </tr>
    </table>
    </form>
    <!--#INCLUDE virtual="/Compartido/Rutinas.htm" -->
</body>
<head>
    <meta http-equiv="CACHE-CONTROL" content="NO-CACHE, must-revalidate">
    <meta http-equiv="Pragma" content="no-cache">
</head>
</html>
