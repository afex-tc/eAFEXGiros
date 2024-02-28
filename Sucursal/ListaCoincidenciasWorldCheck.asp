<!--#INCLUDE virtual="/Compartido/Rutinas.asp" -->
<!--#INCLUDE virtual="/Sucursal/Rutinas.asp" -->
<!--#INCLUDE virtual="/Compartido/Rutinas.htm" -->
<%
    Dim sSQL, rs
    Dim sMensaje
    Dim sPais
    Dim sPorcentage
    Dim sNombres
    
    On Error Resume Next
    
    set rs = nothing
    
    sPais = Request("Pais")
    If sPais = "" Then sPais = "NULL"
    sPorcentage = request("porcentaje")
    If sPorcentage = "" Then sPorcentage = Session("PorcentageCoincidenciaWorldCheck") '"90"
    sNombres = trim(request("nombres"))
    If  trim(request("apellidos")) <> "" Then sNombres = trim(sNombres) & " " & trim(request("apellidos"))
        
    if trim(sNombres) <> "" then
		sSQL = "exec cumplimiento.mostrarcoincidenciasworldcheck " & evaluarstr(sNombres) & _
																", NULL" & _
																", " & cint(100) - cint("0" & sPorcentage) & _
																", " & sPais
		set rs = ejecutarsqlcliente(Session("afxCnxCorporativa"), sSQL)
    
		if err.number <> 0 then
		    sMensaje = "Error al consultar coincidencias WorldCheck. " & err.description		
		elseif rs.eof then
			sMensaje = "No se encontraron coincidencias."
			' JFMG 01-02-2011 se agrega una historia informando la consulta de worldcheck
			If trim(request("cc")) <> "" Then
				AgregarHistoria trim(request("cc")), "Cliente consultado en WorldCheck sin coincidencias (" & trim(sNombres) & ").", 1, 0
			End If			
			' FIN JFMG 01-02-2011
			
		elseif not rs.eof then
			' JFMG 01-02-2011 se agrega una historia informando la consulta de worldcheck
			If trim(request("cc")) <> "" Then
				AgregarHistoria trim(request("cc")), "Cliente consultado en WorldCheck con coincidencias (" & trim(sNombres) & ").", 1, 0
				 Dim Cuerpo, Asunto
				'Se envía mail MS 10-07-2013
				Asunto = "Consulta de coincidencias WorlCheck de cliente " & trim(sNombres)
				
				Cuerpo = "Estimado,<br /><br />Se informa que el usuario " & Session("NombreUsuarioOperador") & _
						 " ha realizado la siguiente acción con el cliente&nbsp;" & ucase(trim(sNombres)) & ":" & _
						 "<br /><br /><b>Consulta de coincidencias WorlCheck con coincidencias.</b>" & _
						 "<br /><br /> Atte,<br /><br />Servicio de Mensajería Afex."
				
				EnviarEMailBD 8, 29,Session("AmbienteServidorCorreo"),Asunto, Cuerpo
				'MS 10-07-2013
			End If			
			' FIN JFMG 01-02-2011
		
		end if
	else
		sMensaje = "No se especificó Nombre o Apellido"
    end if
    
    response.expires = 0
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" >
<head>
    <title>AFEX</title>
    <link rel="stylesheet" type="text/css" href="../Estilos/Cliente.css">
    <link rel="stylesheet" type="text/css" href="../Estilos/AFX2004.css">
</head>
<body>

	<script language="vbscript">
	<!-- 
	
		sub cmdVerificarWorldCheck_onClick()
			if trim(frmLista.txtOtrosNombres.value) <> "" then
			    window.divProcesando.style.display = "block"
				window.navigate "listacoincidenciasworldcheck.asp" & _
																"?cc=" & "<%=Request("cc")%>" & _
																"&nombres=" & trim(frmLista.txtOtrosNombres.value) & _
																"&porcentaje=" & trim(frmLista.txtPorcentaje.value) & _
																"&pais=" & trim(frmLista.txtPais.value)
			else
				msgbox "Debe ingresar los datos solicitados.",,"AFEX"
			end if
		end sub
		
		sub txtPorcentaje_onblur()
			if cint("0" & frmLista.txtPorcentaje.value) < 1 or cint("0" & frmLista.txtPorcentaje.value) > 100 _
			or cint("0" & frmLista.txtPorcentaje.value) = 0 then
				frmLista.txtPorcentaje.value = 100		
			end if
		end sub
			
	-->	
	</script>

    <style >
        
        .ContenedorModal
        {
           display: none;
           width: 100%;
           height: 100%; 
           top: 0;
           left: 0;
           
           position: fixed;
           z-index: 5000;
        }
    
        .Modal
        {
           width: 100%;
           height: 100%; 
           top: 0;
           left: 0;
           
           position: fixed;
           background-color: White;
           z-index: 5000;
           opacity: .5;
           filter: alpha(opacity=50);
        }

        .InteriorModal300
        {
           position: absolute;
           background-color: #eeeeee;
           border-radius: 4px;
           
           padding: 2px;
           
           width: 300px;
           top: 55%;
           left: 50%;
           margin-top: -100px;
           margin-left: -200px;   
           
           z-index: 5001;
        }
    </style>


    <!--INCLUDE file="../Compartido/Encabezado.htm" -->
	<form name="frmLista" method="post" action="">
    <div id="bb" border="0" style="margin: 2 2 2 2" >
        <table class="Borde" id="tabConsulta" BORDER="0" cellpadding="0" cellspacing="0" style="HEIGHT: 150px; width:100%; background-color: #f4f4f4">
			<tr height="40" style="background-color: #ffeeaa; #ffdd77; #e1e1e1">
				<td colspan="3" style="font-size: 16pt">
					&nbsp;&nbsp;Lista de coincidencias WorldCheck <b><%=ucase(trim(sNombres))%> </b>
				</td>
			</tr>
            <tr>
                <td>&nbsp;</td>
            </tr>
			
			<tr>
				<td>
					<table>
						<tr>
							<td>Nombres:</td>
							<td>
								<input name="txtOtrosNombres" style="width: 300px" OnKeyPress="IngresarTexto(3)" maxlength="80" value="<%=sNombres%>">
							</td>
						</tr>
						<tr>
							<td>País:</td>
							<td>
								<input name="txtPais" style="width: 300px" OnKeyPress="IngresarTexto(3)" maxlength="80" value="<%=Request("pais")%>">
							</td>
						</tr>
						<tr>
							<td>Porcentaje de comparación:</td>
							<td>
								<input name="txtPorcentaje" style="width: 28px" OnKeyPress="IngresarTexto(1)" maxlength="3" value="<%=sPorcentage%>">
								&nbsp;%
							</td>
							
						</tr>
						<tr>
							<td align="right" colspan="2">
								<input type="button" name="cmdVerificarWorldCheck" value="Verificar" />
							</td>
						</tr>
					</table>
				</td>
			</tr>				
			
			<tr>
                <td>&nbsp;</td>
            </tr>
            <tr>
				<td align="right">
					<%if request("cc") <> "" then%>
						<input type="button" value="Volver" onclick="javascript:window.navigate('http:DetalleCliente.asp?cc=<%=Request("cc")%>');" / id=button1 name=button1>
					<%end if%>
				</td>
			</tr>
			<tr>
				<td>
					
					<table cellspacing="1" cellpadding="4" width="100%" ID="tbDocumento" STYLE="COLOR: #505050; FONT-FAMILY: Verdana; FONT-SIZE: 10px; POSITION: relative; TOP: 0px; border: 1; background-color: silver; ; display: nones">
					    <tr style="height: 20px" align="center">
							<td style="background-color: #e1e1e1">Nombre</td>                       
					    </tr>
					    
					    <%if sMensaje <> "" then%>
							<tr bgcolor="white" style="color: blue; sbackground-color: white; #f1f1f1; height: 16px; cursor: hand" onMouseOver="javascript:this.bgColor='#f4f4f4'; " onMouseOut="javascript:this.bgColor='white'">
								<td colspan="6" align="center"><%=sMensaje%></td>						
							</tr>
					    <%elseif not rs is nothing then%>
					        <%if not rs.eof then %>					            
					            <%do while not rs.eof %>
					                <tr bgcolor="white" style="color: blue; sbackground-color: white; #f1f1f1; height: 10px; cursor: hand" onMouseOver="javascript:this.bgColor='#f4f4f4'; " onMouseOut="javascript:this.bgColor='white'">
					                    <td height="10px"><%=trim(rs("FirstName")) & " " & trim(rs("LastName"))%></td>
					                    
					                    <td height="10px">
											<table height="10px">
												<%if rs("Aliases") <> "" then%>
												<tr>
													<td style="background-color: #e1e1e1">Aliases:&nbsp; </td>
													<td><%=rs("Aliases")%></td>
												</tr>
												<%end if%>
												<%if rs("title") <> "" then%>
												<tr>
													<td style="background-color: #e1e1e1">Title:&nbsp; </td>
													<td><%=rs("title")%></td>
												</tr>
												<%end if%>
												<%if rs("deceased") <> "" then%>
												<tr>
													<td style="background-color: #e1e1e1">Deceased:&nbsp; </td>
													<td><%=rs("deceased")%></td>
												</tr>
												<%end if%>
												<%if rs("countries")  <> "" then%>
												<tr>
													<td style="background-color: #e1e1e1">Countries:&nbsp; </td>
													<td><%=rs("countries")%></td>
												</tr>
												<%end if%>
												<%if rs("keywords")  <> "" then%>
												<tr>
													<td style="background-color: #e1e1e1">KeyWords:&nbsp; </td>
													<td><%=rs("keywords")%></td>
												</tr>
												<%end if%>
												<%if rs("editor")  <> "" then%>
												<tr>
													<td style="background-color: #e1e1e1">Editor:&nbsp; </td>
													<td><%=rs("editor")%></td>
												</tr>
												<%end if%>
												<%if rs("AlternativeSpelling")  <> "" then%>
												<tr>
													<td style="background-color: #e1e1e1">Alternative Spelling:&nbsp; </td>
													<td><%=rs("AlternativeSpelling")%></td>
												</tr>
												<%end if%>
												<%if rs("position")  <> "" then%>
												<tr>
													<td style="background-color: #e1e1e1">Position:&nbsp; </td>
													<td><%=rs("position")%></td>
												</tr>
												<%end if%>
												<%if rs("passports")  <> "" then%>
												<tr>
													<td style="background-color: #e1e1e1">Passports:&nbsp; </td>
													<td><%=rs("passports")%></td>
												</tr>
												<%end if%>
												<%if rs("companies")  <> "" then%>
												<tr>
													<td style="background-color: #e1e1e1">Companies:&nbsp; </td>
													<td><%=rs("companies")%></td>
												</tr>
												<%end if%>
												<%if rs("externalsource")  <> "" then%>
												<tr>
													<td style="background-color: #e1e1e1">External Source:&nbsp; </td>
													<td><%=rs("externalsource")%></td>
												</tr>
												<%end if%>
												<%if rs("agedate")  <> "" then%>
												<tr>
													<td style="background-color: #e1e1e1">AgeDate:&nbsp; </td>
													<td><%=rs("agedate")%></td>
												</tr>
												<%end if%>
												<%if rs("Category")  <> "" then%>
												<tr>
													<td style="background-color: #e1e1e1">Categoria:&nbsp; </td>
													<td><%=rs("Category")%></td>
												</tr>
												<%end if%>
												<%if rs("dob")  <> "" then%>
												<tr>
													<td style="background-color: #e1e1e1">DOB:&nbsp; </td>
													<td><%=rs("dob")%></td>
												</tr>
												<%end if%>
												<%if rs("ssn")  <> "" then%>
												<tr>
													<td style="background-color: #e1e1e1">SSN:&nbsp; </td>
													<td><%=rs("ssn")%></td>
												</tr>
												<%end if%>
												<%if rs("linkedto")  <> "" then%>
												<tr>
													<td style="background-color: #e1e1e1">Linked To:&nbsp; </td>
													<td><%=rs("linkedto")%></td>
												</tr>
												<%end if%>
												<%if rs("entered")  <> "" then%>
												<tr>
													<td style="background-color: #e1e1e1">Entered:&nbsp; </td>
													<td><%=rs("entered")%></td>
												</tr>
												<%end if%>												
												<%if rs("SubCategory")  <> "" then%>
												<tr>
													<td style="background-color: #e1e1e1">Sub Categoria:&nbsp;</td> 
													<td><%=rs("SubCategory")%></td>
												</tr>
												<%end if%>
												<%if rs("placeofbirth")  <> "" then%>
												<tr>
													<td style="background-color: #e1e1e1">Place Of Birth:&nbsp; </td>
													<td><%=rs("placeofbirth")%></td>
												</tr>
												<%end if%>
												<%if rs("locations")  <> "" then%>
												<tr>
													<td style="background-color: #e1e1e1">Locations:&nbsp; </td>
													<td><%=rs("locations")%></td>
												</tr>
												<%end if%>
												<%if rs("furtherinformation")  <> "" then%>
												<tr>
													<td style="background-color: #e1e1e1">Further Information:&nbsp; </td>
													<td><%=rs("furtherinformation")%></td>
												</tr>
												<%end if%>
												<%if rs("update")  <> "" then%>
												<tr>
													<td style="background-color: #e1e1e1">Update:&nbsp; </td>
													<td><%=rs("update")%></td>
												</tr>
												<%end if%>
											</table>
										</td>
					                </tr>
					            
					                <% rs.movenext %>
					            <%loop %>
					        <%end if %>
					    <%end if%>
					    
					</table>
				</td>
			</tr>
			<tr>
				<td>&nbsp;</td>
			</tr>						
		</table>		   
    </div>
    
    
    <div id="divProcesando" class="ContenedorModal">
        <div class="Modal">
        </div>
        <div class="InteriorModal300" style="text-align: center;">
           <img src="../Images/Procesando.gif" />
        </div>
    </div>
    
    </form>

</body>
</html>
<%
    set rs = nothing
%>
