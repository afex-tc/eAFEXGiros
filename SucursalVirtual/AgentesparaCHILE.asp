<%@ Language=VBScript %>
<!--#INCLUDE virtual="/Compartido/Rutinas.asp" -->
<%
	Dim rsDestino, sSQL
	Dim rsOrigen
	
	If Request.Form("cbxPaisLocal") <> Empty Then AgenteOrigen()
	If Request.Form("cbxCiudadDestino") <> Empty Then AgenteDestino(Request.Form("cbxCiudadDestino"))
	
	Sub AgenteOrigen()	
		On Error Resume Next		
				
		sSQL = " SELECT agente, ciudad, telefono " & _
				 " FROM direccionesagente " & _
				 " WHERE pais = " & EvaluarSTR(Request.Form("cbxPaisLocal")) & _				 
				 " ORDER BY agente "	
		Set rsOrigen = EjecutarSQLCliente(Session("afxCnxAFEXpress"), sSQL)
		If Err.number <> 0 Then
			Set rsOrigen = Nothing			
			MostrarErrorMS ""			
		End If		
	End Sub
	
	Sub AgenteDestino(ByVal Ciudad)
		Dim sSQL
		
	'	On Error Resume Next
		sSQL = " SELECT    DISTINCT ag.codigo_agente, ag.nombre_agente, ag.direccion_agente, ag.fono_agente, ag.codarea_agente, ag.codpais_agente " & _
				 " FROM      agente AG " & _
				 " JOIN      comision CM ON ag.codigo_agente = cm.codigo_agente " & _
				 " WHERE     (cm.pais = " & EvaluarStr("CL") & _
				 "  OR      cm.pais = '**') " & _
				 " AND      (cm.ciudad = " & EvaluarStr(Ciudad) & _
				 "  OR      cm.ciudad = '***') " & _
	          " AND      cm.sentido = 1 " & _
		       " AND      ag.estado_agente <> 0 " & _
			    " AND      cm.fecha_termino is null "
		
		Set rsDestino = EjecutarSQLCliente(Session("afxCnxAFEXpress"), sSQL)
		If Err.number <> 0 Then
			Set rsDestino = Nothing			
			MostrarErrorMS ""
		End If			
	End Sub	
	
	Sub PaisOrigen(Byval Pais)
		Dim rs, sSQL
		
		On Error Resume Next		
				
		sSQL = " SELECT descripcion_pais as nombre, codigo_pais as codigo" & _
				 " FROM pais " & _
				 " WHERE exists (select pais from direccionesagente where pais = codigo_pais) " & _				 
				 " ORDER BY descripcion_pais "	
		Set rs = EjecutarSQLCliente(Session("afxCnxAFEXpress"), sSQL)
		If Err.number <> 0 Then
			Set rs = Nothing			
			MostrarErrorMS ""
			exit sub
		End If		
	
		Response.Write "<option value=></option>"
		If  Not rs.EOF Then
			Do Until rs.eof	
				If UCASE(Trim(rs("codigo"))) = Ucase(Trim(Pais)) Then
					sSelect = "SELECTED"				
				Else
					sSelect = ""
				End If
				sMayMin = trim(rs("nombre"))
				sMayMin = MayMin(sMayMin) 
				Response.write "<option " & sSelect & " value=" & _
				trim(rs("codigo")) & ">" & _
				sMayMin & _
				" </option> "
				If err.number <> 0 Then 
					'response.Redirect "../Compartido/Error.asp?Titulo=Error en HágaseCliente&Number=" & Err.Number  & "&Source=" & Err.source  & "&Description=" & Err.description	
					MostrarErrorMS "País Origen 3"
				End If																
				rs.MoveNext
			Loop
		End If
			
		Set rs = Nothing
	End Sub
	
%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link href="CSS/linkcss_3.css" rel="stylesheet" type="text/css">
<link href="CSS/linkcss_2.css" rel="stylesheet" type="text/css">
<script type="text/JavaScript">
<!--
function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}
//-->
</script>
<title>.:: Enviar a CHILE ::.</title>
<link href="../CSS/CSS_SucursalVirtual.css" rel="stylesheet" type="text/css">
<link href="../CSS/Links_T1T2T3.css" rel="stylesheet" type="text/css"></HEAD>

<script language="VBScript">
<!--
	Sub cbxPaisDestino_onChange()
		frmEnviadores.action = "AgentesparaCHILE.asp" 
		frmEnviadores.submit()
		frmEnviadores.action = "" 
	End Sub
	
	Sub cbxCiudadDestino_onChange()
		frmEnviadores.action = "AgentesparaCHILE.asp" 
		frmEnviadores.submit()
		frmEnviadores.action = "" 
	End Sub
	
	Sub cbxPaisLocal_onChange()
		frmEnviadores.action = "AgentesparaCHILE.asp" 
		frmEnviadores.submit()
		frmEnviadores.action = "" 
	End Sub
	
	Sub window_onLoad()
		If frmEnviadores.cbxCiudadDestino.value <> empty then
			spnDestino.innerText = frmEnviadores.cbxCiudadDestino.options(frmEnviadores.cbxCiudadDestino.selectedIndex).text
		end if
		If frmEnviadores.cbxPaisLocal.value <> empty then
			spnOrigen.innerText = frmEnviadores.cbxPaisLocal.options(frmEnviadores.cbxPaisLocal.selectedIndex).text
		end if
	End Sub
-->
</script>

<BODY leftmargin="2" topmargin="2" rightmargin="0" bottommargin="0" marginwidth="0" marginheight="0" onLoad="MM_preloadImages('Img/botonconsultar_f2.jpg')">

<form method="post" name="frmEnviadores" action="">
	<table width="530" border="0" cellpadding="0" cellspacing="0" class="Borde_tabla_abajo">
		<tr>
		  <td colspan="4"><table width="100%" border="0">
            <tr>
              <td><img src="../Img/Cobertura_Internacional .jpg" width="530" height="35"></td>
            </tr>
          </table></td>
	  </tr>
		<tr>
			<td width="2%"><img src="../Img/img_pagempresa/Link_01.jpg" alt="punto" width="10" height="9" /><br></td>      
			<td width="48%"><strong>&iquest;Donde se ubica Ud.?</strong></td>
			<td width="2%"><img src="../Img/img_pagempresa/Link_01.jpg" alt="punto" width="10" height="9" /><br>
				<select name="cbxPaisDestino" class="linkcss3" style="display: none">
				<% CargarUbicacion 1, "", Request.Form("cbxPaisDestino") %>
				</select></td>
            <td width="48%"><span class="Estilo2">&iquest;Donde quiere enviar?</span></td>
	  </tr>
		<tr>
		  <td colspan="2"><select name="cbxCiudadDestino" class="Borde_tabla_abajo">
              <% CargarCiudadesPais "CL", Request.Form("cbxCiudadDestino") %>
            </select></td>
	      <td colspan="2"><select name="cbxPaisLocal" class="Borde_tabla_abajo">
            <% PaisOrigen Request.Form("cbxPaisLocal") %>
          </select></td>
	  </tr>
		<tr>
		  <td colspan="2">&nbsp;</td>
		  <td colspan="2">&nbsp;</td>
	  </tr>
	</table>
	
<%If Request.Form("cbxCiudadDestino") <> Empty Then%>
	<div class="Borde_tabla_abajo" style="OVERFLOW: auto; WIDTH: 538px; HEIGHT: 200px; align: left">
		<table width="516" border="0" cellpadding="1" cellspacing="1">
<tr>
  <td width="10"><img src="../Img/img_pagempresa/Link_01.jpg" alt="punto" width="10" height="9" /></td>
				<td width="97"><span class="Estilo2">Agentes en :</span>&nbsp;<span id="spnDestino"></td>
		        <td width="5"></td>
        </tr>
			<%							
				Do While Not rsDestino.EOF
			%>
					<tr>
					  <td>&nbsp;</td>
					  <td colspan="6" class="textoempresa"></td>
		  </tr>
					<tr>
					  <td><img src="../Img/img_pagempresa/Link_01.jpg" alt="punto" width="10" height="9" /></td>
						<td colspan="2" class="textoempresa"><strong>Agente o Local Afex</strong></td>
						<td><img src="../Img/img_pagempresa/Link_01.jpg" alt="punto" width="10" height="9" /></td>
						<td class="textoempresa"><strong>Tel&eacute;fono</strong></td>
						<td width="10"><img src="../Img/img_pagempresa/Link_01.jpg" alt="punto" width="10" height="9" /></td>
					    <td width="258" class="textoempresa"><strong>Direcci&oacute;n</strong></td>
					</tr>
					<tr>
					  <td>&nbsp;</td>
					  <td colspan="2"><div align="left"><span class="textoempresa"><%=rsDestino("nombre_agente")%></span></div></td>
					  <td width="10">&nbsp;</td>
					  <td width="104"><div align="left"><span class="textoempresa"><%=trim(rsDestino("codpais_agente")) & " - " & trim(rsDestino("codarea_agente")) & " - " & rsDestino("fono_agente")%></span></div></td>
					  <td colspan="2"><div align="left"><span class="textoempresa"><%=rsDestino("direccion_agente")%></span></div></td>
		  </tr>
					
					
							
			<%
					rsDestino.MoveNext
				Loop						
			%>				
		</table>		
  </div>
	<%End If%>
	<br> 
	<%If Request.Form("cbxPaisLocal") <> Empty Then%>
	<div class="Borde_tabla_abajo" style="OVERFLOW: auto; WIDTH: 538px; HEIGHT: 200px; align: left">
		<table width="516" border="0" cellpadding="1" cellspacing="1" >					
<tr>
  <td width="10"><img src="../Img/img_pagempresa/Link_01.jpg" alt="punto" width="10" height="9" /></td>
				<td width="110"><span class="Estilo2">Agentes en :&nbsp;</span><span id="spnOrigen"></td>
		  </tr>
			<%							
				Do While Not rsOrigen.EOF
			%>
					<tr>
					  <td>&nbsp;</td>
					  <td>&nbsp;</td>
					  <td colspan="2">&nbsp;</td>
					  <td colspan="2">&nbsp;</td>
		  </tr>
					<tr>
					  <td><img src="../Img/img_pagempresa/Link_01.jpg" alt="punto" width="10" height="9" /></td>
					  <td class="textoempresa"><strong>Agente</strong></td>
					  <td width="10"><img src="../Img/img_pagempresa/Link_01.jpg" alt="punto" width="10" height="9" /></td>
					  <td width="104" class="textoempresa"><strong>Tel&eacute;fono</strong></td>
					  <td width="12"><img src="../Img/img_pagempresa/Link_01.jpg" alt="punto" width="10" height="9" /></td>
		              <td width="251"><strong class="textoempresa">Direcci&oacute;n</strong></td>
		  </tr>
					<tr>
					  <td>&nbsp;</td>
						<td><span class="textoempresa"><%=rsOrigen("Agente")%></span></td>
					    <td colspan="2"><span class="textoempresa"><%=rsOrigen("Telefono")%></span></td>
						<td colspan="2"><span class="textoempresa"><%=rsOrigen("Ciudad")%></span></td>
					</tr>
					<tr>
					  <td>&nbsp;</td>
					  <td>&nbsp;</td>
					  <td colspan="2">&nbsp;</td>
					  <td colspan="2">&nbsp;</td>
		  </tr>		
			<%
					rsOrigen.MoveNext
				Loop						
			%>	
	  </table>		
  </div>
	<%End If%>
	
	<br><br>
	
	
</form>

</BODY>
</HTML>

<%
	Set rsOrigen = Nothing
	Set rsDestino = Nothing
%>
