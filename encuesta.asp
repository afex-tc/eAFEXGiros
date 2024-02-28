<% conexion   = true ' Habilita la Conexion %>
<!--#include virtual="asptop.inc"-->
<!--#INCLUDE virtual="/Compartido/Rutinas.asp" -->
<%
  sqla2 = " select 'Menos de 1 mes' valor,'Menos de 1 mes' nombre " & _ 
          " union all select 'Entre 1 y 6 meses','Entre 1 y 6 meses' " & _
 		  " union all select 'Entre 6 meses y 1 año','Entre 6 meses y 1 año' " & _
		  " union all select 'Entre 1 año y 3 años','Entre 1 año y 3 años' " & _
          " union all select 'Más de 3 años','Más de 3 años'"   
  
  sqla3 = " select 'Envío de giros' valor,'Envío de giros' nombre " & _ 
          " union all select 'Recibo de giros','Recibo de giros' " & _
          " union all select 'Transferencias','Transferencias' " & _
		  " union all select 'Compra o venta de moneda extranjera','Compra o venta de moneda extranjera' " & _
          " union all select 'Otros','Otros'" 		  

sqla4 = " select '1 o más veces por semana' valor,'1 o más veces por semana' nombre " & _ 
          " union all select '2 o 3 veces al mes','2 o 3 veces al mes' " & _
          " union all select '1 vez al mes','1 vez al mes' " & _
		  " union all select '1 vez cada 2 meses','1 vez cada 2 meses' " & _
          " union all select 'Entre 3 y 4 veces al año','Entre 3 y 4 veces al año'" 		  


sqla5 = " select 'Por recomendación' valor,'Por recomendación' nombre " & _ 
          " union all select 'Caminando por la calle lo vio','Caminando por la calle lo vio' " & _
          " union all select 'Página web','Página web' " & _
		  " union all select 'Eventos promocionales','Eventos promocionales'" 

sqla6 = " select 'Completamente satisfecho' valor,'Completamente satisfecho' nombre " & _ 
          " union all select 'Satisfecho','Satisfecho' " & _
          " union all select 'Poco satisfecho','Poco satisfecho' " & _
          " union all select 'Insatisfecho','Insatisfecho' " & _
		  " union all select 'Completamente insatisfecho','Completamente insatisfecho'" 

sqla7 = " select 'Excelente' valor,'Excelente' nombre " & _ 
          " union all select 'Muy buena','Muy buena' " & _
          " union all select 'Buena','Buena' " & _
          " union all select 'Regular','Regular' " & _
          " union all select 'Mala','Mala' " & _
		  " union all select 'Muy Mala','Muy Mala'" 

sqla9 = " select 'Sí' valor,'Sí' nombre " & _ 
		  " union all select 'No','No'" 

sqlb1 = " select 'Muy probable' valor,'Muy probable' nombre " & _ 
          " union all select 'Probable','Probable' " & _
          " union all select 'Improbable','Improbable' " & _
		  " union all select 'Muy improbable','Muy improbable'" 

sqlr1 = " select '1' valor,'1' nombre " & _ 
          " union all select '2','2' " & _
          " union all select '3','3' " & _
		  " union all select '4','4'" 

sqlr2 = " select '1' valor,'1' nombre " & _ 
          " union all select '2','2' " & _
          " union all select '3','3' " & _
		  " union all select '4','4'" 

sqlr3 = " select '1' valor,'1' nombre " & _ 
          " union all select '2','2' " & _
          " union all select '3','3' " & _
		  " union all select '4','4'" 

sqlr4 = " select '1' valor,'1' nombre " & _ 
          " union all select '2','2' " & _
          " union all select '3','3' " & _
		  " union all select '4','4'" 


%> 
<html>

<script src="Scripts/AC_RunActiveContent.js" type="text/javascript"></script>
<head>


<title>.:. Encuesta - AFEX .:.</title>
<link href="CSS/Linksnuevos.css" rel="stylesheet" type="text/css">
<link href="CSS/Links_T1T2T3.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
.Estilo3 {color: #FFFFFF}
-->
</style>
<link href="CSS/linkcss_3.css" rel="stylesheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1"><style type="text/css">
<!--
a:link {
	text-decoration: none;
}
a:visited {
	text-decoration: none;
}
a:hover {
	text-decoration: underline;
}
a:active {
	text-decoration: none;
}
-->
</style>
<link href="CSS/linkcss_2.css" rel="stylesheet" type="text/css">
</head>
<Script language="VBScript">
Sub window_onLoad()
  Set WshShell = CreateObject("WScript.Shell")
  dim valor
  valor = WshShell.RegRead("HKCU\Software\Microsoft\Internet Account Manager\Accounts\00000001\SMTP Email Address")
  frm.yo.value = valor
End Sub

'Jonathan Miranda G. 29-12-2006
	sub cmdEnviar_onClick()
		dim sIP
		dim sReferencia
		if valida() then			

			if frm.nombre.value = "davila" then
				sIP = "<%=request.ServerVariables("REMOTE_ADDR")%>"
				sReferencia = window.showmodaldialog("http://jfmiranda/workflow/default.aspx?nombre=" & frm.nombre.value & "&apellido=" & frm.apellido.value & "&pais=" & frm.pais.value & "&ciudad=" & frm.ciudad.value & "&email=" & frm.email.value & "&telefono=" & frm.telefono.value & "&telefonocelular=" & frm.telefonocelular.value & "&empresa=" & frm.empresa.value & "&oficina=" & frm.oficina.value &  "&pregunta2=" & frm.pregunta2.value & "&pregunta3=" & frm.pregunta3.value & "&pregunta4=" & frm.pregunta4.value & "&pregunta5=" & frm.pregunta5.value & "&pregunta6=" & frm.pregunta6.value & "&pregunta7=" & frm.pregunta7.value & "&rapidez=" & frm.rapidez.value & "&confiabilidad=" & frm.confiabilidad.value & "&preciocalidad=" & frm.preciocalidad.value & "&atencion=" & frm.atencion.value & "&pregunta9=" & frm.pregunta9.value & "&pregunta10=" & frm.pregunta10.value & "&comentario=" & frm.comentario.value & "&yo=" & frm.yo.value & "&ip=" & sIP)
				
				frm.nombre.value = empty
				frm.apellido.value  = empty
				frm.pais.value = empty
				frm.ciudad.value  = empty
				frm.email.value = empty
				frm.telefono.value = empty
				frm.telefonocelular.value  = empty
				frm.empresa.value = empty	
				frm.oficina.value = empty			
				frm.pregunta2.value = empty
				frm.pregunta3.value = empty
				frm.pregunta4.value = empty
				frm.pregunta5.value = empty
				frm.pregunta6.value = empty
				frm.pregunta7.value = empty
				frm.rapidez.value = empty
				frm.confiabilidad.value = empty
				frm.preciocalidad.value = empty
				frm.atencion.value = empty
				frm.pregunta9.value = empty
				frm.pregunta10.value = empty
				frm.comentario.value = empty				
				
				if srEferencia <> empty then
					msgbox "Su solicitud ya fué enviada con la REFERENCIA " & sReferencia & ". Nos pondremos prontamente en contacto con Ud. "

				else
					msgbox "Problemas al enviar el requerimiento."
				end if
			else
				frm.submit()
			end if
		end if
	end sub
'--------------------------- Fin -----------------------------------


</Script>
<body leftmargin="2" topmargin="2" rightmargin="0" bottommargin="0" marginwidth="0" marginheight="0">
<form method="post" name="frm" action="enviadoencuesta.asp" o>
  <div align="left">
    <table width="743" border="0" cellspacing="0" class="Borde_tabla_abajo">
      <tr>
        <td width="206" height="364" valign="top"><table width="200" border="0" cellspacing="1" bgcolor="#5A6D6B" class="Borde_tabla_abajo">
          <tr>
            <td colspan="2"><img src="Img/encuesta_01.jpg" width="200" height="151" /></td>
          </tr>
          <tr>
            <td colspan="2">&nbsp;</td>
          </tr>
          <tr>
            <td width="14" valign="top"><p align="justify" class="Estilo3"><img src="Img/boton_link_T1T2T3.jpg" width="14" height="14" /><br>
              </p>              </td>
            <td width="183"><p align="justify" class="Estilo3">Su opini&oacute;n es un valioso aporte, que tiene por objetivo mejorar el servicio que continuamente le brindamos. </p>
              <p align="justify" class="Estilo3">Por favor ingrese la siguiente informaci&oacute;n.</p></td>
          </tr>
          <tr>
            <td height="15" colspan="2">&nbsp;</td>
          </tr>
          <tr>
            <td height="15" colspan="2">
              <noscript>
              </noscript>              <script type="text/javascript">
AC_FL_RunContent( 'codebase','http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=9,0,28,0','width','200','height','250','src','Img/flashencuesta','quality','high','pluginspage','http://www.adobe.com/shockwave/download/download.cgi?P1_Prod_Version=ShockwaveFlash','movie','Img/flashencuesta' ); //end AC code
</script><noscript><object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=9,0,28,0" width="200" height="250">
                <param name="movie" value="Img/flashencuesta.swf">
                <param name="quality" value="high">
                <embed src="Img/flashencuesta.swf" quality="high" pluginspage="http://www.adobe.com/shockwave/download/download.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" width="200" height="250"></embed>
              </object></noscript></td>
          </tr>
          
          
        </table></td>
        <td width="519" valign="top"><table width="531" border="0" cellspacing="0">
          <tr>
            <td align="left" bgcolor="#5A6D6B"><div align="left"><img src="Img/encuesta.jpg" width="208" height="16"></div></td>
          </tr>
          <tr>
            <td valign="middle"></td>
          </tr>
          
          <tr>
            <td><table width="492" border="0" class="Borde_tabla_abajo">
              <tr>
                <td colspan="8" align="right"><table width="100%" border="0" cellpadding="0" cellspacing="0">
                  <tr>
                    <td width="4%"><img src="Img/img_pagempresa/Link_01.jpg" width="10" height="8" /></td>
                    <td width="96%"><span class="Estilo2">Formulario Encuesta</span></td>
                  </tr>
                </table></td>
                </tr>
              
              <tr>
                <td align="right"><img src="Img/img_pagempresa/Link_01.jpg" width="10" height="8" /></td>
                <td width="4%" align="right"><div align="right">
                <p align="left" class="textoempresa">(*)                </div></td>
				<td width="24%" align="right"><div align="left"><span class="textoempresa">Nombres: </span></div></td>
				<td><% campo_texto "nombre","creafila=no|descripcion=Nombre|obligatorio=SI|largo=30" %></td>
				<td colspan="4">&nbsp;</td>
              </tr>
              <tr>
                <td width="4%" align="right" class="textoempresa"><img src="Img/img_pagempresa/Link_01.jpg" width="10" height="8" /></td>
                <td align="right" class="textoempresa"><div align="left">(*) </div></td>
                <td align="right" class="textoempresa"><div align="left">Apellidos:</div></td>
                <td width="23%">
				<% campo_texto "apellido","creafila=no|descripcion=Apellido|obligatorio=SI|largo=30" %>				</td>
                <td colspan="4">&nbsp;</td>
              </tr>
              
              <tr>
                <td width="4%" align="right"><img src="Img/img_pagempresa/Link_01.jpg" width="10" height="8" /></td>
                <td align="right" class="textoempresa"><div align="left">(*) </div></td>
                <td align="right" class="textoempresa"><div align="left">Pa&iacute;s:</div></td>
                <td width="23%">
				 <span class="textoempresa">
				 <% combo_sql "pais","select nombre codigo,nombre from pais order by nombre","sincronizador=si|creafila=no|descripcion=Paises|defecto=" & request("pais") %>
				 </span>				 </font></td>
                <td width="1%"><img src="Img/img_pagempresa/Link_01.jpg" width="10" height="8" /></td>
                <td width="2%">(*)</td>
                <td width="8%" class="textoempresa">Ciudad:</td>
                <td width="34%"><font color="#FFFFFF" face="Arial">
                  <% campo_texto "ciudad","creafila=no|descripcion=Ciudad|obligatorio=SI|largo=30" %>
                </font></td>
              </tr>
              <tr>
                <td align="right" class="textoempresa"><img src="Img/img_pagempresa/Link_01.jpg" width="10" height="8" /></td>
                <td align="right" class="textoempresa"><div align="left">(*) </div></td>
                <td align="right" class="textoempresa"><div align="left">E-mail:</div></td>
                <td><% campo_mail "email","creafila=no|descripcion=Mail|largo=30" %></td>
                <td colspan="4">&nbsp;</td>
              </tr>
              <tr>
                <td width="4%" align="right" class="textoempresa"><img src="Img/img_pagempresa/Link_01.jpg" width="10" height="8" /></td>
                <td align="right" class="textoempresa"><div align="left">(*) </div></td>
                <td align="right" class="textoempresa"><div align="left">Tel&eacute;fono:</div></td>
                <td width="23%">
				<% campo_texto "telefono","creafila=no|descripcion=Telefono|obligatorio=SI|largo=20" %></td>
                <td><img src="Img/img_pagempresa/Link_01.jpg" width="10" height="8" /></td>
                <td colspan="2" class="textoempresa">Celular:</td>
                <td><% campo_texto "telefonocelular","creafila=no|descripcion=Telefono Celular|obligatorio=NO|largo=20" %></td>
              </tr>
              
              <tr>
                <td align="right" class="textoempresa"><img src="Img/img_pagempresa/Link_01.jpg" width="10" height="8" /></td>
                <td colspan="2" align="right" class="textoempresa"> <div align="left">Empresa:</div></td>
                <td><% campo_texto "empresa","creafila=no|descripcion=empresa|obligatorio=NO|largo=30" %></td>
                <td colspan="4">&nbsp;</td>
              </tr>
              <tr>
                <td align="right">&nbsp;</td>
                <td colspan="2" align="right">&nbsp;</td>
                <td>&nbsp;</td>
                <td colspan="4">&nbsp;</td>
              </tr>
              <tr>
                <td colspan="8" align="right"><div align="center"><img src="Img/lineahorizontal.jpg" width="490" height="1"></div></td>
                </tr>
              <tr><td align="right">&nbsp;</td><td colspan="2" align="right">&nbsp;</td><td>&nbsp;</td><td colspan="4">&nbsp;</td></tr>
              
              <tr>
                <td height="21" align="right"><span class="textoempresa"><img src="Img/img_pagempresa/Link_01.jpg" width="10" height="8" /></span></td>
                <td colspan="7" align="right"><div align="left">
                  <table width="100%" border="0">
                    <tr>
                      <td width="4%" class="Estilo2">1.-</td>
                      <td width="96%"><span class="Estilo2">&iquest;En qu&eacute; oficina se atiende Ud. usualmente?</span></td>
                    </tr>
                  </table>
                </div></td>
                </tr>
              <tr>
                <td align="right">&nbsp;</td>
                <td colspan="2" align="right">&nbsp;</td>
                <td colspan="5"><span class="textoempresa">
                  <% combo_sql "oficina","select nombre, nombre from cliente where tipo=4 order by nombre","sincronizador=si|creafila=no|descripcion=Oficina" %>
                </span></td>
              </tr>
              <tr>
                <td align="right">&nbsp;</td>
                <td colspan="2" align="right">&nbsp;</td>
                <td colspan="5">&nbsp;</td>
                </tr>
              <tr>
                <td align="right"><span class="textoempresa"><img src="Img/img_pagempresa/Link_01.jpg" width="10" height="8" /></span></td>
                <td colspan="7" align="right" class="Estilo2"><div align="left">
                  <table width="100%" border="0">
                    <tr>
                      <td width="2%" class="Estilo2">2.-</td>
                      <td width="98%" class="Estilo2">&iquest;Cu&aacute;nto tiempo lleva utilizando los servicios Afex? </td>
                    </tr>
                  </table>
                </div></td>
                </tr>
              <tr>
                <td align="right">&nbsp;</td>
                <td colspan="2" align="right">&nbsp;</td>
                <td colspan="5"><% combo_sql "pregunta2",sqla2,"creafila=no|descripcion=Pregunta 2|obligatorio=SI|defecto= "%></td>
                </tr>
              <tr>
                <td align="right">&nbsp;</td>
                <td colspan="7" align="right">&nbsp;</td>
              </tr>
              <tr>
                <td align="right"><span class="textoempresa"><img src="Img/img_pagempresa/Link_01.jpg" width="10" height="8" /></span></td>
                <td colspan="7" align="right"><div align="left" class="Estilo2">
                  <table width="100%" border="0">
                    <tr>
                      <td width="4%" class="Estilo2">3.-</td>
                      <td width="96%" class="Estilo2">&iquest;Qu&eacute; operaciones hace Ud. usualmente en Afex? </td>
                    </tr>
                  </table>
                </div></td>
                </tr>
              <tr>
                <td align="right">&nbsp;</td>
                <td colspan="2" align="right">&nbsp;</td>
                <td colspan="5"><% combo_sql "pregunta3",sqla3,"creafila=no|descripcion=Pregunta 3|obligatorio=SI|defecto= "%></td>
                </tr>
              <tr>
                <td align="right">&nbsp;</td>
                <td colspan="7" align="right">&nbsp;</td>
              </tr>
              <tr>
                <td align="right"><span class="textoempresa"><img src="Img/img_pagempresa/Link_01.jpg" width="10" height="8" /></span></td>
                <td colspan="7" align="right"><table width="100%" border="0">
                  <tr>
                    <td width="4%" class="Estilo2">4.-</td>
                    <td width="96%" class="Estilo2">&iquest;Con qu&eacute; frecuencia usa Ud. los servicios Afex?</td>
                  </tr>
                </table></td>
                </tr>
              <tr>
                <td align="right">&nbsp;</td>
                <td colspan="2" align="right">&nbsp;</td>
                <td colspan="5"><% combo_sql "pregunta4",sqla4,"creafila=no|descripcion=Pregunta 4|defecto= "%></td>
                </tr>
              <tr>
                <td align="right">&nbsp;</td>
                <td colspan="7" align="right">&nbsp;</td>
              </tr>
              <tr>
                <td align="right"><span class="textoempresa"><img src="Img/img_pagempresa/Link_01.jpg" width="10" height="8" /></span></td>
                <td colspan="7" align="right"><table width="100%" border="0">
                  <tr>
                    <td width="4%" class="Estilo2">5.-</td>
                    <td width="96%" class="Estilo2">&iquest;C&oacute;mo lleg&oacute; a operar a trav&eacute;s de Afex?</td>
                  </tr>
                </table>
                  <div align="left"></div></td>
                </tr>
              <tr>
                <td align="right">&nbsp;</td>
                <td colspan="2" align="right">&nbsp;</td>
                <td colspan="5"><% combo_sql "pregunta5",sqla5,"creafila=no|descripcion=Pregunta 5|obligatorio=SI|defecto= "%></td>
                </tr>
              <tr>
                <td align="right">&nbsp;</td>
                <td colspan="7" align="right">&nbsp;</td>
              </tr>
              <tr>
                <td align="right"><span class="textoempresa"><img src="Img/img_pagempresa/Link_01.jpg" width="10" height="8" /></span></td>
                <td colspan="7" align="right"><table width="100%" border="0">
                  <tr>
                    <td width="4%" class="Estilo2">6.-</td>
                    <td width="96%" class="Estilo2">&iquest;Cuan satisfecho est&aacute; Ud. con los servicios  otorgados por Afex?</td>
                  </tr>
                </table></td>
                </tr>
              
              <tr>
                <td align="right">&nbsp;</td>
                <td colspan="2" align="right">&nbsp;</td>
                <td colspan="5"><% combo_sql "pregunta6",sqla6,"creafila=no|descripcion=Pregunta 6|obligatorio=SI|defecto= "%></td>
                </tr>
              <tr>
                <td align="right">&nbsp;</td>
                <td colspan="7" align="right">&nbsp;</td>
              </tr>
              <tr>
                <td align="right"><span class="textoempresa"><img src="Img/img_pagempresa/Link_01.jpg" width="10" height="8" /></span></td>
                <td colspan="7" align="right"><table width="100%" border="0">
                  <tr>
                    <td width="4%" class="Estilo2">7.-</td>
                    <td width="96%" class="Estilo2">&iquest;C&oacute;mo evaluar&iacute;a Ud. la atenci&oacute;n por parte de los  cajeros de Afex?</td>
                  </tr>
                </table></td>
                </tr>
              <tr>
                <td align="right">&nbsp;</td>
                <td colspan="2" align="right">&nbsp;</td>
                <td colspan="5"><% combo_sql "pregunta7",sqla7,"creafila=no|descripcion=Pregunta 7|obligatorio=SI|defecto= "%></td>
              </tr>
              <tr>
                <td align="right">&nbsp;</td>
                <td colspan="2" align="right">&nbsp;</td>
                <td colspan="5">&nbsp;</td>
                </tr>
              <tr>
                <td align="right"></td>
                <td colspan="7" rowspan="2" align="right"><table width="100%" border="0">
                  <tr>
                    <td width="4%" valign="top" class="Estilo2">8.-</td>
                    <td width="96%" class="Estilo2"><div align="justify">Ordene con n&uacute;meros (1,2,3,4) las siguientes  categor&iacute;as seg&uacute;n la importancia que Ud. les otorga, asignando el 1 a la que le  da mayor importancia y el 4 al que le da menos.</div></td>
                  </tr>
                </table></td>
                </tr>
              <tr>
                <td align="right" valign="top"><span class="textoempresa"><img src="Img/img_pagempresa/Link_01.jpg" width="10" height="8" /></span></td>
              </tr>
              <tr>
                <td colspan="2" align="right">&nbsp;</td>
                </tr>
              <tr>
                <td align="right">&nbsp;</td>
                <td colspan="2" align="right"><div align="left" class="textoempresa">Rapidez:</div></td>
                <td colspan="5"><table width="81%" border="0" cellpadding="0" cellspacing="0">
                  <tr>
                    <td><% combo_sql "rapidez",sqlr1,"creafila=no|descripcion=Pregunta 8|obligatorio=SI|defecto= "%></td>
                    </tr>
                </table></td>
                </tr>
              <tr>
                <td align="right">&nbsp;</td>
                <td colspan="2" align="right"><div align="left" class="textoempresa">Confiabilidad:</div></td>
                <td colspan="5"><table width="81%" border="0" cellpadding="0" cellspacing="0">
                  <tr>
                    <td><% combo_sql "Confiabilidad",sqlr2,"creafila=no|descripcion=Pregunta 8|obligatorio=SI|defecto= "%></td>
                    </tr>
                </table></td>
                </tr>
              <tr>
                <td align="right">&nbsp;</td>
                <td colspan="2" align="right" nowrap><div align="left" class="textoempresa">Precio-Calidad:</div></td>
                <td colspan="5"><table width="81%" border="0" cellpadding="0" cellspacing="0">
                  <tr>
                    <td><% combo_sql "preciocalidad",sqlr3,"creafila=no|descripcion=Pregunta 8|obligatorio=SI|defecto= "%></td>
                    </tr>
                </table></td>
                </tr>
              <tr>
                <td align="right">&nbsp;</td>
                <td colspan="2" align="right"><div align="left" class="textoempresa">Atenci&oacute;n:</div></td>
                <td colspan="5"><table width="81%" border="0" cellpadding="0" cellspacing="0">
                  <tr>
                    <td><% combo_sql "atencion",sqlr4,"creafila=no|descripcion=Pregunta 8|obligatorio=SI|defecto= "%></td>
                    </tr>
                </table></td>
                </tr>
              <tr>
                <td colspan="2" align="right">&nbsp;</td>
                </tr>
              <tr>
                <td align="right"><span class="textoempresa"><img src="Img/img_pagempresa/Link_01.jpg" width="10" height="8" /></span></td>
                <td colspan="7" align="right"><table width="100%" border="0">
                  <tr>
                    <td width="4%" class="Estilo2">9.-</td>
                    <td width="96%" class="Estilo2">&iquest;Recomendar&iacute;a Ud. los servicios de Afex a otra  persona?</td>
                  </tr>
                </table></td>
                </tr>
              <tr>
                <td align="right">&nbsp;</td>
                <td colspan="2" align="right">&nbsp;</td>
                <td colspan="5"><% combo_sql "pregunta9",sqla9,"creafila=no|descripcion=Pregunta 9|obligatorio=SI|defecto= "%></td>
              </tr>
              <tr>
                <td align="right">&nbsp;</td>
                <td colspan="2" align="right">&nbsp;</td>
                <td colspan="5">&nbsp;</td>
                </tr>
              
              <tr>
                <td align="center" valign="top"></td>
                <td colspan="7" rowspan="2" align="right"><table width="100%" border="0">
                  <tr>
                    <td width="4%" valign="top" class="Estilo2">10.-</td>
                    <td width="96%" class="Estilo2"><div align="justify">Bas&aacute;ndose en su propia experiencia, &iquest;buscar&iacute;a Ud.  empresas similares para realizar sus operaciones?</div></td>
                  </tr>
                </table></td>
                </tr>
              <tr>
                <td align="center" valign="top"><img src="Img/img_pagempresa/Link_01.jpg" width="10" height="8" /></td>
              </tr>
              <tr>
                <td align="right">&nbsp;</td>
                <td colspan="2" align="right">&nbsp;</td>
                <td colspan="5"><% combo_sql "pregunta10",sqlb1,"creafila=no|descripcion=Pregunta 10|obligatorio=SI|defecto= "%></td>
                </tr>
              
              <tr>
                <td align="right">&nbsp;</td>
                <td colspan="7" align="right">&nbsp;</td>
              </tr>
              <tr>
                <td align="right"><img src="Img/img_pagempresa/Link_01.jpg" width="10" height="8" /></td>
                <td colspan="7" align="right"><table width="100%" border="0">
                  <tr>
                    <td width="4%" valign="top" class="Estilo2">11.-</td>
                    <td width="96%" class="Estilo2"><div align="justify">Otros comentarios o sugerencias:</div></td>
                  </tr>
                </table>
                  <div align="left"></div></td>
                </tr>
              <tr>
                <td align="right">&nbsp;</td>
                <td colspan="2" align="right">&nbsp;</td>
                <td colspan="5"><% campo_texto "comentario","creafila=no|descripcion=Comentario|obligatorio=NO|largo=30|alto=8" %></td>
                </tr>
              
              <tr>
                <td height="10" colspan="8" align="right"></td>
              </tr>
              <tr>
                <td colspan="8" align="right"><div align="center"><img src="Img/lineahorizontal.jpg" width="490" height="1"></div></td>
                </tr>
              <tr>
                <td width="4%" align="right"></td>
                <td colspan="2" align="right"></td>
                <td width="23%"></td>
                <td colspan="4"></td>
              </tr>
              <tr>
                <td width="4%" align="right"></td>
                <td colspan="2" align="right"></td>
                <td width="23%"></td>
                <td colspan="4"></td>
              </tr>
              <tr>
                <td align="center" colspan="8"><div align="center">
                  <center>
                    <table border="0"
      width="100%" cellspacing="15" cellpadding="0">
                      <tr>
                        <td width="100%"><div align="center">
                          <center>
                            <p align="center"><font face="Arial" color="#000000"><small>
                              
							  <input TYPE="button" class="textoempresa" name="cmdEnviar" VALUE="Enviar">

                              <input name="reset" TYPE="reset" class="textoempresa">
                              </small></font>
                            </center>
                        </div></td>
                      </tr>
                    </table>
                  </center>
                </div></td>
              </tr>
            </table>              </td>
          </tr>
          
        </table></td>
      </tr>
      <tr>
        <td height="90" colspan="2" valign="top"><table width="727" border="0" align="center" cellspacing="0">
          <tr>
            <td height="10" colspan="8"><img src="Img/img_paginahome/lineaHorizontal.jpg" width="730" height="1" /></td>
          </tr>
          <tr>
            <td width="2" height="48">&nbsp;</td>
            <td width="469"><img src="Img/img_paginagiros/giros2006_r40_c4.jpg" alt="" name="giros2006_r40_c4" width="293" height="59" border="0" id="giros2006_r40_c4" /></td>
            <td width="40"><div align="center"><a href="http://www.adobe.com/shockwave/download/download.cgi?P1_Prod_Version=ShockwaveFlash" target="_blank"></a></div></td>
            <td width="40"><div align="center"> <a href="Default.asp" target="_self"> <img src="Img/img_paginagiros/Home.jpg" alt=".:. Home .:." width="34" height="34" border="0" /></a> </div></td>
            <td width="40"><div align="center"> <a href="#" target="Principal"> <img src="Img/img_paginagiros/ayuda.jpg" alt=".:. Ayuda .:."  width="34" height="34" border="0"/></a> </div></td>
            <td width="40"><div align="center"> <a href="contacto.asp" target="_self"> <img src="Img/img_paginagiros/contacto.jpg" alt=".:. Contacto .:." width="35" height="33" border="0" /></a> </div></td>
            <td width="40"><div align="center"> <a href="mapasitio.htm" target="_self"><img src="Img/img_paginagiros/mapa.jpg" alt=".:. Mapa del Sitio .:." width="34" height="33" border="0" /></a> </div></td>
            <td width="40"><div align="center"> <a href="#" target="Principal"><img src="Img/img_paginagiros/usa.jpg" alt=".:. Versi&oacute;n en Ingles .:." width="34" height="33" border="0"/></a> </div></td>
          </tr>
        </table>
          <table width="300" border="0" align="right" cellspacing="0" bgcolor="#31514A">
            <tr>
              <td valign="bottom"><table width="446" border="0" align="right" cellspacing="0" bgcolor="#31514A">
                  <tr>
                    <td width="456" valign="bottom" class="texto_piepagina"><div align="center">Moneda # 1160, Piso 8 y 9 - Santiago, Chile (562) 636 9000 -<a href="contacto.asp" target="_self" class="linkcss2"> contactenos@afex.cl</a></div></td>
                  </tr>
              </table></td>
            </tr>
          </table>
        </td>
      </tr>
    </table>
  </div>
  <input type="hidden" name="yo"> 
</form>
</body>
</html>
<!--#include virtual="aspfin.inc"-->