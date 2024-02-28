<%@ Language=VBScript %>
<!--#INCLUDE virtual="/Compartido/Rutinas.asp" -->
<%
	
	'Objetivo:	Cargar en un combo los agentes
	Function CargarAgente()
		Dim rs, sSQL
		
		On Error Resume Next
		sSQL = "SELECT codigo_agente, nombre_agente FROM agente WHERE sw_consultar = 1 ORDER BY nombre_agente "	
		Set rs = EjecutarSQLCliente(Session("afxCnxAFEXpress"), sSQL)
		
		'response.Write ssql
		'response.End 
		
	
		If Err.number <> 0 Then
			Set rs = Nothing			
			MostrarErrorMS ""
		End If		

		Response.Write "<option value=></option>"
		If  Not rs.EOF Then
			Do Until rs.eof	
				sMayMin = trim(rs("nombre_agente"))
				sMayMin = MayMin(sMayMin) 
				Response.Write "<option " & sSelect & " value=" & _
				trim(rs("codigo_agente")) & ">" & _
				sMayMin & _
				" </option> "
				If Err.number <> 0 Then
					Set rs = Nothing				
					MostrarErrorMS ""
				End If
				sm = rs("nombre_agente") & "//" & rs("codigo_agente")		
				rs.MoveNext
				sm = sm & rs("nombre_agente") & "//" & rs("codigo_agente")	
			Loop
		End If		
		
		Set rs = Nothing		
	End Function
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
<title>.:: Consultar Giro ::.</title>
<link href="CSS/Links_T1T2T3.css" rel="stylesheet" type="text/css">
<link href="CSS/Links_T1T2T3_B.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
.Estilo3 {color: #FFFFFF}
.Estilo4 {color: #333333}
-->
</style>
</HEAD>

<script language="VBScript">
<!--
	Sub Identificador_onClick()
		if frmMail.cbxAgentes.value = empty then exit Sub
		window.showmodaldialog "img/INVOICE_" & frmMail.cbxAgentes.value & ".jpg"
	End Sub

	Sub cmdConsultar_onClick()
		Dim sMensaje
		
		'verifica los datos
		If frmMail.txtCodigo.value = Empty Or frmMail.txtCorreo.value = Empty Or frmMail.cbxAgentes.value = Empty Then
			msgbox "Debe ingresar todos los datos que se solicitan. ",,"Consultar Giro"
			exit sub
		End If
		
		sMensaje = window.showModalDialog ("EnviarMailHistoria.asp?Giro=" & frmMail.txtCodigo.value & "&Mail=" & frmMail.txtCorreo.value & _
														"&Agente=" & frmMail.cbxAgentes.value)
		
		msgbox sMensaje,,"Consultar Giro"
	End Sub

-->
</script>

<BODY leftmargin="2" topmargin="2" rightmargin="0" bottommargin="0" marginwidth="0" marginheight="0" onLoad="MM_preloadImages('Img/botonconsultar_f2.jpg')">

<form method="post" name="frmMail" action="">
  <table width="540" height="296" border="0" cellspacing="0" class="Borde_tabla_abajo">
    <tr>
      <td width="561" height="135"><table width="530" border="0" align="left" cellpadding="0" cellspacing="0">
        
        <tr>
          <td ><table width="527" border="0" cellpadding="0" cellspacing="0">
            
            <tr>
              <td height="59" colspan="3"><div align="left"><img src="Img/verifique.jpg" width="531" height="62"></div></td>
            </tr>
            <tr>
              <td width="209" height="50" bgcolor="#003031"><img src="Img/1paso.jpg" width="203" height="47"></td>
              <td colspan="2" bgcolor="#003031"><img src="Img/2paso.jpg" width="201" height="48"></td>
              </tr>
            <tr>
					<td height="40" bgcolor="#003031">
						<select name="cbxAgentes" class="linkcss3">
							<% CargarAgente() %>
						</select>					</td>
              <td colspan="2" bgcolor="#003031"><input name="txtCodigo" type="Text" class="linkcss3"></td>
              </tr>
            <tr>
              <td height="22" bgcolor="#003031">&nbsp;</td>
              <td colspan="2" bgcolor="#003031"><table width="143" border="0" cellspacing="1">
                <tr>
                  <td><img style="cursor:hand" name="Identificador" src="Img/click4.jpg" width="107" height="16">&nbsp; </td>
                  </tr>
              </table></td>
              </tr>
            <tr>
              <td height="8" colspan="3" bgcolor="#003031"></td>
              </tr>
            <tr>
              <td bgcolor="#003031">&nbsp;</td>
              <td width="138" bgcolor="#003031"><img src="Img/3paso.jpg" width="134" height="48"></td>
              <td width="168" rowspan="2" valign="bottom" bgcolor="#003031"><img style="cursor:hand" name="cmdConsultar" src="Img/botonconsultar.jpg" width="90" height="35" border="0" alt=""></td>
              </tr>
            <tr>
              <td bgcolor="#003031">&nbsp;</td>
              <td bgcolor="#003031"><input name="txtCorreo" type="Text" class="linkcss3"></td>
              </tr>
            <tr>
              <td height="8" colspan="3" bgcolor="#003031"></td>
            </tr>
            <tr>
              <td height="8" colspan="3" bgcolor="#003031"><div align="left"> &nbsp;</div></td>
            </tr>
            <tr>
              <td height="8" colspan="3" bgcolor="#003031"></td>
            </tr>
          </table>
           <img src="Img/abajo_envio.jpg" width="532" height="30"></td>
        </tr>
        
      </table></td>
    </tr>
  </table>
    <table width="540" border="0" align="left" cellspacing="0" background="../Img/pie_01.jpg" bgcolor="#F6f6f6" class="Borde_tabla_abajo">
    <tr>
      <td colspan="2"><div align="left"></div></td>
    </tr>
    <tr>
      <td width="696"><div align="center"><span><a href="Default.asp"  class="linkcss3" target="_blank"><span class="Estilo3 Estilo4">&#8226;</span> Home <span class="Estilo4">&#8226;</span></a></span></div></td>
      <td width="43" rowspan="2"><div align="right"><img src="Img/logitoafex.jpg" alt="AFEX" width="40" height="40" border="0" class="bordelogo" /></div></td>
    </tr>
    <tr>
      <td height="15"><div align="center"></div></td>
    </tr>
  </table>
  <p>&nbsp;</p>
</form>

</BODY>
</HTML>
