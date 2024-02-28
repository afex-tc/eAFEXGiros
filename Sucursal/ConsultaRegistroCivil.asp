<%@ Language=VBScript %>
<%
	Dim sRespuestaRegistro, sNumeroID, sTipo, sNumeroSerie
	
	sRespuestaRegistro = request("RespuestaRegistro")
	
	if sRespuestaRegistro = "" then
		sTipo = request("Tipo")
		sNumeroID = request("NumeroID")
		sNumeroSerie = request("NumeroSerie")
	
		Response.Redirect  "http://peumo/registrocivil/default.aspx?TipoDocumento=" & sTipo & "&NumeroDocumento=" & sNumeroID & "&NumeroSerie=" & trim(sNumeroSerie)	
	end if
%>

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>

<script language="vbscript">
<!--
	sub window_onload()	
		if "<%=sRespuestaRegistro%>" <> "" then
			window.returnvalue = "<%=sRespuestaRegistro%>"			
		end if
		window.close() 
	end sub
-->
</script>

<BODY>
<P>&nbsp;</P>

</BODY>
</HTML>
