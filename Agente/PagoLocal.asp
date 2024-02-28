<%@ Language=VBScript %>
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

Dim cTC
Dim sSQL
Dim rsTC
Dim cReales

		sSql = " SELECT  top 1 * " & _
			   " FROM    Tipo_Cambio " & _
			   " WHERE   fecha_termino IS Null " & _
			   "  AND     sw_tipo = 0" & _
			   "  AND     codigo_moneda = 'ARP'"
			  

			Set rsTC =EjecutarSQLCliente(Session("afxCnxAFEXpress"),sSQl)
			
				if Err.Number <> 0 Then
					Set rsTC = Nothing
					Response.Redirect "../Compartido/Error.asp?description=" & err.Description
				End If
			
				If rsTC.EOF Then
					cTC = 0
				Else
					cTC = ccur(rsTC("valor"))				
										
					'cReales=ccur(Request("Dolares"))*(ccur(cTC))
					cReales = ccur(cTC) * ccur(Request("Dolares")) 'cdbl(Request("Dolares")) * cdbl(cTC)	
					
				End If
				'SESSION("MontoLocal")=cReales
				
				set rsTc=nothing
							
	
	%>

<script LANGUAGE="VBScript">
<!--
		window.returnvalue = "<%=cReales%>"
		window.close		
//-->
</script>



