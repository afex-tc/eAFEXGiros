<%@ Language=VBScript %>
<!--#INCLUDE virtual="/compartido/Rutinas.asp" -->
<!--#INCLUDE virtual="/compartido/Errores.asp" -->
<%
	dim bVigencia
	dim sSQL
	dim rs

	bVigencia = False

	sSQL = " execute MostrarParidadesMantenedor " & EvaluarStr(request("Moneda"))
	
	set rs = EjecutarSQLCliente(Session("afxCnxAfexchange"), sSQL)
	if err.number <> 0 then
		set rs = nothing
		mostrarerrorms "Validar vigencia paridad"
	end if		
	If Not rs.EOF Then
		bVigencia = True 
	End If

	set rs = nothing
%>

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>

<script language="vbscript">
<!--
	window.dialogWidth = 0
	window.dialogHeight = 0
	window.dialogLeft = 0
	window.dialogTop = 0
	window.defaultstatus = ""
	
	window.returnvalue = "<%=bVigencia%>" 
	window.close 
-->
</script>

<BODY>

<P>&nbsp;</P>

</BODY>
</HTML>
