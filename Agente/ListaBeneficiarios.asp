<%@ Language=VBScript %>
<!--#INCLUDE virtual="/Compartido/Rutinas.asp" -->
<!--#INCLUDE virtual="/Compartido/Errores.asp" -->
<%
	dim rs, sSQL, i
	
	sSQL = " select b.*, p.descripcion_pais as nombrepais, c.descripcion_ciudad as nombreciudad " & _
			" from beneficiarios b " & _
				" inner join pais p on p.codigo_pais = b.pais " & _
				" inner join ciudad c on c.codigo_ciudad = b.ciudad " & _
			" where b.codigocliente = " & evaluarstr(trim(request("Cliente")))
	set rs = ejecutarsqlcliente(Session("afxCnxAFEXpress"), sSQL)
	if err.number <> 0 then
		mostrarerrorms "Error al buscar Beneficiarios."
	end if
	
%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
	<script language="vbscript">
	<!--
		window.dialogleft = 230
		window.dialogtop = 250
		window.dialogheight = 20
		window.dialogwidth = 30
	
		sub Copiar(Fila)
			dim sBeneficiario
			
			sBeneficiario = tblB.rows(fila).cells(0).innerText & ";" & tblB.rows(fila).cells(1).innerText & _
					";" & tblB.rows(fila).cells(2).innerText & ";" & tblB.rows(fila).cells(4).innerText			
			window.returnvalue = sBeneficiario
			
			window.close 
		end sub
	
	-->
	</script>
	
	<style type="text/css">	
	<!--
		.tabla 
			{
				border-top: solid 2px #999999;
				border-right: solid 1px #999999;
				border-left: solid 1px #999999;
				border-bottom: solid 1px #999999;
			}
		
		.fila
			{				
				border-bottom: solid 1px #999999;
			}		
	-->
	</style>

<BODY>
	
	<br><br>
	<table id="tblB" align="center" width="80%" class="tabla">
		<tr>
			<td style="font-family: verdana; font-size: 12px" class="fila"><b>Nombres</b></td>
			<td style="font-family: verdana; font-size: 12px" class="fila"><b>Apellidos</b></td>
			<td style="font-family: verdana; font-size: 12px" class="fila"><b>País</b></td>
			<td style="font-family: verdana; font-size: 12px" class="fila"><b>Ciudad</b></td>
		</tr>
		<tr>
			<td style="font-family: verdana; font-size: 12xp">&nbsp;</td>
		</tr>
		<%i = 2%>
		<%do while not rs.eof %>		
			<tr style="cursor: hand" onClick="Copiar <%=i%>">
				<td style="font-family: verdana; font-size: 10px"><%=rs("nombres")%></td>
				<td style="font-family: verdana; font-size: 10px"><%=rs("apellidos")%></td>
				<td style="font-family: verdana; font-size: 10px" style="display: none"><%=rs("pais")%></td>
				<td style="font-family: verdana; font-size: 10px"><%=rs("nombrepais")%></td>
				<td style="font-family: verdana; font-size: 10px" style="display: none"><%=rs("ciudad")%></td>
				<td style="font-family: verdana; font-size: 10px"><%=rs("nombreciudad")%></td>
			</tr>
			<%i = i + 1%>
			<%rs.movenext%>
		<%loop%>
	</table>
	
</BODY>
</HTML>
<%set rs = nothing%>