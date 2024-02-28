<%@ Language=VBScript %>
<%
	Response.Redirect "http://laurel:87/loginmanual.aspx?CodigoUsuario=" & trim(Session("NombreUsuarioOperador")) & "&CodigoSucursal=" & trim(Session("CodigoAgente")) & "&CodigoCliente=" & "trim(frmCliente.txtExpress.value)" & "&ClienteAgente=" & Session("CodigoCliente") & "&CategoriaAgente=" & Session("Categoria") ', "AFEX", "height=900, width=1000""
%>
