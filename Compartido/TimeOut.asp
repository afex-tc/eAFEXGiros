<!-- TimeOut.asp -->
<%
	'Se asegura que la página no se almacene en la memoria cache
	'Response.expires = 0
	'Response.expiresabsolute = Now() - 1
	'Response.addHeader "pragma", "no-cache"
	'Response.addHeader "cache-control", "private"
	'Response.CacheControl = "no-cache"

	If Session("CodigoCliente") = "" Then
		response.Redirect "http:TimeOut.htm"
		response.end
	End If
%>
